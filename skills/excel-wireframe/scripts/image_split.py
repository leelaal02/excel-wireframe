# -*- coding: utf-8 -*-
"""긴 화면 이미지를 PPT 이미지 자리에 맞춰 세로로 나눈다.

세 가지를 지킨다.

1. **의미가 끊기지 않게 자른다.** 일정 높이로 기계적으로 자르지 않고, 가로로
   이어진 여백 띠(섹션 경계)를 찾아 거기서 자른다. 목표 높이는 상한이
   아니라 참고값이다 — 섹션이 끊기느니 조각이 길어져 조금 축소되는 편이 낫다.
   섹션 헤더 아래 틈에서는 자르지 않는다. 거기서 자르면 제목만 앞 장에 남고
   내용이 다음 장으로 넘어간다.
2. **상세를 구간에 맞춘다.** 스크린샷에는 상세 번호에 대응하는 SoM 뱃지(노란 원)가
   찍혀 있다. 각 조각에 뱃지가 몇 개 들어가는지 세어 그 수만큼 상세를 배분한다.
   뱃지 안의 숫자를 읽지 않는 이유는 순서 추정이 불안정하기 때문이다 — 같은 줄에
   있는 뱃지들의 y가 몇 px씩 어긋나 행 묶음 임계값 하나로 가를 수 없다. 개수만
   세면 행 안의 순서가 틀려도 구간 경계만 맞으면 결과가 정확하다.
3. **넘치면 다음 장으로 잇는다.** 한 조각의 상세가 슬롯 수를 넘으면 같은 조각을
   여러 장에 반복해 싣는다. 그 상세가 가리키는 화면이 그 조각 안에 있기 때문이다.
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
from PIL import Image

# --- 뱃지 검출 (노란 원) ---
BADGE_MIN_PX = 60          # 이보다 작은 덩어리는 잡티
BADGE_MIN_SIDE = 10
BADGE_MAX_SIDE = 60
BADGE_ASPECT = (0.6, 1.6)  # 원이므로 가로세로가 비슷

# --- 여백 띠 검출 ---
QUIET_STD = 12             # 행의 색 분산이 이보다 낮으면 조용한 행
QUIET_MEAN = 200           # 그리고 밝아야 한다 (어두운 단색 띠는 여백이 아니다)
QUIET_MIN_RUN = 5          # 이만큼 연속돼야 가로 여백 띠로 본다. 5픽셀
NAV_SKIP_RATIO = 0.12      # 좌측 네비게이션은 세로로 계속 이어져 여백 판단을 방해한다

# --- 분할 계획 ---
CUT_MIN_RATIO = 0.5        # 목표 높이의 이 배수부터 자를 곳을 찾는다
CUT_MAX_RATIO = 1.5        # 이 배수를 넘으면 여백이 없어도 자른다
HEADER_MAX_PX = 40         # 띠와 띠 사이가 이보다 얇으면 섹션 헤더 한 줄로 본다. 헤더 고아 방지


# detect_badges() 함수: 화면에 있는 SoM 노란 번호를 모두 찾는 것
def detect_badges(img: Image.Image) -> list[tuple[int, int]]:
    """SoM 뱃지(노란 원)의 중심 좌표를 (y, x) 목록으로 돌려준다."""
    a = np.asarray(img.convert("RGB")).astype(np.int16)  # (1) 이미지를 RGB 배열로 변환
    r, g, b = a[:, :, 0], a[:, :, 1], a[:, :, 2]
    mask = (r > 180) & (g > 140) & (b < 110) & (r - b > 90) & (g - b > 60)  # (2) 노란색만 추출

    h, w = mask.shape # mask의 크기를 가져오기
    seen = np.zeros_like(mask, dtype=bool) # 이미 방문한 픽셀인지 기록하는 배열을 생성
    out: list[tuple[int, int]] = [] # 최종적으로 검출한 뱃지 중심 좌표를 저장할 리스트
    ys, xs = np.nonzero(mask) # mask에서 True인 픽셀의 좌표를 모두 찾기
    for i in np.argsort(ys): # y좌표(위에서 아래 순서)를 기준으로 픽셀 반복
        y0, x0 = int(ys[i]), int(xs[i]) # 현재 탐색을 시작할 픽셀의 좌표를 가져
        if seen[y0, x0]:
            continue
        stack = [(y0, x0)]   # (3) 연결된 픽셀 찾기
        seen[y0, x0] = True  # 시작 픽셀을 방문 처리
        pix = [] # 현재 덩어리에 속하는 픽셀들을 저장할 리스트
        while stack: # 스택이 빌 때까지 DFS를 수행
            y, x = stack.pop() # 스택에서 픽셀 하나를 꺼내기
            pix.append((y, x)) # 현재 픽셀을 이 덩어리에 추가
            for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)): # 현재 픽셀의 상, 하, 좌, 우 네 방향을 확인
                ny, nx = y + dy, x + dx # 이웃 픽셀의 좌표를 계산
                if 0 <= ny < h and 0 <= nx < w and mask[ny, nx] and not seen[ny, nx]: 
                # 이미지 범위를 벗어나지 않았는지 확인, 이웃 픽셀이 노란색(True)인지 확인
                    seen[ny, nx] = True # 방문 처리
                    stack.append((ny, nx)) # 이어진 픽셀이므로 스택에 넣어 나중에 계속 탐색
        if len(pix) < BADGE_MIN_PX:  # (4) 너무 작은 것은 제거
            continue
        p = np.array(pix) # 리스트를 NumPy 배열로 변환하여 계산
        y1, y2 = int(p[:, 0].min()), int(p[:, 0].max())
        x1, x2 = int(p[:, 1].min()), int(p[:, 1].max())
        bh, bw = y2 - y1 + 1, x2 - x1 + 1
        if not (BADGE_MIN_SIDE <= bh <= BADGE_MAX_SIDE): # 높이가 너무 작거나 크면 제외
            continue
        if not (BADGE_MIN_SIDE <= bw <= BADGE_MAX_SIDE): # 너비가 너무 작거나 크면 제외
            continue
        if not (BADGE_ASPECT[0] <= bw / bh <= BADGE_ASPECT[1]): # (5) 원형에 가까운지 확인
            continue
        out.append(((y1 + y2) // 2, (x1 + x2) // 2))  # (6) 중심 좌표를 계산하여 out 리스트에 저장
    out.sort() # 중심 좌표를 y축 기준(위에서 아래 순서)으로 정렬
    return out # 최종적으로 검출된 모든 SoM 뱃지의 중심 좌표를 반환



def find_quiet_bands(img: Image.Image) -> list[tuple[int, int]]:
    """가로로 이어진 여백 띠를 (시작y, 끝y) 목록으로 돌려준다.

    좌측 네비게이션은 세로로 끊김 없이 이어져 어느 행에서도 '조용하지 않게'
    만든다. 본문만 보도록 왼쪽 일부를 잘라내고 판단한다.
    """
    a = np.asarray(img.convert("RGB")).astype(np.int16)
    skip = int(a.shape[1] * NAV_SKIP_RATIO)
    body = a[:, skip:, :]
    quiet = (body.std(axis=(1, 2)) < QUIET_STD) & (body.mean(axis=(1, 2)) > QUIET_MEAN)

    bands: list[tuple[int, int]] = []
    start = None
    for y, q in enumerate(quiet):
        if q and start is None:
            start = y
        elif not q and start is not None:
            if y - start >= QUIET_MIN_RUN:
                bands.append((start, y - 1))
            start = None
    if start is not None and len(quiet) - start >= QUIET_MIN_RUN:
        bands.append((start, len(quiet) - 1))
    return bands



def leads_section(band, bands) -> bool:
    """이 띠 바로 위에 섹션 헤더 한 줄만 얹혀 있는지 본다.

    화면의 섹션 헤더는 위아래로 여백을 거느린다. 아래쪽 여백에서 자르면 헤더
    글자만 앞 조각 끝에 남고 그 헤더가 이끄는 표와 입력란은 다음 조각으로
    넘어간다 — 읽는 사람에게는 제목과 내용이 갈라져 보인다. 위 띠와의 간격이
    한 줄 높이밖에 안 되면 그 사이는 헤더이므로, 이 띠는 자를 곳이 아니다.
    """
    start = band[0]
    above = [end for _s, end in bands if end < start]
    if not above:
        return False
    return start - max(above) - 1 <= HEADER_MAX_PX


def plan_cuts(height: int, target_h: int, bands, badges) -> list[tuple[int, int, int]]:
# height   : 원본 이미지의 높이(px)
# target_h : PPT 이미지 영역에 맞춘 목표 분할 높이(px)
# bands    : 검출된 여백의 위치 목록 [(시작 y, 끝 y) ...]
# badges   : 검출된 SoM 뱃지의 중심 좌표 목록 [(y, x) ...]
    """이미지를 (시작y, 끝y, 뱃지수) 조각으로 나눈다.

    자를 곳은 목표 높이의 CUT_MIN_RATIO~CUT_MAX_RATIO 범위에서 고른다. 그 안에
    여백 띠가 있으면 섹션 헤더를 고아로 남기지 않는 것들 중 목표 높이에 가장
    가까운 것을 쓰고, 없으면 뱃지를 관통하지 않는 지점을, 그것도 없으면 범위
    끝에서 자른다.
    """
    if target_h <= 0:
        return [(0, height - 1, len(badges))] # 분할하지 않고 이미지 전체를 하나의 조각으로 반환

    badge_ys = [y for y, _ in badges] # SoM 뱃지의 y좌표만 따로 추출,  # [결과값 for 변수 in 반복대상]
    cuts: list[tuple[int, int, int]] = [] # 최종 분할 결과를 저장하는 리스트
    top = 0 # 현재 조각의 시작 위치(y좌표)
    while top < height:
        if height - top <= target_h * CUT_MAX_RATIO:
            bottom = height - 1
        else:
            lo = top + int(target_h * CUT_MIN_RATIO)
            hi = min(height - 1, top + int(target_h * CUT_MAX_RATIO))
            inside = [(s, e) for s, e in bands if lo <= s and e <= hi]
            # 헤더를 고아로 만드는 띠는 뺀다. 다만 그것뿐이면 안 자를 수는
            # 없으니 그대로 쓴다 — 조각이 무한정 길어지는 편이 더 나쁘다.
            usable = [b for b in inside if not leads_section(b, bands)] or inside
            if usable:
                # 두께가 아니라 목표 높이와의 거리로 고른다. 띠 두께는 섹션이
                # 얼마나 크게 갈리는지와 상관이 없고(5px과 6px이 갈린다),
                # 두꺼운 쪽을 우선하면 조각이 목표보다 훨씬 짧아져 이미지만
                # 납작해진 채 슬라이드의 자리가 남는다.
                s, e = min(usable,
                           key=lambda b: abs((b[0] + b[1]) // 2 - (top + target_h)))
                bottom = (s + e) // 2
            else:
                free = [y for y in range(hi, lo - 1, -1)
                        if all(abs(y - by) > BADGE_MAX_SIDE // 2 for by in badge_ys)]
                bottom = free[0] if free else hi
        n = sum(1 for y in badge_ys if top <= y <= bottom)
        cuts.append((top, bottom, n))
        top = bottom + 1
    return cuts


def slice_image(path: Path, cuts, out_dir: Path, stem: str) -> list[Path]:
    """계획대로 이미지를 잘라 저장하고 조각 경로를 돌려준다."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    made = []
    with Image.open(path) as im:
        for i, (top, bottom, _n) in enumerate(cuts):
            piece = im.crop((0, top, im.width, bottom + 1))
            p = out_dir / ("%s_part%d.png" % (stem, i + 1))
            piece.save(p)
            made.append(p)
    return made


def split_for_box(path: Path, box_w: int, box_h: int, out_dir: Path):
    """이미지가 자리보다 길면 나눈다.

    (조각 경로 목록, 조각별 뱃지 수, 뱃지 총수)를 돌려준다. 나눌 필요가 없으면
    ([원본], [], 0)이다 — 뱃지 수가 비어 있으면 호출자는 기존 방식대로 상세를
    순차 배분한다.
    """
    path = Path(path)
    with Image.open(path) as im:
        iw, ih = im.size
        target_h = int(iw * box_h / box_w) if box_w else ih
        if ih <= target_h:
            return [path], [], 0
        img = im.convert("RGB")
        badges = detect_badges(img)
        bands = find_quiet_bands(img)
    cuts = plan_cuts(ih, target_h, bands, badges)
    if len(cuts) <= 1:
        return [path], [], len(badges)
    pieces = slice_image(path, cuts, out_dir, path.stem)
    return pieces, [n for _, _, n in cuts], len(badges)
