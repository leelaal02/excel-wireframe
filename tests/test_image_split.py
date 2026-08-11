"""긴 화면을 어디서 자르는지 검증한다.

실제 화면(CPTIPMT1002)에서 섹션 헤더만 앞 장에 남고 본문이 다음 장으로
넘어간 일이 있었다. 그 구조를 합성 이미지로 재현해 둔다 — 섹션 사이 여백은
5px인데 헤더 아래 틈만 6px이어서, 하필 그 틈이 유일한 절단 후보가 됐다.
"""
import numpy as np
import pytest
from PIL import Image

from image_split import find_quiet_bands, plan_cuts

WIDTH = 400  # 좌측 12%를 건너뛰어도 콘텐츠가 남는 폭


def make_image(spec):
    """('q'|'c', 높이) 목록으로 세로로 쌓은 이미지를 만든다.

    q는 흰 여백 행, c는 좌우가 검고 흰 콘텐츠 행이다. 콘텐츠 행은 평균이
    낮아 여백 판정을 통과하지 못한다.
    """
    rows = []
    for kind, height in spec:
        band = np.full((height, WIDTH, 3), 255, dtype=np.uint8)
        if kind == "c":
            band[:, ::2, :] = 0
        rows.append(band)
    return Image.fromarray(np.vstack(rows))


def owner(cuts, y):
    """y가 들어간 조각의 번호."""
    return next(i for i, (top, bottom, _n) in enumerate(cuts) if top <= y <= bottom)


def test_finds_five_pixel_gap():
    """섹션 사이 여백이 5px뿐인 화면이 있다. 그것도 경계로 봐야 한다."""
    img = make_image([("c", 100), ("q", 5), ("c", 100)])
    assert find_quiet_bands(img) == [(100, 104)]


def test_cut_keeps_section_header_with_its_body():
    """헤더 아래 틈이 더 두꺼워도 거기서 자르면 안 된다.

    헤더 줄만 앞 조각에 남고 그 헤더가 이끄는 본문은 다음 조각으로 넘어간다.
    """
    img = make_image([
        ("c", 1000),   # 앞 섹션
        ("q", 5),      # 헤더 위 여백
        ("c", 10),     # 섹션 헤더 한 줄
        ("q", 6),      # 헤더 아래 여백 — 여기서 자르면 헤더가 고아가 된다
        ("c", 280),    # 헤더가 이끄는 본문
        ("q", 5),      # 본문이 끝난 자리
        ("c", 700),
    ])
    header_y, body_y = 1010, 1100
    cuts = plan_cuts(img.height, 1300, find_quiet_bands(img), [])
    assert owner(cuts, header_y) == owner(cuts, body_y)


def test_prefers_band_nearest_target_height():
    """후보가 여럿이면 목표 높이에 가까운 것을 고른다.

    두꺼운 띠를 우선하면 조각이 목표보다 훨씬 짧아져, 슬라이드의 이미지
    자리를 남기고 그림만 납작해진다.
    """
    cuts = plan_cuts(2000, 900, [(500, 514), (890, 894)], [])
    assert cuts[0][1] == 892
