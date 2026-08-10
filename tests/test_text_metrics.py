from text_metrics import EMU_PER_PT, plan_row_heights, row_height, text_lines

# 기본 서식의 설명 칸 폭(EMU). content_area 폭 9957099를 표 5개로 나누고
# 번호 칸과 좌우 여백을 뺀 값이다. 5.03cm 남짓이고 7pt 전각이 한 줄에 20자,
# 6pt면 23자 들어간다.
DESC_W = 1810509


def test_text_lines_counts_one_line_for_short_text():
    assert text_lines("짧은 글", DESC_W, 7.0) == 1


def test_text_lines_wraps_by_width():
    # 전각 한 글자는 폰트 크기와 폭이 같다. 7pt면 한 줄에 20자가 들어간다.
    assert text_lines("가" * 40, DESC_W, 7.0) == 2
    assert text_lines("가" * 41, DESC_W, 7.0) == 3


def test_text_lines_treats_halfwidth_as_half():
    """반각은 전각의 절반 폭으로 센다. 같은 글자 수라도 줄 수가 적어야 한다."""
    full = text_lines("가" * 60, DESC_W, 7.0)
    half = text_lines("a" * 60, DESC_W, 7.0)
    assert half < full


def test_text_lines_empty_is_one_line():
    """빈 셀도 한 줄 높이는 차지한다. 0을 돌려주면 행 높이가 0이 된다."""
    assert text_lines("", DESC_W, 7.0) == 1
    assert text_lines(None, DESC_W, 7.0) == 1


def test_text_lines_single_char_wider_than_box():
    """폭보다 넓은 한 글자도 한 줄이다. 무한 루프나 0이 되면 안 된다."""
    assert text_lines("가", 1000, 7.0) == 1


def test_text_lines_counts_explicit_newlines():
    """set_cell_text가 \\n을 문단으로 나누므로 줄 수에 반영해야 한다."""
    assert text_lines("첫 줄\n둘째 줄", DESC_W, 7.0) == 2


def test_text_lines_shrinks_with_smaller_font():
    # 20자/줄과 23자/줄은 대부분의 길이에서 같은 줄 수를 낸다(100자면 둘 다 5줄).
    # 차이가 확실히 나는 길이를 골라야 검증이 된다.
    long = "가" * 69
    assert text_lines(long, DESC_W, 7.0) == 4
    assert text_lines(long, DESC_W, 6.0) == 3


def test_row_height_grows_with_lines():
    one = row_height(1, 7.0)
    two = row_height(2, 7.0)
    assert two - one == int(7.0 * 1.2 * EMU_PER_PT)


def test_row_height_includes_margins():
    assert row_height(1, 7.0, margin_top=9525, margin_bottom=0) == \
        int(7.0 * 1.2 * EMU_PER_PT) + 9525


def test_plan_row_heights_returns_one_height_per_row():
    """표 다섯 개가 나란히 놓이므로 행 높이는 표를 가로질러 하나로 통일된다."""
    pages = [["짧게"] * 20]
    heights = plan_row_heights(pages, 4, DESC_W, 7.0, [100, 200, 300, 400])
    assert len(heights) == 4


def test_plan_row_heights_respects_floors():
    """계산값이 실측 하한보다 작으면 하한이 이긴다."""
    floors = [382457, 268746, 496168, 268746]
    heights = plan_row_heights([["짧게"] * 20], 4, DESC_W, 7.0, floors)
    assert heights == floors


def test_plan_row_heights_grows_for_long_text():
    floors = [382457, 268746, 496168, 268746]
    texts = ["짧게"] * 20
    texts[1] = "가" * 300           # 슬롯 1 -> 표 0, 행 1
    heights = plan_row_heights([texts], 4, DESC_W, 7.0, floors)
    assert heights[1] > floors[1]
    assert heights[0] == floors[0]  # 다른 행은 그대로


def test_plan_row_heights_takes_max_across_tables():
    """같은 행 인덱스면 표가 달라도 가장 큰 것에 맞춘다."""
    floors = [100, 100, 100, 100]
    texts = ["짧게"] * 20
    texts[13] = "가" * 300          # 슬롯 13 -> 표 3, 행 1
    heights = plan_row_heights([texts], 4, DESC_W, 7.0, floors)
    assert heights[1] > heights[0]


def test_plan_row_heights_takes_max_across_pages():
    """한 화면의 모든 장이 같은 표 높이를 쓴다. 최댓값이 이겨야 한다."""
    floors = [100, 100, 100, 100]
    short = ["짧게"] * 20
    long = list(short)
    long[0] = "가" * 300
    one = plan_row_heights([short], 4, DESC_W, 7.0, floors)
    both = plan_row_heights([short, long], 4, DESC_W, 7.0, floors)
    assert both[0] > one[0]


def test_plan_row_heights_handles_partial_page():
    """마지막 장은 슬롯을 다 채우지 못한다. 빈 슬롯이 있어도 터지지 않는다."""
    floors = [100, 100, 100, 100]
    heights = plan_row_heights([["하나", "둘"]], 4, DESC_W, 7.0, floors)
    assert len(heights) == 4
    assert all(h >= 100 for h in heights)


def test_plan_row_heights_handles_empty_page():
    floors = [100, 100, 100, 100]
    assert plan_row_heights([[]], 4, DESC_W, 7.0, floors) == floors


def test_plan_row_heights_shrinks_with_smaller_font():
    floors = [100, 100, 100, 100]
    texts = ["가" * 200] * 20
    big = plan_row_heights([texts], 4, DESC_W, 7.0, floors)
    small = plan_row_heights([texts], 4, DESC_W, 6.0, floors)
    assert sum(small) < sum(big)


def test_fits_lines_counts_what_a_row_holds():
    """실측 행 높이가 7pt를 몇 줄 담는지. 이 수치가 넘침 판정의 기준이다."""
    from text_metrics import fits_lines
    assert [fits_lines(h, 7.0) for h in (382457, 268746, 496168, 268746)] == \
        [3, 2, 4, 2]


def test_fits_lines_grows_with_smaller_font():
    from text_metrics import fits_lines
    assert fits_lines(382457, 6.0) > fits_lines(382457, 7.0)


def test_fits_lines_never_returns_zero():
    """행이 아무리 낮아도 한 줄은 있는 것으로 친다. 0을 돌려주면 전부 넘침이 된다."""
    from text_metrics import fits_lines
    assert fits_lines(1, 7.0) == 1
    assert fits_lines(0, 7.0) == 1
