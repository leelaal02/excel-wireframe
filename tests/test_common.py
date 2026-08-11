from pathlib import Path

from common import (
    Warnings,
    migrate_legacy_work,
    migration_notice,
    read_json,
    resolve_output_path,
    work_dir,
    write_json,
)


def test_json_roundtrip_keeps_korean(tmp_path: Path):
    p = tmp_path / "a.json"
    write_json(p, {"name": "이용기관 목록"})
    assert read_json(p) == {"name": "이용기관 목록"}
    assert "이용기관" in p.read_text(encoding="utf-8")


def test_warnings_collects_and_formats():
    w = Warnings()
    w.add("B2BISMT1001", "no-image", "이미지를 찾지 못했습니다")
    w.add(None, "shape-not-found", "제목 도형 없음")
    assert len(w) == 2
    assert w.to_list()[0] == {
        "screen_id": "B2BISMT1001",
        "code": "no-image",
        "message": "이미지를 찾지 못했습니다",
    }
    text = w.format()
    assert "B2BISMT1001" in text
    assert "no-image" in text


def test_warnings_format_is_empty_when_none():
    assert Warnings().format() == ""


def test_output_path_uses_excel_stem(tmp_path: Path):
    assert resolve_output_path(tmp_path, "짧은 버전.xlsx") == tmp_path / "짧은 버전.pptx"


def test_output_path_takes_stem_from_full_path(tmp_path: Path):
    # meta.source에는 추출할 때 쓴 경로가 통째로 들어간다. 파일명만 써야 한다.
    src = str(Path("C:/설계/입력") / "KB차세대_화면설계.xlsx")
    assert resolve_output_path(tmp_path, src) == tmp_path / "KB차세대_화면설계.pptx"


def test_output_path_numbers_when_name_taken(tmp_path: Path):
    (tmp_path / "짧은 버전.pptx").write_bytes(b"")
    assert resolve_output_path(tmp_path, "짧은 버전.xlsx") == tmp_path / "짧은 버전2.pptx"

    (tmp_path / "짧은 버전2.pptx").write_bytes(b"")
    assert resolve_output_path(tmp_path, "짧은 버전.xlsx") == tmp_path / "짧은 버전3.pptx"


def test_output_path_falls_back_when_source_missing(tmp_path: Path):
    # 손으로 만든 screens.json이나 meta.source가 없던 시절의 파일도 이름을 받아야 한다.
    assert resolve_output_path(tmp_path, "") == tmp_path / "화면설계서.pptx"


def test_output_path_does_not_create_directory(tmp_path: Path):
    missing = tmp_path / "output"
    assert resolve_output_path(missing, "a.xlsx") == missing / "a.pptx"
    assert not missing.exists()


def test_work_dir_is_dot_work_under_output():
    assert work_dir(Path("output")) == Path("output") / ".work"


def test_work_dir_does_not_create_directory(tmp_path: Path):
    # 경로를 정하는 일과 만드는 일을 섞지 않는다. mkdir은 쓰는 쪽이 한다.
    assert work_dir(tmp_path) == tmp_path / ".work"
    assert not (tmp_path / ".work").exists()


def _make_legacy_output(root: Path) -> None:
    root.mkdir(parents=True, exist_ok=True)
    for name in ("structure-report.json", "mapping.json", "screens.json",
                 "default-template.pptx"):
        (root / name).write_bytes(b"legacy")
    (root / "images").mkdir()
    (root / "images" / "img1.png").write_bytes(b"png")


def test_migrate_moves_legacy_files_into_work(tmp_path: Path):
    _make_legacy_output(tmp_path)

    moved = migrate_legacy_work(tmp_path)

    assert sorted(moved) == sorted([
        "structure-report.json", "mapping.json", "screens.json",
        "default-template.pptx", "images",
    ])
    for name in ("structure-report.json", "mapping.json", "screens.json",
                 "default-template.pptx"):
        assert (tmp_path / ".work" / name).read_bytes() == b"legacy"
        assert not (tmp_path / name).exists()
    assert (tmp_path / ".work" / "images" / "img1.png").read_bytes() == b"png"
    assert not (tmp_path / "images").exists()


def test_migrate_leaves_result_pptx_in_output(tmp_path: Path):
    # 결과물은 옮기면 안 된다. default-template.pptx와 확장자가 같으므로
    # 화이트리스트에 없는 pptx는 손대지 않는다는 규칙이 유일한 방어선이다.
    _make_legacy_output(tmp_path)
    (tmp_path / "짧은 버전.pptx").write_bytes(b"result")
    (tmp_path / "짧은 버전2.pptx").write_bytes(b"result2")

    migrate_legacy_work(tmp_path)

    assert (tmp_path / "짧은 버전.pptx").read_bytes() == b"result"
    assert (tmp_path / "짧은 버전2.pptx").read_bytes() == b"result2"
    assert not (tmp_path / ".work" / "짧은 버전.pptx").exists()


def test_migrate_keeps_existing_work_file(tmp_path: Path):
    # .work/ 쪽이 새 경로에서 만들어진 최신 파일이다. 구버전이 덮어쓰면 안 된다.
    _make_legacy_output(tmp_path)
    (tmp_path / ".work").mkdir()
    (tmp_path / ".work" / "mapping.json").write_bytes(b"current")

    moved = migrate_legacy_work(tmp_path)

    assert "mapping.json" not in moved
    assert (tmp_path / ".work" / "mapping.json").read_bytes() == b"current"
    assert (tmp_path / "mapping.json").read_bytes() == b"legacy"


def test_migrate_does_nothing_when_output_is_clean(tmp_path: Path):
    (tmp_path / "짧은 버전.pptx").write_bytes(b"result")

    assert migrate_legacy_work(tmp_path) == []
    assert not (tmp_path / ".work").exists()


def test_migrate_keeps_explicitly_given_paths(tmp_path: Path):
    """CLI가 --mapping으로 가리킨 파일을 옮기면 그 실행이 파일을 잃는다."""
    _make_legacy_output(tmp_path)

    moved = migrate_legacy_work(tmp_path, keep=[tmp_path / "mapping.json"])

    assert "mapping.json" not in moved
    assert (tmp_path / "mapping.json").read_bytes() == b"legacy"
    assert "screens.json" in moved


def test_migrate_keep_ignores_paths_outside_output(tmp_path: Path):
    # 다른 폴더의 mapping.json을 지정한 경우다. 같은 파일명이라고 해서
    # output/ 안의 것을 남길 이유가 없다.
    _make_legacy_output(tmp_path)
    other = tmp_path / "다른곳"
    other.mkdir()
    (other / "mapping.json").write_bytes(b"other")

    moved = migrate_legacy_work(tmp_path, keep=[other / "mapping.json"])

    assert "mapping.json" in moved


def test_migration_notice_names_what_moved():
    text = migration_notice(["mapping.json", "images"])
    assert "2개" in text
    assert "mapping.json" in text
    assert "images" in text
    assert ".work" in text


def test_migration_notice_is_empty_when_nothing_moved():
    # Warnings.format과 같은 규약이다 — 빈 문자열이면 호출부가 아무것도 찍지 않는다.
    assert migration_notice([]) == ""


def test_migrate_does_nothing_when_output_missing(tmp_path: Path):
    missing = tmp_path / "output"

    assert migrate_legacy_work(missing) == []
    assert not missing.exists()
