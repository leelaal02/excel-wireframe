from pathlib import Path

import re

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "skills" / "excel-wireframe"


def test_skill_md_has_frontmatter():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    assert text.startswith("---\n")
    fm = text.split("---", 2)[1]
    assert re.search(r"^name:\s*excel-wireframe\s*$", fm, re.M)
    assert re.search(r"^description:\s*\S", fm, re.M)


def test_skill_md_documents_three_scripts():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    for script in ("analyze.py", "extract.py", "build.py"):
        assert script in text


def test_skill_md_covers_missing_template():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    assert "기본 템플릿" in text
    assert "--template" in text


def test_reference_exists():
    ref = SKILL / "references" / "mapping-schema.md"
    text = ref.read_text(encoding="utf-8")
    assert "mapping.json" in text
    assert "screens.json" in text
    assert "detail_tables" in text
    assert "상세표1" in text
