import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "skills" / "excel-wireframe" / "scripts"
sys.path.insert(0, str(SCRIPTS))
