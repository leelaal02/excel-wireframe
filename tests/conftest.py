import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "skills" / "excel-to-wireframe-ppt" / "scripts"
sys.path.insert(0, str(SCRIPTS))
