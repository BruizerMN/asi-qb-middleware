"""Called by the pre-commit hook to increment the BUILD number in version.py."""
import re
import pathlib

f = pathlib.Path(__file__).parent.parent / "version.py"
content = f.read_text()
match = re.search(r'BUILD = "(\d+)"', content)
if not match:
    raise SystemExit("Could not find BUILD in version.py")

new_build = f"{int(match.group(1)) + 1:04d}"
content = re.sub(r'BUILD = "\d+"', f'BUILD = "{new_build}"', content)
f.write_text(content)
print(f"Build → {new_build}")
