from __future__ import annotations

from pathlib import Path
import re
import sys


EM_DASH = "—"
EMOJI_PATTERN = re.compile(
    "["
    "\U0001F300-\U0001F5FF"
    "\U0001F600-\U0001F64F"
    "\U0001F680-\U0001F6FF"
    "\U0001F700-\U0001F77F"
    "\U0001F780-\U0001F7FF"
    "\U0001F800-\U0001F8FF"
    "\U0001F900-\U0001F9FF"
    "\U0001FA00-\U0001FAFF"
    "\U00002700-\U000027BF"
    "]"
)

ALLOWED_EXTENSIONS = {
    ".md",
    ".txt",
    ".py",
    ".toml",
    ".yml",
    ".yaml",
    ".json",
}


def _line_positions(text: str, needle: str) -> list[int]:
    positions: list[int] = []
    for i, line in enumerate(text.splitlines(), start=1):
        if needle in line:
            positions.append(i)
    return positions


def _emoji_line_positions(text: str) -> list[int]:
    positions: list[int] = []
    for i, line in enumerate(text.splitlines(), start=1):
        if EMOJI_PATTERN.search(line):
            positions.append(i)
    return positions


def _should_check(path: Path) -> bool:
    if not path.exists() or path.is_dir():
        return False
    return path.suffix.lower() in ALLOWED_EXTENSIONS


def main(argv: list[str]) -> int:
    failed = False

    for raw_path in argv[1:]:
        path = Path(raw_path)
        if not _should_check(path):
            continue

        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue

        em_dash_lines = _line_positions(text, EM_DASH)
        emoji_lines = _emoji_line_positions(text)

        for line in em_dash_lines:
            print(f"{path}:{line}: avoid em dash character '{EM_DASH}'")
            failed = True

        for line in emoji_lines:
            print(f"{path}:{line}: avoid emojis in repository text")
            failed = True

    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
