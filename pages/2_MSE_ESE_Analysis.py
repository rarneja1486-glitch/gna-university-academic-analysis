from pathlib import Path

exec(
    Path(__file__).resolve()
    .parents[1]
    .joinpath("marks.py")
    .read_text(encoding="utf-8"),
    globals()
)
