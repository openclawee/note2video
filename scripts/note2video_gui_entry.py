from __future__ import annotations

import sys


def main() -> int:
    from note2video.gui.app import main as gui_main

    return int(gui_main(sys.argv[1:]) or 0)


if __name__ == "__main__":
    raise SystemExit(main())

