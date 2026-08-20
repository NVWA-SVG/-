from __future__ import annotations

import argparse
import json
from dataclasses import asdict
from pathlib import Path

from renamer import (
    RenameError,
    build_plan,
    execute_plan,
    latest_history,
    plan_summary,
    undo_history,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="PowerShell GUI JSON bridge")
    parser.add_argument("action", choices=("preview", "execute", "undo"))
    parser.add_argument("--output", required=True)
    parser.add_argument("--folder", required=True)
    parser.add_argument("--roster", default="")
    parser.add_argument("--project", default="")
    parser.add_argument("--template", default="{姓名}_{学号}_{项目}")
    args = parser.parse_args()

    try:
        if args.action == "undo":
            history = latest_history(args.folder)
            count = undo_history(history)
            result = {"ok": True, "count": count}
        else:
            items = build_plan(args.folder, args.roster, args.project, args.template)
            summary = plan_summary(items)
            result = {
                "ok": True,
                "items": [asdict(item) for item in items],
                "summary": summary,
            }
            if args.action == "execute":
                history = execute_plan(items, args.folder)
                result["history"] = str(history)
    except RenameError as exc:
        result = {"ok": False, "error": str(exc)}
    except Exception as exc:  # GUI 边界：确保任何错误都能显示，不闪退。
        result = {"ok": False, "error": f"未预期错误：{exc}"}

    Path(args.output).write_text(
        json.dumps(result, ensure_ascii=False), encoding="utf-8-sig"
    )
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
