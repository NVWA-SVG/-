import csv
import tempfile
import unittest
from pathlib import Path

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from renamer import RenameError, build_plan, execute_plan, load_roster, undo_history


class RenamerTests(unittest.TestCase):
    def make_roster(self, folder: Path, rows: list[tuple[str, str, str]]) -> Path:
        roster = folder / "roster.csv"
        with roster.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(["姓名", "学号", "项目"])
            writer.writerows(rows)
        return roster

    def test_preview_execute_and_undo(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            roster = self.make_roster(folder, [("张三", "001", "AI大赛"), ("李四", "002", "AI大赛")])
            (folder / "张三.docx").write_bytes(b"one")
            (folder / "李四.pdf").write_bytes(b"two")
            (folder / "其他.txt").write_bytes(b"three")

            plan = build_plan(folder, roster)
            self.assertEqual(sum(item.status == "ready" for item in plan), 2)
            self.assertEqual(sum(item.status == "skipped" for item in plan), 2)  # 其他 + 花名册

            history = execute_plan(plan, folder)
            self.assertTrue((folder / "张三_001_AI大赛.docx").is_file())
            self.assertTrue((folder / "李四_002_AI大赛.pdf").is_file())
            self.assertEqual(undo_history(history), 2)
            self.assertTrue((folder / "张三.docx").is_file())
            self.assertTrue((folder / "李四.pdf").is_file())

    def test_project_override(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            roster = self.make_roster(folder, [("张三", "001", "旧项目")])
            (folder / "张三.txt").write_text("x", encoding="utf-8")
            plan = build_plan(folder, roster, project="新项目")
            ready = next(item for item in plan if item.status == "ready")
            self.assertEqual(Path(ready.target).name, "张三_001_新项目.txt")

    def test_duplicate_names_are_rejected(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            roster = self.make_roster(folder, [("张三", "001", "A"), ("张三", "002", "B")])
            with self.assertRaises(RenameError):
                load_roster(roster)

    def test_existing_target_is_conflict(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            roster = self.make_roster(folder, [("张三", "001", "AI")])
            (folder / "张三.txt").write_text("source", encoding="utf-8")
            (folder / "张三_001_AI.txt").write_text("existing", encoding="utf-8")
            plan = build_plan(folder, roster)
            match = next(item for item in plan if Path(item.source).name == "张三.txt")
            self.assertEqual(match.status, "conflict")

    def test_name_embedded_in_mixed_filename(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            roster = self.make_roster(folder, [("梁书前", "202434610215", "图像处理")])
            source = folder / "202434610215梁书前24AI2实验二_.docx"
            source.write_text("x", encoding="utf-8")
            plan = build_plan(folder, roster)
            match = next(item for item in plan if Path(item.source) == source)
            self.assertEqual(match.status, "ready")
            self.assertEqual(Path(match.target).name, "梁书前_202434610215_图像处理.docx")
            self.assertEqual(match.reason, "按文件名中的学号识别")

    def test_name_only_embedded_in_mixed_filename(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            roster = self.make_roster(folder, [("刘鸣诚", "202434610099", "图像处理")])
            source = folder / "图像实验2_24AI2_刘鸣诚.docx"
            source.write_text("x", encoding="utf-8")
            plan = build_plan(folder, roster)
            match = next(item for item in plan if Path(item.source) == source)
            self.assertEqual(match.status, "ready")
            self.assertEqual(match.matched_name, "刘鸣诚")
            self.assertEqual(match.reason, "按文件名中的姓名识别")

    def test_conflicting_student_id_and_name_is_blocked(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            roster = self.make_roster(
                folder,
                [("张三", "001", "AI"), ("李四", "002", "AI")],
            )
            source = folder / "001_李四_作业.docx"
            source.write_text("x", encoding="utf-8")
            plan = build_plan(folder, roster)
            match = next(item for item in plan if Path(item.source) == source)
            self.assertEqual(match.status, "conflict")
            self.assertIn("学号指向", match.reason)

    def test_longer_name_wins_when_names_overlap(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            roster = self.make_roster(
                folder,
                [("王明", "001", "AI"), ("王小明", "002", "AI")],
            )
            source = folder / "24AI2_王小明_实验.docx"
            source.write_text("x", encoding="utf-8")
            plan = build_plan(folder, roster)
            match = next(item for item in plan if Path(item.source) == source)
            self.assertEqual(match.status, "ready")
            self.assertEqual(match.matched_name, "王小明")


if __name__ == "__main__":
    unittest.main()
