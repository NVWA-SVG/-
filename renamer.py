from __future__ import annotations

import argparse
import csv
import json
import re
import unicodedata
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


NAME_HEADERS = {"姓名", "名字", "name"}
ID_HEADERS = {"学号", "学生学号", "student_id", "studentid", "id"}
PROJECT_HEADERS = {"项目", "项目名称", "project"}
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
HISTORY_DIR_NAME = ".rename_history"


class RenameError(Exception):
    """可以直接显示给用户的错误。"""


@dataclass(frozen=True)
class Student:
    name: str
    student_id: str
    project: str = ""


@dataclass
class RenameItem:
    source: str
    target: str
    status: str
    reason: str = ""
    matched_name: str = ""


def normalize_name(value: str) -> str:
    """用于匹配的名称标准化：统一全/半角并去掉首尾空白。"""
    return unicodedata.normalize("NFKC", str(value or "")).strip()


def compact_match_text(value: str) -> str:
    """去掉空格、下划线和标点，用于从混合文件名中识别姓名/学号。"""
    return "".join(ch for ch in normalize_name(value).casefold() if ch.isalnum())


def safe_filename_part(value: str) -> str:
    value = normalize_name(value)
    value = INVALID_FILENAME_CHARS.sub("＿", value).rstrip(". ")
    return value


def _find_header(fieldnames: Iterable[str], aliases: set[str]) -> str | None:
    lookup = {normalize_name(field).casefold(): field for field in fieldnames if field}
    for alias in aliases:
        found = lookup.get(alias.casefold())
        if found:
            return found
    return None


def _read_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        sample = handle.read(4096)
        handle.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(handle, dialect=dialect)
        if not reader.fieldnames:
            raise RenameError("花名册没有表头。")
        return [
            {str(key or ""): str(value or "") for key, value in row.items()}
            for row in reader
            if any(str(value or "").strip() for value in row.values())
        ]


def _column_index(cell_reference: str) -> int:
    letters = "".join(ch for ch in cell_reference if ch.isalpha()).upper()
    result = 0
    for ch in letters:
        result = result * 26 + ord(ch) - ord("A") + 1
    return result - 1


def _read_xlsx_rows(path: Path) -> list[dict[str, str]]:
    """用标准库读取 xlsx 的第一个工作表，避免额外安装依赖。"""
    namespaces = {
        "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "p": "http://schemas.openxmlformats.org/package/2006/relationships",
    }
    try:
        with zipfile.ZipFile(path) as book:
            shared: list[str] = []
            if "xl/sharedStrings.xml" in book.namelist():
                root = ET.fromstring(book.read("xl/sharedStrings.xml"))
                for item in root.findall("m:si", namespaces):
                    shared.append("".join(node.text or "" for node in item.iterfind(".//m:t", namespaces)))

            workbook = ET.fromstring(book.read("xl/workbook.xml"))
            first_sheet = workbook.find("m:sheets/m:sheet", namespaces)
            if first_sheet is None:
                raise RenameError("Excel 花名册中没有工作表。")
            relation_id = first_sheet.attrib.get(f"{{{namespaces['r']}}}id")

            relations = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
            target = None
            for relation in relations.findall("p:Relationship", namespaces):
                if relation.attrib.get("Id") == relation_id:
                    target = relation.attrib.get("Target")
                    break
            if not target:
                raise RenameError("无法找到 Excel 的第一个工作表。")
            sheet_path = target.lstrip("/")
            if not sheet_path.startswith("xl/"):
                sheet_path = f"xl/{sheet_path}"
            sheet = ET.fromstring(book.read(sheet_path))

            table: list[list[str]] = []
            for row in sheet.findall(".//m:sheetData/m:row", namespaces):
                values: list[str] = []
                for cell in row.findall("m:c", namespaces):
                    index = _column_index(cell.attrib.get("r", "A1"))
                    while len(values) <= index:
                        values.append("")
                    cell_type = cell.attrib.get("t")
                    if cell_type == "inlineStr":
                        value = "".join(node.text or "" for node in cell.iterfind(".//m:t", namespaces))
                    else:
                        value_node = cell.find("m:v", namespaces)
                        value = value_node.text if value_node is not None and value_node.text else ""
                        if cell_type == "s" and value:
                            value = shared[int(value)]
                    values[index] = value
                table.append(values)
    except (KeyError, zipfile.BadZipFile, ET.ParseError, ValueError, IndexError) as exc:
        raise RenameError(f"无法读取 Excel 花名册：{exc}") from exc

    if not table:
        raise RenameError("花名册是空的。")
    headers = [normalize_name(value) for value in table[0]]
    return [
        {headers[index]: row[index] if index < len(row) else "" for index in range(len(headers))}
        for row in table[1:]
        if any(str(value).strip() for value in row)
    ]


def load_roster(path: str | Path) -> dict[str, Student]:
    roster_path = Path(path).expanduser().resolve()
    if not roster_path.is_file():
        raise RenameError(f"花名册不存在：{roster_path}")
    suffix = roster_path.suffix.casefold()
    if suffix in {".csv", ".tsv", ".txt"}:
        rows = _read_csv_rows(roster_path)
    elif suffix == ".xlsx":
        rows = _read_xlsx_rows(roster_path)
    else:
        raise RenameError("花名册仅支持 .csv、.tsv 或 .xlsx 格式。")
    if not rows:
        raise RenameError("花名册中没有数据。")

    fields = list(rows[0].keys())
    name_field = _find_header(fields, NAME_HEADERS)
    id_field = _find_header(fields, ID_HEADERS)
    project_field = _find_header(fields, PROJECT_HEADERS)
    if not name_field or not id_field:
        raise RenameError("花名册必须包含“姓名”和“学号”两列。")

    students: dict[str, Student] = {}
    duplicates: list[str] = []
    used_student_ids: dict[str, str] = {}
    duplicate_ids: list[str] = []
    for line_number, row in enumerate(rows, start=2):
        name = normalize_name(row.get(name_field, ""))
        student_id = normalize_name(row.get(id_field, ""))
        project = normalize_name(row.get(project_field, "")) if project_field else ""
        if not name and not student_id:
            continue
        if not name or not student_id:
            raise RenameError(f"花名册第 {line_number} 行的姓名或学号为空。")
        key = name.casefold()
        if key in students:
            duplicates.append(name)
        student_id_key = compact_match_text(student_id)
        if student_id_key in used_student_ids:
            duplicate_ids.append(student_id)
        used_student_ids[student_id_key] = name
        students[key] = Student(name=name, student_id=student_id, project=project)
    if duplicates:
        raise RenameError("花名册存在重名，无法安全匹配：" + "、".join(sorted(set(duplicates))))
    if duplicate_ids:
        raise RenameError("花名册存在重复学号，无法安全匹配：" + "、".join(sorted(set(duplicate_ids))))
    if not students:
        raise RenameError("花名册中没有有效学生数据。")
    return students


def match_student(file_stem: str, students: dict[str, Student]) -> tuple[Student | None, str, str]:
    """
    返回（学生，匹配说明，冲突说明）。

    优先级：完整姓名 > 文件名内学号 > 文件名内姓名。学号与姓名指向
    不同学生时不猜测，直接报冲突。
    """
    normalized_stem = normalize_name(file_stem)
    exact = students.get(normalized_stem.casefold())
    if exact:
        return exact, "完整姓名匹配", ""

    compact_stem = compact_match_text(file_stem)
    roster = list(students.values())
    id_matches = [
        student
        for student in roster
        if compact_match_text(student.student_id)
        and compact_match_text(student.student_id) in compact_stem
    ]
    name_matches = [
        student
        for student in roster
        if compact_match_text(student.name)
        and compact_match_text(student.name) in compact_stem
    ]

    # 学号长度不一时，优先最长的完整学号，避免 123 误中 1234。
    if id_matches:
        longest = max(len(compact_match_text(student.student_id)) for student in id_matches)
        id_matches = [
            student for student in id_matches if len(compact_match_text(student.student_id)) == longest
        ]
    if len(id_matches) > 1:
        names = "、".join(student.name for student in id_matches)
        return None, "", f"文件名中匹配到多个学号：{names}"

    if len(id_matches) == 1:
        by_id = id_matches[0]
        other_names = [student for student in name_matches if student != by_id]
        if other_names:
            names = "、".join(student.name for student in other_names)
            return None, "", f"学号指向“{by_id.name}”，但文件名又包含其他姓名：{names}"
        return by_id, "按文件名中的学号识别", ""

    if name_matches:
        longest = max(len(compact_match_text(student.name)) for student in name_matches)
        longest_matches = [
            student for student in name_matches if len(compact_match_text(student.name)) == longest
        ]
        if len(longest_matches) > 1:
            names = "、".join(student.name for student in longest_matches)
            return None, "", f"文件名中匹配到多个姓名：{names}"
        return longest_matches[0], "按文件名中的姓名识别", ""

    return None, "", ""


def render_target(template: str, student: Student, project_override: str) -> str:
    project = normalize_name(project_override) or student.project
    values = {
        "姓名": safe_filename_part(student.name),
        "学号": safe_filename_part(student.student_id),
        "项目": safe_filename_part(project),
    }
    if "{项目}" in template and not project:
        raise RenameError(f"学生“{student.name}”没有项目名称。")
    try:
        result = template.format(**values)
    except KeyError as exc:
        raise RenameError(f"命名模板中存在未知字段：{exc}") from exc
    result = safe_filename_part(result)
    if not result:
        raise RenameError("命名模板生成了空文件名。")
    return result


def build_plan(
    folder: str | Path,
    roster_path: str | Path,
    project: str = "",
    template: str = "{姓名}_{学号}_{项目}",
) -> list[RenameItem]:
    target_folder = Path(folder).expanduser().resolve()
    roster_file = Path(roster_path).expanduser().resolve()
    if not target_folder.is_dir():
        raise RenameError(f"目标文件夹不存在：{target_folder}")
    students = load_roster(roster_file)
    items: list[RenameItem] = []
    reserved: dict[str, Path] = {}

    files = sorted(
        (path for path in target_folder.iterdir() if path.is_file()),
        key=lambda path: path.name.casefold(),
    )
    for source in files:
        if source.resolve() == roster_file:
            items.append(RenameItem(str(source), str(source), "skipped", "花名册文件"))
            continue
        student, match_reason, match_conflict = match_student(source.stem, students)
        if match_conflict:
            items.append(RenameItem(str(source), str(source), "conflict", match_conflict))
            continue
        if not student:
            items.append(
                RenameItem(
                    str(source),
                    str(source),
                    "skipped",
                    "文件名中未识别到花名册姓名或学号",
                )
            )
            continue
        try:
            new_stem = render_target(template, student, project)
        except RenameError as exc:
            items.append(RenameItem(str(source), str(source), "conflict", str(exc), student.name))
            continue
        target = source.with_name(new_stem + source.suffix)
        if target.name == source.name:
            items.append(RenameItem(str(source), str(target), "skipped", "文件名已符合模板", student.name))
            continue
        target_key = str(target).casefold()
        if target_key in reserved:
            items.append(RenameItem(str(source), str(target), "conflict", "多个文件将生成同一目标名", student.name))
            continue
        if target.exists() and target.resolve() != source.resolve():
            items.append(RenameItem(str(source), str(target), "conflict", "目标文件已存在", student.name))
            continue
        reserved[target_key] = source
        items.append(RenameItem(str(source), str(target), "ready", match_reason, student.name))
    return items


def plan_summary(items: Iterable[RenameItem]) -> dict[str, int]:
    summary = {"ready": 0, "skipped": 0, "conflict": 0}
    for item in items:
        summary[item.status] = summary.get(item.status, 0) + 1
    return summary


def execute_plan(items: list[RenameItem], folder: str | Path) -> Path:
    conflicts = [item for item in items if item.status == "conflict"]
    if conflicts:
        raise RenameError("预览中存在冲突，请先解决后再执行。")
    ready = [item for item in items if item.status == "ready"]
    if not ready:
        raise RenameError("没有可以重命名的文件。")

    completed: list[RenameItem] = []
    try:
        for item in ready:
            source, target = Path(item.source), Path(item.target)
            if not source.is_file():
                raise RenameError(f"源文件已不存在：{source.name}")
            if target.exists() and target.resolve() != source.resolve():
                raise RenameError(f"目标文件已存在：{target.name}")
            source.rename(target)
            completed.append(item)
    except Exception as exc:
        rollback_errors: list[str] = []
        for item in reversed(completed):
            try:
                Path(item.target).rename(Path(item.source))
            except OSError as rollback_exc:
                rollback_errors.append(str(rollback_exc))
        detail = f"重命名失败，已尝试回滚：{exc}"
        if rollback_errors:
            detail += "；回滚错误：" + " | ".join(rollback_errors)
        raise RenameError(detail) from exc

    history_dir = Path(folder).expanduser().resolve() / HISTORY_DIR_NAME
    history_dir.mkdir(exist_ok=True)
    history_path = history_dir / f"rename_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.json"
    payload = {
        "created_at": datetime.now().astimezone().isoformat(),
        "folder": str(Path(folder).expanduser().resolve()),
        "operations": [asdict(item) for item in completed],
    }
    history_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return history_path


def latest_history(folder: str | Path) -> Path:
    history_dir = Path(folder).expanduser().resolve() / HISTORY_DIR_NAME
    histories = sorted(history_dir.glob("rename_*.json"), reverse=True) if history_dir.is_dir() else []
    if not histories:
        raise RenameError("该文件夹没有可撤销的重命名记录。")
    return histories[0]


def undo_history(history_path: str | Path) -> int:
    path = Path(history_path).expanduser().resolve()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        operations = payload["operations"]
    except (OSError, json.JSONDecodeError, KeyError, TypeError) as exc:
        raise RenameError(f"无法读取撤销记录：{exc}") from exc

    for operation in operations:
        original = Path(operation["source"])
        current = Path(operation["target"])
        if not current.is_file():
            raise RenameError(f"无法撤销，文件不存在：{current.name}")
        if original.exists():
            raise RenameError(f"无法撤销，原文件名已被占用：{original.name}")

    completed: list[tuple[Path, Path]] = []
    try:
        for operation in reversed(operations):
            original = Path(operation["source"])
            current = Path(operation["target"])
            current.rename(original)
            completed.append((original, current))
    except OSError as exc:
        for original, current in reversed(completed):
            try:
                original.rename(current)
            except OSError:
                pass
        raise RenameError(f"撤销失败：{exc}") from exc

    undone_path = path.with_suffix(path.suffix + ".undone")
    path.rename(undone_path)
    return len(operations)


def _print_plan(items: list[RenameItem]) -> None:
    labels = {"ready": "待改名", "skipped": "跳过", "conflict": "冲突"}
    for item in items:
        source = Path(item.source).name
        target = Path(item.target).name
        arrow = f" -> {target}" if source != target else ""
        reason = f" ({item.reason})" if item.reason else ""
        print(f"[{labels.get(item.status, item.status)}] {source}{arrow}{reason}")
    summary = plan_summary(items)
    print(f"\n合计：待改名 {summary['ready']}，跳过 {summary['skipped']}，冲突 {summary['conflict']}")


def main() -> int:
    parser = argparse.ArgumentParser(description="按花名册批量重命名文件")
    subparsers = parser.add_subparsers(dest="command", required=True)
    for command in ("preview", "rename"):
        child = subparsers.add_parser(command, help="预览" if command == "preview" else "执行重命名")
        child.add_argument("--folder", required=True, help="待处理文件夹")
        child.add_argument("--roster", required=True, help="花名册 CSV/TSV/XLSX")
        child.add_argument("--project", default="", help="统一项目名，留空则使用花名册的项目列")
        child.add_argument("--template", default="{姓名}_{学号}_{项目}", help="命名模板")
    undo = subparsers.add_parser("undo", help="撤销最近一次重命名")
    undo.add_argument("--folder", required=True, help="已处理文件夹")
    args = parser.parse_args()

    try:
        if args.command == "undo":
            count = undo_history(latest_history(args.folder))
            print(f"已撤销 {count} 个文件。")
            return 0
        items = build_plan(args.folder, args.roster, args.project, args.template)
        _print_plan(items)
        if args.command == "rename":
            history = execute_plan(items, args.folder)
            print(f"\n重命名完成。撤销记录：{history}")
        return 2 if any(item.status == "conflict" for item in items) else 0
    except RenameError as exc:
        print(f"错误：{exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
