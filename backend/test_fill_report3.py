from pathlib import Path
import sys
import re
from datetime import datetime, date, time

try:
    import win32com.client as win32
except ImportError:
    print("pywin32가 설치되어 있지 않습니다.")
    print(r".venv\Scripts\python.exe -m pip install pywin32")
    sys.exit(1)

from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

REPORT3_SHEET_NAME = "보고서3"


def find_latest_file(patterns):
    files = []

    if UPLOAD_DIR.exists():
        for pattern in patterns:
            files.extend(UPLOAD_DIR.glob(pattern))

    if not files:
        return None

    return max(files, key=lambda file: file.stat().st_mtime)


def is_empty(value):
    return value is None or str(value).strip() == ""


def normalize_key(value):
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\n", "")
    text = text.replace("\r", "")
    text = text.replace(" ", "")
    text = text.replace("(", "")
    text = text.replace(")", "")
    text = text.replace("（", "")
    text = text.replace("）", "")
    text = text.replace("㎡", "")
    text = text.replace("㎥", "")
    text = text.replace("㎍", "")
    text = text.replace("μg", "")
    text = text.replace("ug", "")
    text = text.replace("/", "")
    text = text.replace("%", "")
    text = text.lower()

    return text


def decimal_places_from_format(number_format):
    if not number_format:
        return None

    fmt = str(number_format)

    if "." not in fmt:
        if "0" in fmt:
            return 0
        return None

    after_dot = fmt.split(".", 1)[1]
    count = 0

    for ch in after_dot:
        if ch in ["0", "#"]:
            count += 1
        else:
            break

    return count


def format_excel_cell(cell):
    value = cell.value

    if is_empty(value):
        return "-"

    if isinstance(value, datetime):
        return value.strftime("%Y. %m. %d.")

    if isinstance(value, date):
        return value.strftime("%Y. %m. %d.")

    if isinstance(value, time):
        return value.strftime("%H:%M")

    if isinstance(value, str):
        text = value.strip()
        return text if text else "-"

    if isinstance(value, float):
        places = decimal_places_from_format(cell.number_format)

        if places is not None:
            if places == 0:
                return str(int(round(value)))
            return f"{value:.{places}f}"

        if value.is_integer():
            return str(int(value))

        return f"{value:.3f}".rstrip("0").rstrip(".")

    return str(value).strip()


def read_block(sheet, start_row, end_row):
    rows = range(start_row, end_row + 1)

    return {
        "검사장소": [format_excel_cell(sheet.cell(row=r, column=2)) for r in rows],
        "검사횟수": [format_excel_cell(sheet.cell(row=r, column=4)) for r in rows],
        "최소": [format_excel_cell(sheet.cell(row=r, column=5)) for r in rows],
        "최대": [format_excel_cell(sheet.cell(row=r, column=6)) for r in rows],
        "평균": [format_excel_cell(sheet.cell(row=r, column=7)) for r in rows],
    }


def extract_report3_1_data(sheet):
    return [
        {
            "section": "3-1",
            "name": "PM10",
            "anchors": ["PM10"],
            "start_row": 5,
            "end_row": 11,
            "data": read_block(sheet, 5, 11),
        },
        {
            "section": "3-1",
            "name": "PM2.5",
            "anchors": ["PM2.5", "PM2"],
            "start_row": 12,
            "end_row": 16,
            "data": read_block(sheet, 12, 16),
        },
        {
            "section": "3-1",
            "name": "CO2",
            "anchors": ["CO₂", "CO2", "이산화탄소"],
            "start_row": 17,
            "end_row": 22,
            "data": read_block(sheet, 17, 22),
        },
        {
            "section": "3-1",
            "name": "HCHO",
            "anchors": ["HCHO", "포름알데하이드"],
            "start_row": 23,
            "end_row": 30,
            "data": read_block(sheet, 23, 30),
        },
        {
            "section": "3-1",
            "name": "총부유세균",
            "anchors": ["총부유세균", "부유세균"],
            "start_row": 31,
            "end_row": 34,
            "data": read_block(sheet, 31, 34),
        },
        {
            "section": "3-1",
            "name": "낙하세균",
            "anchors": ["낙하세균"],
            "start_row": 35,
            "end_row": 38,
            "data": read_block(sheet, 35, 38),
        },
        {
            "section": "3-1",
            "name": "CO",
            "anchors": ["일산화탄소"],
            "start_row": 39,
            "end_row": 41,
            "data": read_block(sheet, 39, 41),
        },
        {
            "section": "3-1",
            "name": "NO2",
            "anchors": ["NO₂", "NO2", "이산화질소"],
            "start_row": 42,
            "end_row": 44,
            "data": read_block(sheet, 42, 44),
        },
        {
            "section": "3-1",
            "name": "Rn",
            "anchors": ["Rn", "라돈"],
            "start_row": 45,
            "end_row": 46,
            "data": read_block(sheet, 45, 46),
        },
        {
            "section": "3-1",
            "name": "석면",
            "anchors": ["석면"],
            "start_row": 47,
            "end_row": 52,
            "data": read_block(sheet, 47, 52),
        },
        {
            "section": "3-1",
            "name": "오존",
            "anchors": ["오존", "O₃", "O3"],
            "start_row": 53,
            "end_row": 55,
            "data": read_block(sheet, 53, 55),
        },
        {
            "section": "3-1",
            "name": "진드기 등",
            "anchors": ["진드기"],
            "start_row": 56,
            "end_row": 56,
            "data": read_block(sheet, 56, 56),
        },
        {
            "section": "3-1",
            "name": "TVOC",
            "anchors": ["TVOC"],
            "start_row": 57,
            "end_row": 59,
            "data": read_block(sheet, 57, 59),
        },
        {
            "section": "3-1",
            "name": "벤젠",
            "anchors": ["벤젠"],
            "start_row": 60,
            "end_row": 62,
            "data": read_block(sheet, 60, 62),
        },
        {
            "section": "3-1",
            "name": "톨루엔",
            "anchors": ["톨루엔"],
            "start_row": 63,
            "end_row": 65,
            "data": read_block(sheet, 63, 65),
        },
        {
            "section": "3-1",
            "name": "에틸벤젠",
            "anchors": ["에틸벤젠"],
            "start_row": 66,
            "end_row": 68,
            "data": read_block(sheet, 66, 68),
        },
        {
            "section": "3-1",
            "name": "자일렌",
            "anchors": ["자일렌"],
            "start_row": 69,
            "end_row": 71,
            "data": read_block(sheet, 69, 71),
        },
        {
            "section": "3-1",
            "name": "스티렌",
            "anchors": ["스티렌"],
            "start_row": 72,
            "end_row": 74,
            "data": read_block(sheet, 72, 74),
        },
    ]


def find_excel_item_rows(sheet, item_specs, search_start=75, search_end=119):
    found_items = []

    for row in range(search_start, search_end + 1):
        value = sheet.cell(row=row, column=1).value
        key = normalize_key(value)

        if not key:
            continue

        for spec in item_specs:
            for excel_key in spec["excel_keys"]:
                if normalize_key(excel_key) in key:
                    found_items.append(
                        {
                            **spec,
                            "start_row": row,
                        }
                    )
                    break

    unique = []
    seen = set()

    for item in found_items:
        if item["name"] in seen:
            continue
        seen.add(item["name"])
        unique.append(item)

    unique.sort(key=lambda item: item["start_row"])

    for index, item in enumerate(unique):
        if index + 1 < len(unique):
            item["end_row"] = unique[index + 1]["start_row"] - 1
        else:
            item["end_row"] = search_end

    return unique


def extract_report3_2_data(sheet):
    item_specs = [
        {
            "section": "3-2",
            "name": "조도_칠판면",
            "excel_keys": ["조도칠판면", "조도(칠판면)"],
            "anchors": ["조도"],
            "hwp_occurrence": 1,
        },
        {
            "section": "3-2",
            "name": "조도_책상면",
            "excel_keys": ["조도책상면", "조도(책상면)"],
            "anchors": ["조도"],
            "hwp_occurrence": 2,
        },
        {
            "section": "3-2",
            "name": "조도비_칠판면",
            "excel_keys": ["조도비칠판면", "조도비(칠판면)"],
            "anchors": ["조도비"],
            "hwp_occurrence": 1,
        },
        {
            "section": "3-2",
            "name": "조도비_책상면",
            "excel_keys": ["조도비책상면", "조도비(책상면)"],
            "anchors": ["조도비"],
            "hwp_occurrence": 2,
        },
        {
            "section": "3-2",
            "name": "온도",
            "excel_keys": ["온도"],
            "anchors": ["온도"],
            "hwp_occurrence": 1,
        },
        {
            "section": "3-2",
            "name": "습도",
            "excel_keys": ["습도"],
            "anchors": ["습도"],
            "hwp_occurrence": 1,
        },
        {
            "section": "3-2",
            "name": "소음",
            "excel_keys": ["소음"],
            "anchors": ["소음"],
            "hwp_occurrence": 1,
        },
        {
            "section": "3-2",
            "name": "환기",
            "excel_keys": ["환기"],
            "anchors": ["환기"],
            "hwp_occurrence": 1,
        },
    ]

    items = find_excel_item_rows(
        sheet=sheet,
        item_specs=item_specs,
        search_start=75,
        search_end=119,
    )

    blocks = []

    for item in items:
        blocks.append(
            {
                "section": "3-2",
                "name": item["name"],
                "anchors": item["anchors"],
                "hwp_occurrence": item["hwp_occurrence"],
                "start_row": item["start_row"],
                "end_row": item["end_row"],
                "data": read_block(sheet, item["start_row"], item["end_row"]),
            }
        )

    return blocks


def extract_report3_data(excel_path):
    workbook = load_workbook(excel_path, data_only=True, keep_vba=True)

    if REPORT3_SHEET_NAME not in workbook.sheetnames:
        raise ValueError(f"{REPORT3_SHEET_NAME} 시트를 찾지 못했습니다.")

    sheet = workbook[REPORT3_SHEET_NAME]

    blocks_3_1 = extract_report3_1_data(sheet)
    blocks_3_2 = extract_report3_2_data(sheet)

    return blocks_3_1 + blocks_3_2


def run(hwp, action_name):
    try:
        hwp.HAction.Run(action_name)
        return True
    except Exception:
        return False


def move_doc_begin(hwp):
    run(hwp, "MoveDocBegin")


def find_text(hwp, text, from_begin=True):
    if from_begin:
        move_doc_begin(hwp)

    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = text
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1

    return hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)


def insert_text(hwp, text):
    text = "-" if text is None or str(text).strip() == "" else str(text)

    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    return hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)


def cancel_selection(hwp):
    run(hwp, "Cancel")


def move_right(hwp, count=1):
    for _ in range(count):
        if not run(hwp, "TableRightCell"):
            return False
    return True


def move_down(hwp, count=1):
    for _ in range(count):
        if not run(hwp, "TableLowerCell"):
            return False
    return True


def clear_current_cell_text(hwp):
    cancel_selection(hwp)

    try:
        run(hwp, "MoveLineBegin")
        run(hwp, "MoveSelLineEnd")
        run(hwp, "Delete")
        cancel_selection(hwp)
        return True
    except Exception:
        cancel_selection(hwp)

    return False


def write_current_cell(hwp, value):
    value = "-" if value is None or str(value).strip() == "" else str(value)

    clear_current_cell_text(hwp)

    try:
        insert_text(hwp, value)
        return True
    except Exception as e:
        print("[실패] 셀 입력 실패")
        print(e)
        return False


def go_to_section(hwp, section_name):
    found = find_text(hwp, section_name, from_begin=True)

    if not found:
        print(f"[실패] HWP에서 '{section_name}'를 찾지 못했습니다.")
        return False

    cancel_selection(hwp)
    return True


def go_to_item_anchor(hwp, block):
    section_name = block["section"]

    if not go_to_section(hwp, section_name):
        return False

    occurrence = block.get("hwp_occurrence", 1)

    for anchor in block["anchors"]:
        found_count = 0

        # 3-2의 조도/조도비처럼 같은 단어가 반복될 수 있음
        while found_count < occurrence:
            found = find_text(hwp, anchor, from_begin=False)

            if not found:
                break

            found_count += 1
            cancel_selection(hwp)

        if found_count == occurrence:
            print(f"[성공] {block['name']} 기준 글자 찾음: {anchor} / {occurrence}번째")
            return True

    print(f"[실패] HWP에서 {block['name']} 항목을 찾지 못했습니다. anchors={block['anchors']}")
    return False


def go_to_item_cell(hwp, block, right_count):
    if right_count >= 7:
        raise ValueError(
            "이번 버전에서는 검사시간/유지기준/평가결과/측정기기 사양 영역을 건드리지 않습니다."
        )

    if not go_to_item_anchor(hwp, block):
        return False

    if not move_right(hwp, right_count):
        return False

    return True


def fill_vertical_column(hwp, block, label, right_count, values):
    print()
    print(f"[{block['section']} / {block['name']}] 세로 입력: {label} / 오른쪽 {right_count}칸")

    if not go_to_item_cell(hwp, block, right_count):
        print(f"[실패] {block['name']} {label} 첫 칸 이동 실패")
        return 0

    success_count = 0

    for index, value in enumerate(values, start=1):
        print(f"  {index}/{len(values)} → {value}")

        if write_current_cell(hwp, value):
            success_count += 1

        if index < len(values):
            if not move_down(hwp, 1):
                print(f"[실패] {block['name']} {label} 아래 이동 실패")
                return success_count

    return success_count


def fill_block(hwp, block):
    data = block["data"]
    total = 0

    total += fill_vertical_column(hwp, block, "검사장소", 1, data["검사장소"])
    total += fill_vertical_column(hwp, block, "검사횟수", 3, data["검사횟수"])
    total += fill_vertical_column(hwp, block, "최소", 4, data["최소"])
    total += fill_vertical_column(hwp, block, "최대", 5, data["최대"])
    total += fill_vertical_column(hwp, block, "평균", 6, data["평균"])

    return total


def main():
    excel_path = find_latest_file(["*.xlsm", "*.xlsx", "*.xls"])
    hwp_path = find_latest_file(["*.hwp"])

    if excel_path is None:
        print("uploads 폴더에서 엑셀 파일을 찾지 못했습니다.")
        print("웹사이트에서 엑셀 파일 업로드 후 ANALYZE EXCEL을 먼저 눌러주세요.")
        return

    if hwp_path is None:
        print("uploads 폴더에서 HWP 템플릿을 찾지 못했습니다.")
        print("웹사이트에서 원본 HWP 템플릿 업로드 후 ANALYZE EXCEL을 먼저 눌러주세요.")
        return

    blocks = extract_report3_data(excel_path)

    print("=" * 60)
    print("사용할 엑셀:")
    print(excel_path)
    print()
    print("사용할 HWP 템플릿:")
    print(hwp_path)
    print()
    print("입력 대상:")
    for block in blocks:
        print(
            f"- {block['section']} / {block['name']} / "
            f"엑셀 {block['start_row']}~{block['end_row']}행"
        )
        print("  검사장소:", block["data"]["검사장소"])
        print("  검사횟수:", block["data"]["검사횟수"])
        print("  최소:", block["data"]["최소"])
        print("  최대:", block["data"]["최대"])
        print("  평균:", block["data"]["평균"])
    print("=" * 60)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = OUTPUT_DIR / f"{timestamp}_보고서3_3-1_3-2_측정값만_입력테스트.hwp"

    hwp = None

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

        try:
            hwp.XHwpWindows.Item(0).Visible = True
        except Exception:
            pass

        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass

        hwp.Open(str(hwp_path))

        total = 0

        for block in blocks:
            print()
            print("=" * 60)
            print(f"[{block['section']} / {block['name']}] 입력 시작")
            print("=" * 60)

            total += fill_block(hwp, block)

        hwp.SaveAs(str(output_path))

        print()
        print("[완료] 보고서3 3-1 + 3-2 측정값 입력 테스트 HWP를 저장했습니다.")
        print(output_path)
        print(f"[결과] 총 입력 셀 수: {total}")
        print()
        print("이번 버전은 검사시간, 유지기준, 평가결과, 측정기기 사양을 건드리지 않습니다.")

    except Exception as e:
        print()
        print("[오류] 보고서3 입력 중 문제가 발생했습니다.")
        print(str(e))

    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
    