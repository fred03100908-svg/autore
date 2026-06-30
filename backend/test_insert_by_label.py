from pathlib import Path
import sys
from datetime import datetime

try:
    import win32com.client as win32
except ImportError:
    print("pywin32가 설치되어 있지 않습니다.")
    print(r".venv\Scripts\python.exe -m pip install pywin32")
    sys.exit(1)


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


def find_latest_hwp_file():
    files = []

    if UPLOAD_DIR.exists():
        files.extend(UPLOAD_DIR.glob("*.hwp"))

    if not files:
        return None

    return max(files, key=lambda file: file.stat().st_mtime)


def move_doc_begin(hwp):
    try:
        hwp.HAction.Run("MoveDocBegin")
    except Exception:
        pass


def find_text(hwp, text):
    move_doc_begin(hwp)

    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = text
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1

    return hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)


def insert_text(hwp, text):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = str(text)
    return hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)


def cancel_find_selection(hwp):
    """
    검색된 글자 선택 상태만 해제.
    여기서 MoveRight를 쓰면 한 칸 더 밀려서 교장 칸에 들어가므로 쓰지 않는다.
    """
    try:
        hwp.HAction.Run("Cancel")
    except Exception:
        pass


def move_to_next_cell(hwp):
    """
    현재 셀에서 바로 다음 셀로 이동.
    학 교 명 셀 → 학교명 입력 빈칸
    """
    try:
        hwp.HAction.Run("TableRightCell")
        return True
    except Exception as e:
        print("[실패] 오른쪽 셀 이동 실패")
        print(e)
        return False


def fill_right_cell_of_label(hwp, label_text, value):
    print(f"검색 시도: {label_text}")

    found = find_text(hwp, label_text)

    if not found:
        print(f"[실패] '{label_text}'를 찾지 못했습니다.")
        return False

    print(f"[성공] '{label_text}'를 찾았습니다.")

    # 선택 상태만 해제. 커서를 오른쪽으로 움직이지 않음.
    cancel_find_selection(hwp)

    # 오른쪽 셀로 딱 한 번만 이동
    if not move_to_next_cell(hwp):
        return False

    print("[성공] 오른쪽 빈칸으로 이동했습니다.")

    try:
        insert_text(hwp, value)
        print(f"[입력 완료] {label_text} → {value}")
        return True
    except Exception as e:
        print("[실패] 값 입력 실패")
        print(e)
        return False


def main():
    hwp_path = find_latest_hwp_file()

    if hwp_path is None:
        print("uploads 폴더에서 HWP 템플릿을 찾지 못했습니다.")
        print("웹사이트에서 원본 HWP 템플릿을 다시 업로드한 뒤 ANALYZE EXCEL을 눌러주세요.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = OUTPUT_DIR / f"{timestamp}_학교명_한칸이동테스트.hwp"

    print("=" * 60)
    print("사용할 템플릿:")
    print(hwp_path)
    print()
    print("저장될 파일:")
    print(output_path)
    print("=" * 60)

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

        success = fill_right_cell_of_label(
            hwp=hwp,
            label_text="학 교 명",
            value="인덕원고등학교",
        )

        hwp.SaveAs(str(output_path))

        print()
        print("[완료] 파일을 저장했습니다.")
        print(output_path)

        if success:
            print("저장된 HWP에서 '학 교 명' 오른쪽 빈칸을 확인하세요.")
        else:
            print("아직 위치 조정이 필요합니다.")

    except Exception as e:
        print()
        print("[오류] HWP 자동 입력 중 문제가 발생했습니다.")
        print(str(e))

    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
    