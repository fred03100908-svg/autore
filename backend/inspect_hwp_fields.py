from pathlib import Path
import sys

try:
    import win32com.client as win32
except ImportError:
    print("pywin32가 설치되어 있지 않습니다.")
    print("아래 명령어를 먼저 실행하세요:")
    print(r".venv\Scripts\python.exe -m pip install pywin32")
    sys.exit(1)


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"


def find_latest_hwp_file():
    files = []

    if UPLOAD_DIR.exists():
        files.extend(UPLOAD_DIR.glob("*.hwp"))
        files.extend(UPLOAD_DIR.glob("*.hwpx"))

    if not files:
        return None

    return max(files, key=lambda file: file.stat().st_mtime)


def get_hwp_field_list(hwp):
    """
    HWP 템플릿 안에 있는 누름틀/필드 목록을 가져온다.
    한글 버전에 따라 호출 방식이 다를 수 있어서 여러 방식으로 시도한다.
    """
    candidates = [
        (),
        (0,),
        (0, 0),
        (1, 0),
    ]

    for args in candidates:
        try:
            raw = hwp.GetFieldList(*args)
            if raw:
                return raw
        except Exception:
            pass

    return ""


def split_fields(raw_field_text):
    if not raw_field_text:
        return []

    separators = ["\x02", "\r\n", "\n", "\r", ";", ","]

    fields = [raw_field_text]

    for sep in separators:
        next_fields = []
        for item in fields:
            next_fields.extend(item.split(sep))
        fields = next_fields

    cleaned = []

    for field in fields:
        name = field.strip()
        if name and name not in cleaned:
            cleaned.append(name)

    return cleaned


def main():
    if len(sys.argv) >= 2:
        hwp_path = Path(sys.argv[1])
    else:
        hwp_path = find_latest_hwp_file()

    if hwp_path is None:
        print("검사할 HWP/HWPX 파일을 찾지 못했습니다.")
        print("먼저 웹에서 템플릿 파일을 한 번 업로드하거나, 아래처럼 직접 경로를 넣어 실행하세요.")
        print(r'.venv\Scripts\python.exe inspect_hwp_fields.py "C:\경로\템플릿.hwp"')
        return

    if not hwp_path.exists():
        print(f"파일이 존재하지 않습니다: {hwp_path}")
        return

    print("=" * 60)
    print("검사할 한글 템플릿:")
    print(hwp_path)
    print("=" * 60)

    hwp = None

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass

        hwp.Open(str(hwp_path))

        raw_fields = get_hwp_field_list(hwp)
        fields = split_fields(raw_fields)

        print()
        print("[필드 검사 결과]")

        if fields:
            print(f"총 {len(fields)}개의 필드를 찾았습니다.")
            print()

            for index, field in enumerate(fields, start=1):
                print(f"{index}. {field}")

            print()
            print("결론: 이 템플릿은 필드 방식으로 자동 입력할 가능성이 높습니다.")
            print("다음 단계에서 PutFieldText 방식으로 값을 넣으면 됩니다.")

        else:
            print("필드를 찾지 못했습니다.")
            print()
            print("결론: 이 템플릿은 눈으로 보기에는 칸이 있지만,")
            print("프로그램이 바로 찾을 수 있는 필드명은 없는 상태일 가능성이 큽니다.")
            print()
            print("이 경우 다음 단계는 표 위치 방식으로 갑니다.")
            print("예: 첫 번째 표 2행 2열 = 학교명")

    except Exception as e:
        print("한글 템플릿 검사 중 오류가 발생했습니다.")
        print(str(e))

    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
    