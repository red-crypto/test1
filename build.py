# -*- coding: utf-8 -*-
"""
급여명세서 발송기 — exe 빌드 스크립트 (빌드.bat 대체)

사용법 (Windows의 cmd, PowerShell, MSYS2, Git Bash 어디서든):
    py build.py                    # 일반 빌드
    py build.py --sample 샘플.pdf   # 실제 샘플 PDF로 전체 검사까지 한 뒤 빌드
    py build.py --clean-venv       # 빌드용 가상환경을 새로 만든 뒤 빌드
    python build.py                # MSYS용 python으로 실행해도 됨(아래 참고)

결과물:
    dist\\급여명세서발송기.exe
    package\\  ← exe + template.txt + 예시 명부 + README + CHANGELOG (배포용 묶음)

왜 .bat 대신 .py 인가
--------------------------------------------------------------
1) MSYS2 / Git Bash 에서는 PATH 앞쪽에 MSYS용 python이 잡힌다.
   그 python은 '윈도우용'이 아니라서 pywin32(Outlook 연동)가 설치되지 않는다
   (requirements.txt의 pywin32는 sys_platform == "win32" 일 때만 설치).
   → 빌드가 실패하거나, 빌드는 되는데 'Outlook 발송이 안 되는 exe'가 나온다.
   이 스크립트는 윈도우용 정식 파이썬인지 확인하고, 아니면 py 런처로
   정식 파이썬을 찾아 스스로 다시 실행한다.
2) 한글이 들어간 .bat은 cmd 코드페이지에 따라 줄이 깨지고, bash에서는
   ^ 줄 이어쓰기, %~dp0 같은 cmd 문법이 동작하지 않는다.
3) 빌드 전용 가상환경(.venv-build)을 써서 PC에 깔린 다른 패키지와 섞이지 않게 한다.
4) GitHub Actions도 이 스크립트로 빌드한다 → 로컬과 CI가 같은 옵션으로 빌드된다.
   PyInstaller 옵션을 바꿀 때는 이 파일만 고치면 된다.

빌드 전후 안전 점검
--------------------------------------------------------------
* 빌드 전: tkinter / win32com / pythoncom / pdfplumber / pypdf / pandas / openpyxl
  가 실제로 import 되는지 확인 → 하나라도 안 되면 빌드하지 않는다.
* 빌드 후: exe 안에 win32com·pdfplumber·payroll_core 등이 들어갔는지 확인
  → 빠졌으면 실패 처리(발송이 안 되는 exe가 배포되는 사고 방지).
"""

import argparse
import hashlib
import os
import re
import shutil
import subprocess
import sys
import sysconfig
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VENV_DIR = ROOT / ".venv-build"
BUILD_NAME = "PayrollSender"            # PyInstaller에는 ASCII 이름으로 빌드
EXE_NAME = "급여명세서발송기.exe"         # 빌드 후 한글 이름으로 바꾼다
PACKAGE_DIR = ROOT / "package"
PACKAGE_FILES = ["template.txt", "mapping_예시.xlsx", "README.md", "CHANGELOG.md"]
REEXEC_FLAG = "PAYROLL_BUILD_REEXEC"     # 무한 재실행 방지용 표시

# PyInstaller 옵션 (CI도 이 값을 그대로 쓴다)
#   --onefile   : 단일 실행파일
#   --windowed  : 콘솔 창 숨김(GUI 전용)
#   --hidden-import win32timezone : pywin32 빌드 시 누락되는 모듈
#   --collect-data pdfminer       : pdfplumber가 쓰는 pdfminer cmap 데이터(한글 추출에 필요)
PYINSTALLER_ARGS = [
    "--noconfirm", "--clean", "--onefile", "--windowed",
    "--name", BUILD_NAME,
    "--hidden-import", "win32timezone",
    "--collect-data", "pdfminer",
    "payroll_gui.py",
]

# 빌드 전에 import 되어야 하는 모듈과, 안 될 때 보여줄 설명
REQUIRED_MODULES = [
    ("tkinter", "화면(GUI). python.org 설치 프로그램에서 'tcl/tk and IDLE' 항목을 켜고 다시 설치하세요."),
    ("win32com.client", "Outlook 연동(pywin32). 윈도우용 정식 파이썬에서만 설치됩니다."),
    ("pythoncom", "Outlook 연동(pywin32). 윈도우용 정식 파이썬에서만 설치됩니다."),
    ("pdfplumber", "PDF 글자 읽기"),
    ("pypdf", "PDF 암호화"),
    ("pandas", "명부 읽기"),
    ("openpyxl", "엑셀 명부 읽기"),
]

# 빌드된 exe 안에 반드시 들어 있어야 하는 모듈
REQUIRED_IN_BUNDLE = [
    "win32com.client", "pdfplumber", "pypdf", "pandas", "openpyxl",
    "payroll_core", "payroll_outlook", "payroll_version",
]

NO_NATIVE_PYTHON_MSG = (
    "윈도우용 정식 파이썬을 찾지 못했습니다.\n"
    "  1) https://www.python.org/downloads/windows/ 에서 Python 3.11 이상을 설치하세요.\n"
    "     (설치 화면에서 'py launcher'와 'tcl/tk and IDLE'이 켜져 있어야 합니다)\n"
    "  2) 그다음 이 폴더에서  py build.py  로 실행하세요.\n"
    "  ※ MSYS2/Git Bash의 python(MSYS용)으로는 Outlook 연동이 들어간 exe를 만들 수 없습니다."
)


class BuildError(Exception):
    """사용자에게 그대로 보여줄 빌드 오류."""


# ------------------------------------------------------------
# 출력
# ------------------------------------------------------------
def _utf8_output():
    """MSYS 터미널(mintty)·CI 로그·cmd 어디서든 한글이 깨지지 않게 UTF-8로 출력."""
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def step(n, total, msg):
    print(f"\n[{n}/{total}] {msg}", flush=True)


def info(msg):
    print(f"      {msg}", flush=True)


# ------------------------------------------------------------
# 파이썬 판별 / MSYS 대응
# ------------------------------------------------------------
def is_native_windows_python(platform=None, version=None, sysconf_platform=None) -> bool:
    """
    윈도우용 정식 파이썬(python.org, MS Store 등 MSVC 빌드)인지 판별한다.
      - MSYS2/Cygwin python : sys.platform 이 'msys' 또는 'cygwin'
      - MinGW(GCC) python   : sys.platform 은 'win32'지만 MSVC가 아닌 GCC 빌드
    두 경우 모두 pywin32를 쓸 수 없으므로 '정식 아님'으로 본다.
    (인자는 테스트용. 보통은 비워 두면 현재 파이썬을 검사한다)
    """
    platform = sys.platform if platform is None else platform
    version = sys.version if version is None else version
    sysconf_platform = sysconfig.get_platform() if sysconf_platform is None else sysconf_platform
    if platform != "win32":
        return False
    if "MSC" not in version:
        return False
    if sysconf_platform.startswith("mingw"):
        return False
    return True


def to_windows_path(p) -> str:
    """
    MSYS/Cygwin 경로(/c/Users/...)를 윈도우 경로(C:\\Users\\...)로 바꾼다.
    윈도우용 파이썬에 넘기는 경로는 반드시 윈도우 경로여야 한다.
    """
    s = str(p)
    if os.name == "nt":                 # 이미 윈도우 경로
        return s
    cygpath = shutil.which("cygpath")
    if cygpath:
        try:
            out = subprocess.run([cygpath, "-w", s], capture_output=True, text=True, check=True)
            return out.stdout.strip() or s
        except Exception:
            pass
    return s


def clean_child_env(env=None) -> dict:
    """
    윈도우용 파이썬 자식 프로세스에 넘길 환경변수를 정리한다.
      - PYTHONHOME / PYTHONPATH : MSYS용 파이썬 경로가 들어 있으면 정식 파이썬이
        시작조차 못 한다 → 제거
      - TEMP / TMP : MSYS의 /tmp 같은 값이면 pip·PyInstaller가 임시 폴더를 못 찾는다
        → 윈도우 경로가 아니면 사용자 임시 폴더로 바꾼다
      - 한글 출력이 깨지지 않게 UTF-8 모드
    """
    env = dict(os.environ if env is None else env)
    for k in ("PYTHONHOME", "PYTHONPATH"):
        env.pop(k, None)
    if os.name == "nt":
        fallback = os.path.join(env.get("LOCALAPPDATA", os.path.expanduser("~")), "Temp")
        for k in ("TEMP", "TMP"):
            if not re.match(r"^[A-Za-z]:[\\/]", env.get(k, "")):
                env[k] = fallback
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def build_child_args(args) -> list:
    """다시 실행할 때 넘길 인자를 만든다(경로는 윈도우 경로로 변환)."""
    out = []
    if args.sample:
        out += ["--sample", to_windows_path(Path(args.sample).resolve())]
    if args.clean_venv:
        out.append("--clean-venv")
    if args.no_venv:
        out.append("--no-venv")
    if args.no_pause:
        out.append("--no-pause")
    return out


def reexec_with_native_python(args):
    """
    지금 파이썬이 정식 윈도우 파이썬이 아니면 py 런처로 정식 파이썬을 찾아
    이 스크립트를 다시 실행한다. (돌아오지 않음: 자식의 종료 코드로 끝낸다)
    """
    if os.environ.get(REEXEC_FLAG):
        raise BuildError(
            "py 런처로 다시 실행했지만 여전히 윈도우용 정식 파이썬이 아닙니다.\n"
            f"  현재: {sys.executable}\n  ({sys.version.splitlines()[0]})\n\n"
            + NO_NATIVE_PYTHON_MSG)
    py = shutil.which("py") or shutil.which("py.exe")
    if not py:
        raise BuildError(
            f"지금 실행 중인 파이썬은 윈도우용이 아닙니다: {sys.executable}\n\n"
            + NO_NATIVE_PYTHON_MSG)
    print(f"지금 파이썬은 윈도우용이 아닙니다({sys.executable}).")
    print("py 런처로 윈도우용 정식 파이썬을 찾아 다시 실행합니다...\n", flush=True)
    env = clean_child_env()
    env[REEXEC_FLAG] = "1"
    cmd = [py, "-3", to_windows_path(Path(__file__).resolve()), *build_child_args(args)]
    r = subprocess.run(cmd, env=env)
    sys.exit(r.returncode)


# ------------------------------------------------------------
# 빌드 단계
# ------------------------------------------------------------
def run(cmd, what, env=None, capture=False):
    """명령 실행. 실패하면 BuildError."""
    r = subprocess.run([str(c) for c in cmd], cwd=ROOT, env=env or clean_child_env(),
                       capture_output=capture, text=capture)
    if r.returncode != 0:
        detail = ""
        if capture:
            detail = "\n" + ((r.stdout or "") + (r.stderr or "")).strip()[-2000:]
        raise BuildError(f"{what}에 실패했습니다. 위 메시지를 확인하세요.{detail}")
    return r


def venv_python(vdir: Path) -> Path:
    return vdir / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def prepare_python(args) -> str:
    """빌드에 쓸 파이썬을 정한다(기본: 빌드 전용 가상환경)."""
    if args.no_venv:
        info(f"가상환경 없이 현재 파이썬 사용: {sys.executable}")
        return sys.executable
    if args.clean_venv and VENV_DIR.exists():
        info("기존 빌드용 가상환경 삭제")
        shutil.rmtree(VENV_DIR)
    vpy = venv_python(VENV_DIR)
    if vpy.exists():
        # 파이썬을 지우거나 업그레이드했으면 가상환경이 깨져 있을 수 있다 → 실행 확인
        ok = subprocess.run([str(vpy), "-c", "import sys"], capture_output=True,
                            env=clean_child_env()).returncode == 0
        if not ok:
            info("빌드용 가상환경이 깨져 있어 다시 만듭니다")
            shutil.rmtree(VENV_DIR, ignore_errors=True)
    if not vpy.exists():
        info("빌드용 가상환경(.venv-build)을 만드는 중... (처음 한 번만)")
        run([sys.executable, "-m", "venv", VENV_DIR], "가상환경 만들기")
    info(f"빌드용 파이썬: {vpy}")
    return str(vpy)


def install_packages(py: str, use_venv: bool):
    """requirements.txt + PyInstaller 설치. 가상환경은 내용이 바뀌었을 때만 다시 설치."""
    req = ROOT / "requirements.txt"
    key = hashlib.sha256(req.read_bytes() + b"|pyinstaller").hexdigest()
    stamp = VENV_DIR / ".installed-requirements" if use_venv else None
    if stamp and stamp.exists() and stamp.read_text(encoding="utf-8") == key:
        info("이미 설치되어 있습니다(requirements.txt 변경 없음) — 건너뜀")
        return
    info("필요한 패키지 설치 중... (처음엔 몇 분 걸립니다)")
    run([py, "-m", "pip", "install", "--disable-pip-version-check", "-q", "--upgrade", "pip"],
        "pip 업데이트")
    run([py, "-m", "pip", "install", "--disable-pip-version-check", "-r", req, "pyinstaller"],
        "패키지 설치")
    if stamp:
        stamp.write_text(key, encoding="utf-8")


IMPORT_CHECK_CODE = (
    "import importlib, sys\n"
    "bad = []\n"
    "for m in sys.argv[1:]:\n"
    "    try:\n"
    "        importlib.import_module(m)\n"
    "    except Exception as e:\n"
    "        bad.append(m + '\\t' + type(e).__name__ + ': ' + str(e))\n"
    "print('\\n'.join(bad))\n"
    "sys.exit(1 if bad else 0)\n"
)


def verify_imports(py: str):
    """빌드 전에 필수 모듈이 실제로 import 되는지 확인."""
    mods = [m for m, _ in REQUIRED_MODULES]
    r = subprocess.run([py, "-c", IMPORT_CHECK_CODE, *mods], cwd=ROOT, env=clean_child_env(),
                       capture_output=True, text=True)
    if r.returncode == 0:
        info("필수 모듈 모두 정상: " + ", ".join(mods))
        return
    why = dict(REQUIRED_MODULES)
    lines = []
    for line in (r.stdout or "").strip().splitlines():
        name, _, err = line.partition("\t")
        lines.append(f"  - {name}: {why.get(name, '')}\n      ({err})")
    if not lines:
        lines.append((r.stderr or "").strip()[-1500:])
    raise BuildError("필수 모듈을 불러오지 못해 빌드하지 않습니다.\n" + "\n".join(lines))


def run_tests(py: str, sample):
    cmd = [py, "test_core.py"]
    if sample:
        cmd.append(str(Path(sample).resolve()))
        info("샘플 PDF로 전체 검사 실행")
    else:
        info("PDF 없이 가능한 검사만 실행 (전체 검사: --sample 샘플.pdf)")
    run(cmd, "로직 점검")
    # 화면·Outlook 없이 실제 GUI 발송 흐름을 끝까지 검사(중복 실행 차단, 전달 확인 등)
    info("발송 흐름 검사 (가짜 화면·가짜 Outlook)")
    run([py, str(ROOT / "tests" / "test_gui_flow.py")], "발송 흐름 검사")


def run_pyinstaller(py: str):
    run([py, "-m", "PyInstaller", *PYINSTALLER_ARGS], "exe 빌드")


def verify_bundle(toc_path: Path = None):
    """
    빌드된 exe에 필수 모듈이 들어갔는지 PyInstaller 목록(PYZ-00.toc)으로 확인.
    목록 파일을 못 찾으면(PyInstaller 버전 차이 등) 경고만 하고 넘어간다.
    """
    toc = toc_path or (ROOT / "build" / BUILD_NAME / "PYZ-00.toc")
    if not toc.exists():
        info(f"⚠ exe 구성 목록({toc.name})을 찾지 못해 내용 확인을 건너뜁니다.")
        return
    text = toc.read_text(encoding="utf-8", errors="replace")
    missing = [m for m in REQUIRED_IN_BUNDLE if f"'{m}'" not in text]
    if missing:
        raise BuildError(
            "빌드된 exe에 필수 모듈이 빠져 있습니다: " + ", ".join(missing) + "\n"
            "이 exe는 발송이 안 될 수 있어 실패로 처리합니다.")
    info("exe 안에 필수 모듈 포함 확인: " + ", ".join(REQUIRED_IN_BUNDLE))


def finalize() -> Path:
    """한글 이름으로 바꾸고 배포용 묶음(package/)을 만든다."""
    built = ROOT / "dist" / (BUILD_NAME + (".exe" if os.name == "nt" else ""))
    if not built.exists():
        raise BuildError(f"빌드 결과물이 없습니다: {built}")
    final = ROOT / "dist" / EXE_NAME
    try:
        os.replace(built, final)
    except PermissionError:
        raise BuildError(f"{final} 을(를) 덮어쓰지 못했습니다. 그 exe가 실행 중이면 닫고 다시 하세요.")

    if PACKAGE_DIR.exists():
        shutil.rmtree(PACKAGE_DIR)
    PACKAGE_DIR.mkdir()
    shutil.copy2(final, PACKAGE_DIR / EXE_NAME)
    for name in PACKAGE_FILES:
        src = ROOT / name
        if src.exists():
            shutil.copy2(src, PACKAGE_DIR / name)
        else:
            info(f"⚠ 배포 묶음에 넣을 {name} 이(가) 없어 건너뜁니다")
    return final


def read_version() -> str:
    sys.path.insert(0, str(ROOT))
    try:
        from payroll_version import APP_VERSION
        return APP_VERSION
    except Exception:
        return "?"


# ------------------------------------------------------------
def parse_args(argv=None):
    p = argparse.ArgumentParser(description="급여명세서 발송기 exe 빌드")
    p.add_argument("--sample", metavar="PDF", help="실제 샘플 PDF로 전체 검사까지 실행")
    p.add_argument("--clean-venv", action="store_true", help="빌드용 가상환경을 새로 만듦")
    p.add_argument("--no-venv", action="store_true", help="가상환경 없이 현재 파이썬으로 빌드")
    p.add_argument("--no-pause", action="store_true", help="끝난 뒤 엔터 대기 안 함")
    return p.parse_args(argv)


def should_pause(args) -> bool:
    """탐색기에서 더블클릭으로 실행했을 때 창이 바로 닫히지 않게."""
    return (os.name == "nt" and not args.no_pause and not os.environ.get("CI")
            and sys.stdin is not None and sys.stdin.isatty())


def main(argv=None) -> int:
    _utf8_output()
    args = parse_args(argv)

    if not is_native_windows_python():
        reexec_with_native_python(args)          # 여기서 돌아오지 않는다

    os.chdir(ROOT)
    version = read_version()
    total = 8
    print("=" * 52)
    print(f"  급여명세서 발송기 exe 빌드  (프로그램 v{version})")
    print("=" * 52)

    step(1, total, "파이썬 확인")
    if sys.version_info < (3, 9):
        raise BuildError(f"파이썬 3.9 이상이 필요합니다(현재 {sys.version.split()[0]}).")
    info(f"{sys.executable}  ({sys.version.split()[0]}, 윈도우용 정식 파이썬)")

    step(2, total, "빌드용 파이썬 준비")
    py = prepare_python(args)

    step(3, total, "패키지 설치")
    install_packages(py, use_venv=not args.no_venv)

    step(4, total, "필수 모듈 확인 (Outlook 연동·화면·PDF)")
    verify_imports(py)

    step(5, total, "핵심 로직 점검")
    run_tests(py, args.sample)

    step(6, total, "exe 빌드 (2~5분)")
    run_pyinstaller(py)

    step(7, total, "exe 내용 확인")
    verify_bundle()

    step(8, total, "이름 변경 + 배포 묶음 만들기")
    final = finalize()
    size_mb = final.stat().st_size / (1024 * 1024)

    print("\n" + "=" * 52)
    print("  빌드 완료!")
    print(f"  결과물   : {final}  ({size_mb:.1f} MB)")
    print(f"  배포 묶음: {PACKAGE_DIR}")
    print(f"  버전     : v{version}  (실행하면 창 제목에 표시됨)")
    print("=" * 52)
    print("  대표님 PC에는 exe만 교체하세요(template.txt는 덮어쓰지 않기).")
    print("  처음 실행 시 '테스트 수신자'에 본인 메일을 넣고 한 번 시험 발송해 보세요.")
    return 0


if __name__ == "__main__":
    _args = None
    code = 1
    try:
        _args = parse_args()
        code = main()
    except BuildError as e:
        print(f"\n[오류] {e}", file=sys.stderr, flush=True)
        code = 1
    except KeyboardInterrupt:
        print("\n[중단] 사용자가 빌드를 중단했습니다.", file=sys.stderr)
        code = 130
    finally:
        if _args is not None and should_pause(_args):
            try:
                input("\n엔터를 누르면 창이 닫힙니다...")
            except Exception:
                pass
    sys.exit(code)
