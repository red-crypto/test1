@echo off
chcp 65001 > nul
REM ============================================================
REM  급여명세서 발송기 - exe 빌드 스크립트 (Windows 전용)
REM
REM  사용법: 이 파일을 더블클릭하거나, 명령 프롬프트에서 실행
REM  결과물: dist\급여명세서발송기.exe
REM
REM  ※ exe는 반드시 Windows에서 빌드해야 합니다.
REM    (PyInstaller는 크로스 컴파일을 지원하지 않습니다)
REM ============================================================

cd /d "%~dp0"
echo.
echo ============================================
echo   급여명세서 발송기 exe 빌드
echo ============================================
echo.

REM --- 파이썬 확인 ---
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] 파이썬을 찾을 수 없습니다.
    echo        https://www.python.org 에서 설치 후 다시 실행하세요.
    echo        설치 시 "Add Python to PATH" 체크 필수입니다.
    pause
    exit /b 1
)
echo [1/4] 파이썬 확인 완료
python --version

REM --- 패키지 설치 ---
echo.
echo [2/4] 필요한 패키지 설치 중... (처음엔 몇 분 걸립니다)
python -m pip install --upgrade pip > nul
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [오류] 패키지 설치에 실패했습니다. 인터넷 연결을 확인하세요.
    pause
    exit /b 1
)
python -m pip install pyinstaller
if errorlevel 1 (
    echo [오류] PyInstaller 설치에 실패했습니다.
    pause
    exit /b 1
)

REM --- 스모크 테스트 ---
echo.
echo [3/4] 핵심 로직 점검 중...
python test_core.py
if errorlevel 1 (
    echo [오류] 로직 점검에 실패했습니다. 빌드를 중단합니다.
    pause
    exit /b 1
)

REM --- 빌드 ---
echo.
echo [4/4] exe 빌드 중... (2~5분 걸립니다)
pyinstaller --noconfirm --clean --onefile --windowed ^
    --name PayrollSender ^
    --hidden-import win32timezone ^
    --collect-data pdfminer ^
    payroll_gui.py
if errorlevel 1 (
    echo [오류] 빌드에 실패했습니다. 위 메시지를 확인하세요.
    pause
    exit /b 1
)

REM --- 한글 이름으로 변경 ---
if exist "dist\급여명세서발송기.exe" del "dist\급여명세서발송기.exe"
ren "dist\PayrollSender.exe" "급여명세서발송기.exe"

echo.
echo ============================================
echo   빌드 완료!
echo   결과물: dist\급여명세서발송기.exe
echo ============================================
echo.
echo  대표님 PC로 옮길 때 함께 복사할 파일:
echo    - 급여명세서발송기.exe   (필수)
echo    - template.txt           (메일 문구, 없으면 자동 생성됨)
echo.
echo  ※ 처음 실행 시 '테스트 수신자'에 본인 메일을 넣고
echo     한 번 시험 발송해 보세요.
echo.
pause
