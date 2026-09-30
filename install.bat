@echo off
chcp 65001 >nul
REM missa_ppt 최초 1회 설치 스크립트.
REM 대상 PC에 Python(3.9 이상 권장)이 이미 설치·PATH 등록돼 있어야 한다.
REM 이 파일이 있는 폴더에서 실행한다(저장소 zip을 이미 받아 압축을 푼 상태여야 함).

echo missa_ppt 의존 패키지를 설치합니다...
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [오류] 패키지 설치에 실패했습니다. Python/pip이 PATH에 등록돼 있는지 확인해 주세요.
    pause
    exit /b 1
)

echo.
if exist "%~dp0create_shortcut.py" (
    echo "청년미사 PPT" 바로가기를 생성합니다...
    python "%~dp0create_shortcut.py"
    if errorlevel 1 (
        echo [경고] 바로가기 생성에 실패했습니다 — run_missa_청년.bat을 직접 실행해도 됩니다.
    )
)

echo.
echo 설치가 완료되었습니다. run_missa.bat(또는 바탕화면 바로가기)로 실행할 수 있습니다.
pause
