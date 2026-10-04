@echo off
chcp 65001 >nul
REM missa_ppt 최초 1회 설치 스크립트(성인·청년 배포판 공용).
REM 대상 PC에 Python(3.9 이상 권장)이 이미 설치·PATH 등록돼 있어야 한다.
REM 이 파일이 있는 폴더에서 실행한다(배포 zip을 받아 압축을 푼 상태여야 함).

REM 어느 배포판인지는 폴더에 있는 실행 파일로 판단해 안내 문구에 그 유형만 표시한다.
set "KIND=미사 PPT"
set "RUNBAT=run_missa_*.bat"
if exist "%~dp0run_missa_성인.bat" (
    set "KIND=성인미사 PPT"
    set "RUNBAT=run_missa_성인.bat"
)
if exist "%~dp0run_missa_청년.bat" (
    set "KIND=청년미사 PPT"
    set "RUNBAT=run_missa_청년.bat"
)

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
    echo "%KIND%" 바로가기를 생성합니다...
    python "%~dp0create_shortcut.py"
    if errorlevel 1 (
        echo [경고] 바로가기 생성에 실패했습니다. %RUNBAT% 을 직접 실행해도 됩니다.
    )
)

echo.
echo 설치가 완료되었습니다. 바탕화면 또는 폴더 안의 "%KIND%" 바로가기 [또는 %RUNBAT%] 로 실행할 수 있습니다.
pause
