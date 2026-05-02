@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
title 출장 영수증 보고서 생성기

echo ============================================
echo   출장 영수증 보고서 생성기 v1.1
echo ============================================
echo.

REM 현재 bat 파일 위치로 이동
pushd "%~dp0"

REM Python 확인
python --version >nul 2>&1
if errorlevel 1 goto :no_python

echo   [OK] Python 확인 완료
echo.

REM 첫 실행 시 의존성 설치
if exist ".installed" goto :skip_install
echo   [설치] 필요한 라이브러리를 설치합니다...
echo   (최초 1회만, 몇 분 소요될 수 있습니다)
echo.
python -m pip install -r requirements.txt
if errorlevel 1 goto :install_fail
echo done> .installed
echo.
echo   [OK] 라이브러리 설치 완료

:skip_install
echo.
echo   서버를 시작합니다...
echo   브라우저에서 http://127.0.0.1:8500 이 열립니다.
echo   종료하려면 이 창을 닫거나 Ctrl+C를 누르세요.
echo ============================================
echo.

start "" "http://127.0.0.1:8500"
python app.py
goto :done

:no_python
echo   [ERROR] Python을 찾을 수 없습니다.
echo   Python 3.11 이상을 설치해주세요.
goto :done

:install_fail
echo.
echo   [ERROR] 라이브러리 설치 실패
goto :done

:done
echo.
popd
pause
