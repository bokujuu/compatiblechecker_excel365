@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul

rem --- Python 検出: Windows python → py -3 → WSL python3 ---
set "_PYMODE=win"
set "_PY=python"
%_PY% -V >nul 2>&1
if errorlevel 1 (
  set "_PY=py -3"
  %_PY% -V >nul 2>&1
)
if errorlevel 1 (
  wsl.exe -e python3 -V >nul 2>&1
  if errorlevel 1 (
    echo Python が見つかりません。Windows か WSL に Python を用意してください。
    exit /b 2
  )
  set "_PYMODE=wsl"
  set "_PY=wsl.exe -e python3"
)

rem --- 依存パッケージ確認＆必要なら自動インストール ---
%_PY% -c "import importlib.util,sys;sys.exit(0 if importlib.util.find_spec('openpyxl') else 1)"
if errorlevel 1 (
  echo 必要パッケージ^(openpyxl^)をインストールします...
  %_PY% -m pip install --quiet --upgrade pip
  %_PY% -m pip install --quiet openpyxl
)

rem pyxlsb は任意（.xlsb対応の参考用）
%_PY% -c "import importlib.util,sys;sys.exit(0 if importlib.util.find_spec('pyxlsb') else 1)"
if errorlevel 1 (
  rem 任意: 必要なら有効化してください
  rem %_PY% -m pip install --quiet pyxlsb
)

if "%~1"=="" (
  echo 使い方: この .bat に .xlsx/.xlsm ファイルをドラッグ&ドロップしてください。
  echo 例: 2016_365_test.xlsx をこの bat に重ねて離します。
  exit /b 1
)

rem --- スクリプトのパス（WSL用に変換も準備） ---
set "_SCRIPT_WIN=%~dp0excel2016_compat_check.py"
set "_SCRIPT_WSL="
if "%_PYMODE%"=="wsl" (
  for /f "usebackq delims=" %%S in (`wsl.exe wslpath -a "%_SCRIPT_WIN%"`) do set "_SCRIPT_WSL=%%S"
)

rem --- 引数（ドラッグ&ドロップされた各ファイル）を安全に1件ずつ処理 ---
:_loop
if "%~1"=="" goto _done
echo 解析中: %~f1
if "%_PYMODE%"=="wsl" (
  for /f "usebackq delims=" %%W in (`wsl.exe wslpath -a "%~f1"`) do set "_ARG_WSL=%%W"
  %_PY% "%_SCRIPT_WSL%" "%_ARG_WSL%"
) else (
  %_PY% "%_SCRIPT_WIN%" "%~f1"
)
if errorlevel 1 (
  echo 失敗: %~f1
) else (
  echo 完了: %~f1
)
shift
goto _loop

:_done
echo すべての処理が終了しました。
endlocal
exit /b 0
