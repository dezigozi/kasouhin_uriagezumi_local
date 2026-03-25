@echo off
chcp 65001 >nul 2>&1
set "SRC=%~dp0.."
set "DST=%~dp0..\バックアップ_20260220"

echo Backing up to: %DST%

robocopy "%SRC%\lease-report-app" "%DST%\lease-report-app" /MIR /XD node_modules /NFL /NDL /NP
if exist "%SRC%\リース実績レポート.bat" copy /Y "%SRC%\リース実績レポート.bat" "%DST%\"
if exist "%SRC%\app-icon.ico" copy /Y "%SRC%\app-icon.ico" "%DST%\"
if exist "%SRC%\EXEビルド手順.md" copy /Y "%SRC%\EXEビルド手順.md" "%DST%\"
if exist "%SRC%\リース実績レポート.lnk" copy /Y "%SRC%\リース実績レポート.lnk" "%DST%\"

echo Backup complete!
