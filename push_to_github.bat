@echo off
chcp 65001 > nul
title បង្ហោះប្រព័ន្ធ POS FIS ទៅកាន់ GitHub & Streamlit Cloud
cd /d "%~dp0"

echo ======================================================================
echo   🚀 UPLOAD / PUSH ប្រព័ន្ធ POS FIS ទៅកាន់ GITHUB
echo ======================================================================
echo.
echo   គណនី GitHub: Leang-Huoy (loanghuoy12@gmail.com)
echo   ទីតាំង Folder: %CD%
echo.
echo   ----------------------------------------------------------------------
echo   ជំហានទី ១: បញ្ចូលតំណភ្ជាប់ GitHub Repository របស់អ្នក
echo   (ប្រសិនបើមិនទាន់បានបង្កើត សូមចូលទៅ https://github.com/new រួចបង្កើត Repository ឈ្មោះ: pos-fis)
echo   ----------------------------------------------------------------------
echo.

set DEFAULT_REPO=https://github.com/Leang-Huoy/pos-fis.git
set /p REPO_URL="សូមបញ្ចូល GitHub Repo URL (ចុច Enter ដើម្បីប្រើ: %DEFAULT_REPO%): "
if "%REPO_URL%"=="" set REPO_URL=%DEFAULT_REPO%

echo.
echo   កំពុងកំណត់ Remote URL ទៅកាន់: %REPO_URL% ...
git remote remove origin 2>nul
git remote add origin %REPO_URL%
git branch -M main

echo   កំពុងរៀបចំឯកសារ និង Commit...
git add .
git commit -m "Update POS FIS system for Streamlit Cloud" 2>nul

echo.
echo   ----------------------------------------------------------------------
echo   ជំហានទី ២: កំពុង Push កូដទៅកាន់ GitHub...
echo   (ប្រសិនបើមានផ្ទាំង Browser លោតឡើង សូមចុច "Sign in with your browser")
echo   ----------------------------------------------------------------------
echo.

git push -u origin main

if %errorlevel% equ 0 (
    echo.
    echo ======================================================================
    echo   🎉 អបអរសាទរ! កូដត្រូវបាន Push ចូល GitHub ដោយជោគជ័យ!
    echo   🌐 ពិនិត្យ Repo របស់អ្នក: %REPO_URL%
    echo.
    echo   ----------------------------------------------------------------------
    echo   ជំហានបន្ទាប់ដើម្បី Hosting ឥតគិតថ្លៃលើ Streamlit Community Cloud:
    echo   1. បើកគេហទំព័រ: https://share.streamlit.io/
    echo   2. ចុចប៊ូតុង "Create app" ឬ "New app"
    echo   3. ជ្រើសរើស Repository របស់អ្នក (ឧ. Leang-Huoy/pos-fis)
    echo   4. កំណត់ Main file path: app.py
    echo   5. ចុចប៊ូតុង "Deploy!"
    echo   =^> ប្រព័ន្ធរបស់អ្នកនឹងទទួលបាន Link HTTPS ដំណើរការ 24/7 ឥតគិតថ្លៃ!
    echo ======================================================================
) else (
    echo.
    echo   ----------------------------------------------------------------------
    echo   ⚠️ ប្រសិនបើមានបញ្ហា Authentication សូមបញ្ចូល GitHub Personal Access Token (PAT):
    echo   (របៀបបង្កើត Token: GitHub -> Settings -> Developer settings -> Personal access tokens -> Tokens (classic))
    echo   ----------------------------------------------------------------------
    echo.
    set /p GH_TOKEN="សូម Paste GitHub Token (បើគ្មានទេ សូមចុច Enter ដើម្បីចាកចេញ): "
    if not "%GH_TOKEN%"=="" (
        git remote set-url origin https://%GH_TOKEN%@github.com/Leang-Huoy/pos-fis.git
        git push -u origin main
        if %errorlevel% equ 0 (
            echo.
            echo   🎉 ជោគជ័យ! កូដបាន Push ទៅកាន់ GitHub ដោយជោគជ័យតាម Token!
            echo   🌐 សូមចូលទៅកាន់: https://share.streamlit.io/ ដើម្បី Deploy ភ្លាមៗ!
        )
    )
)

echo.
pause
