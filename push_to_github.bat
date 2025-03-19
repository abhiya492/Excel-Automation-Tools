@echo off
echo Excel Automation Tools - GitHub Push Script
echo ===========================================
echo.

REM Check if Git is installed
where git >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Git is not installed or not in your PATH.
    echo Please install Git from https://git-scm.com/downloads
    pause
    exit /b 1
)

echo Initializing Git repository...
git init

echo Creating .gitignore file...
REM The .gitignore file should already exist in the directory

echo Adding files to Git...
git add .

echo Committing changes...
set /p commit_msg="Enter commit message (default: Initial commit): "
if "%commit_msg%"=="" set commit_msg=Initial commit: Excel Automation Tools project
git commit -m "%commit_msg%"

echo.
echo Connecting to GitHub repository...
echo.
echo Please choose an option:
echo 1. Connect to existing repository
echo 2. Create new repository instructions
echo.
set /p choice="Enter your choice (1 or 2): "

if "%choice%"=="1" (
    set /p repo_url="Enter your GitHub repository URL: "
    git remote add origin %repo_url%
    
    echo.
    echo Pushing to GitHub...
    git push -u origin master
    
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo Failed to push to GitHub. You might need to:
        echo 1. Check your internet connection
        echo 2. Verify repository URL is correct
        echo 3. Ensure you have write access to the repository
        echo 4. Set up authentication (GitHub token or SSH key)
    ) else (
        echo.
        echo Successfully pushed to GitHub!
    )
) else (
    echo.
    echo To create a new repository on GitHub:
    echo 1. Go to https://github.com/new
    echo 2. Name your repository "Excel-Automation-Tools"
    echo 3. Add a description: "A comprehensive set of Excel VBA tools for automating data processing, reporting, and validation."
    echo 4. Set repository visibility (public or private)
    echo 5. Do NOT initialize with README, .gitignore, or license
    echo 6. Click "Create repository"
    echo.
    echo After creating the repository, return here and run this script again,
    echo selecting option 1 and entering the repository URL.
)

echo.
echo Done!
pause 