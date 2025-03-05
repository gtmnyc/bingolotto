@echo off
cd /d "C:\Users\gerry\Desktop\Scripts\kivy project"

:: Set GitHub repository details
set GITHUB_USERNAME=gtmnyc
set GITHUB_REPO=bingolotto

:: Initialize Git (if not already initialized)
git init

:: Add files
git add .

:: Commit the changes
git commit -m "Initial commit for Kivy project"

:: Set remote repository (change to your GitHub repo URL)
git remote add origin https://github.com/gtmnyc/bingolotto

:: Push the project
git branch -M main
git push -u origin main

echo Git push completed! Your project is now on GitHub.
pause
