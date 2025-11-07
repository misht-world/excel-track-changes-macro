@echo off
set /p repo_name="Enter repository name: "
set /p description="Enter repository description: "

mkdir %repo_name%
cd %repo_name%

git init
echo # %repo_name% > README.md
echo. >> README.md
echo %description% >> README.md

:: Стандартный .gitignore для VBA/Excel
echo # VBA and Excel files > .gitignore
echo *.xls >> .gitignore
echo *.xlsx >> .gitignore
echo *.xlsm >> .gitignore
echo *.xlam >> .gitignore
echo *.vbw >> .gitignore
echo. >> .gitignore
echo # Temporary files >> .gitignore
echo ~$* >> .gitignore

git add .
git commit -m "Initial commit"

echo.
echo Local repository created!
echo Now go to GitHub.com and create a new repository named: %repo_name%
echo Then run: git remote add origin https://github.com/your_username/%repo_name%.git
echo And: git push -u origin main
pause