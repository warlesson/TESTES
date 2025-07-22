@echo off
cd /d %USERPROFILE%\Desktop\AUTOMAÇÃO

echo Adicionando arquivos ao Git...
git add .

echo Fazendo commit...
git commit -m "Atualização automática"

echo Enviando para o GitHub...
git push

pause
