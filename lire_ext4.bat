@echo off
:: Utilisation de l'encodage UTF-8 pour supporter les accents français
chcp 65001 >nul
title Lecteur de Disques EXT4 pour Windows 11
color 0A

:: Vérification des privilèges Administrateur (requis pour wsl --mount)
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo =====================================================================
    echo    CE SCRIPT DOIT ÊTRE EXÉCUTÉ EN TANT QU'ADMINISTRATEUR
    echo =====================================================================
    echo.
    echo Demande d'élévation des privilèges en cours...
    echo Veuillez accepter l'invite de commande de l'UAC si elle apparaît.
    echo.
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:menu
cls
echo =====================================================================
echo           LECTEUR DE DISQUES EXT4 / LINUX SUR WINDOWS 11
echo =====================================================================
echo.
echo Ce script utilise WSL 2 (Windows Subsystem for Linux) pour monter
echo et lire nativement des disques formatés en EXT4 (USB ou autre).
echo.
echo [1] Lister les disques physiques et monter un disque/partition
echo [2] Démonter un disque physique proprement
echo [3] Afficher les dossiers actuellement montés dans WSL
echo [4] Ouvrir le dossier de montage dans l'Explorateur Windows
echo [5] Quitter
echo.
echo =====================================================================
set /p choix="Veuillez choisir une option [1-5] : "

if "%choix%"=="1" goto lister_monter
if "%choix%"=="2" goto demonter
if "%choix%"=="3" goto lister_montages
if "%choix%"=="4" goto ouvrir_explorateur
if "%choix%"=="5" goto quitter
goto menu


:lister_monter
cls
echo =====================================================================
echo               LISTE DES DISQUES PHYSIQUES CONNECTÉS
echo =====================================================================
echo.
:: Vérification de la présence de WSL
where wsl >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERREUR] WSL (Windows Subsystem for Linux) n'est pas installé sur ce système.
    echo Veuillez l'installer en exécutant la commande : wsl --install
    echo Puis redémarrez votre ordinateur.
    echo.
    pause
    goto menu
)

:: Affichage de la liste des disques connectés en utilisant PowerShell
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_DiskDrive | Select-Object @{Name='Numero_Disque';Expression={[int]($_.DeviceID -replace '[^\d]+', '')}}, Model, @{Name='Taille_Go';Expression={[math]::round($_.Size/1GB,1)}}, InterfaceType | Format-Table -AutoSize"
echo.
echo Remarque : Repérez le numéro du disque correspondant à votre SSD/clé USB EXT4.
echo.
set /p disk_num="Entrez le numéro du disque à monter (ex: 1) : "
if "%disk_num%"=="" (
    echo [ERREUR] Numéro de disque invalide.
    pause
    goto lister_monter
)

echo.
echo Optionnel : Si le disque possède plusieurs partitions, spécifiez l'index (ex: 1).
echo Laissez vide pour monter le disque entier ou si vous ne savez pas.
set /p part_num="Entrez le numéro de la partition (laisser vide par défaut) : "

echo.
echo Tentative de montage du disque %disk_num%...
if "%part_num%"=="" (
    echo Commande : wsl --mount \\.\PHYSICALDRIVE%disk_num% --type ext4
    wsl --mount \\.\PHYSICALDRIVE%disk_num% --type ext4
) else (
    echo Commande : wsl --mount \\.\PHYSICALDRIVE%disk_num% --partition %part_num% --type ext4
    wsl --mount \\.\PHYSICALDRIVE%disk_num% --partition %part_num% --type ext4
)

if %errorlevel% neq 0 (
    echo.
    echo [ÉCHEC] Impossible de monter le disque.
    echo Raisons possibles :
    echo   1. Le disque est déjà monté.
    echo   2. Le numéro de disque ou de partition est incorrect.
    echo   3. Le système de fichiers n'est pas de l'EXT4.
    echo   4. WSL n'est pas démarré (lancez d'abord un terminal Linux ou attendez quelques secondes).
) else (
    echo.
    echo [SUCCÈS] Le disque a été monté avec succès !
    echo.
    echo Tentative d'ouverture de l'Explorateur Windows...
    :: Utilisation de WSL pour ouvrir le dossier de montage dans l'Explorateur Windows
    if "%part_num%"=="" (
        wsl.exe -e sh -c "if [ -d '/mnt/wsl/PHYSICALDRIVE%disk_num%' ]; then cd '/mnt/wsl/PHYSICALDRIVE%disk_num%' && explorer.exe .; else cd /mnt/wsl && explorer.exe .; fi"
    ) else (
        wsl.exe -e sh -c "if [ -d '/mnt/wsl/PHYSICALDRIVE%disk_num%p%part_num%' ]; then cd '/mnt/wsl/PHYSICALDRIVE%disk_num%p%part_num%' && explorer.exe .; elif [ -d '/mnt/wsl/PHYSICALDRIVE%disk_num%' ]; then cd '/mnt/wsl/PHYSICALDRIVE%disk_num%' && explorer.exe .; else cd /mnt/wsl && explorer.exe .; fi"
    )
)
echo.
pause
goto menu


:demonter
cls
echo =====================================================================
echo                     DÉMONTAGE D'UN DISQUE PHYSIQUE
echo =====================================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_DiskDrive | Select-Object @{Name='Numero_Disque';Expression={[int]($_.DeviceID -replace '[^\d]+', '')}}, Model, @{Name='Taille_Go';Expression={[math]::round($_.Size/1GB,1)}}, InterfaceType | Format-Table -AutoSize"
echo.
set /p disk_num="Entrez le numéro du disque à démonter (ex: 1) : "
if "%disk_num%"=="" (
    echo [ERREUR] Numéro de disque invalide.
    pause
    goto demonter
)

echo.
echo Démontage du disque %disk_num%...
wsl --unmount \\.\PHYSICALDRIVE%disk_num%

if %errorlevel% neq 0 (
    echo.
    echo [ÉCHEC] Impossible de démonter le disque. Il se peut qu'il ne soit pas monté.
) else (
    echo.
    echo [SUCCÈS] Le disque a été démonté proprement. Vous pouvez le déconnecter en toute sécurité.
)
echo.
pause
goto menu


:lister_montages
cls
echo =====================================================================
echo             DOSSIERS ACTUELLEMENT MONTÉS DANS WSL
echo =====================================================================
echo.
where wsl >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERREUR] WSL n'est pas disponible.
    pause
    goto menu
)

echo Voici les répertoires présents sous /mnt/wsl :
echo ----------------------------------------------------
wsl.exe ls -lh /mnt/wsl 2>nul
if %errorlevel% neq 0 (
    echo Aucun dossier de montage actif trouvé sous /mnt/wsl.
)
echo ----------------------------------------------------
echo.
pause
goto menu


:ouvrir_explorateur
cls
echo =====================================================================
echo             OUVERTURE DE L'EXPLORATEUR DE FICHIERS
echo =====================================================================
echo.
echo Ouverture du dossier général de montage dans l'Explorateur Windows...
explorer.exe \\wsl.localhost\
echo.
echo Si l'Explorateur ne s'ouvre pas automatiquement, saisissez
echo '\\wsl.localhost\' dans la barre d'adresse de votre explorateur.
echo.
pause
goto menu


:quitter
cls
echo Merci d'avoir utilisé le Lecteur de Disques EXT4.
echo À bientôt !
timeout /t 3 >nul
exit /b
