# Lecteur de Disques EXT4 pour Windows 11 (SSD externe connecté en USB)

Ce dépôt contient un script de commande Windows Batch (`.bat`) conçu pour faciliter la lecture et le montage de disques ou partitions formatés en **EXT4** (et autres systèmes de fichiers Linux) directement depuis Windows 11.

## Contexte et Fonctionnement
Windows 11 ne supporte pas nativement les systèmes de fichiers Linux comme l'EXT4, affichant souvent ces disques comme « RAW » et proposant à tort de les formater.

Pour résoudre ce problème de manière sécurisée et gratuite, ce script utilise la fonctionnalité native `wsl --mount` de **WSL 2** (Windows Subsystem for Linux 2). Il permet d'attacher un disque physique USB (comme un SSD externe ou une clé USB) directement dans le noyau Linux virtuel de Windows, rendant ainsi les fichiers lisibles et modifiables depuis l'Explorateur de fichiers Windows à l'adresse réseau `\\wsl.localhost\`.

---

## Prérequis
1. **Windows 11** (Famille, Professionnel ou Éducation).
2. **WSL 2** installé sur votre machine.
   - Si WSL n'est pas installé, ouvrez un terminal en administrateur et tapez : `wsl --install`.
   - Redémarrez ensuite votre ordinateur.
3. Le script doit impérativement être exécuté avec les **privilèges Administrateur** (le script demandera automatiquement l'élévation si nécessaire).

---

## Fonctionnalités du Script (`lire_ext4.bat`)
Le script est entièrement interactif et en français. Il propose un menu simple :

1. **Lister les disques physiques connectés** : Affiche les informations essentielles (Numéro de disque, Modèle, Taille en Go, Type d'interface USB/SATA).
2. **Monter un disque ou une partition spécifique** : Permet de monter le disque choisi (par exemple le disque `1` ou la partition `1` du disque `1`) avec la commande appropriée.
3. **Ouverture automatique de l'Explorateur Windows** : Une fois le montage réussi, l'Explorateur s'ouvre directement sur le dossier réseau contenant vos fichiers Linux.
4. **Afficher les dossiers actuellement montés** : Pour voir rapidement quels disques sont actifs dans WSL.
5. **Démonter proprement** : Permet de démonter le disque physique avant de le débrancher pour éviter toute corruption de données.

---

## Comment l'utiliser ?
1. Branchez votre SSD externe EXT4 en USB.
2. Double-cliquez sur `lire_ext4.bat` (ou faites un clic droit -> *Exécuter en tant qu'administrateur*).
3. Choisissez l'option **1** pour lister vos disques et monter votre SSD.
4. Une fois votre travail terminé, retournez sur le script et choisissez l'option **2** pour démonter le disque de manière sécurisée avant de le déconnecter physiquement.
