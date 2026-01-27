# Instructions d'Installation de la Macro

Suivez ces étapes pour installer la macro de correction de ponctuation dans votre LibreOffice Writer.

### Étape 1 : Ouvrir l'éditeur de macros

1.  Ouvrez un nouveau document dans **LibreOffice Writer**.
2.  Allez dans le menu en haut de l'écran : `Outils` -> `Macros` -> `Gérer les macros` -> `LibreOffice Basic...`.
    ![Menu Macros](https://i.imgur.com/8Q2gK4h.png)

### Étape 2 : Créer un nouveau module

1.  Une fenêtre intitulée **Macros LibreOffice Basic** va s'ouvrir.
2.  Dans la partie gauche (`Macro de`), cliquez sur **Mes macros** pour la développer, puis sur **Standard**. `Standard` est la bibliothèque par défaut où sont stockées les macros personnelles.
3.  Cliquez sur le bouton **Nouveau**.
    ![Nouveau Module](https://i.imgur.com/u5V3gYx.png)
4.  LibreOffice vous demandera de donner un nom au nouveau module. Vous pouvez l'appeler `CorrectionPonctuation`, par exemple. Cliquez ensuite sur **OK**.

### Étape 3 : Insérer le code de la macro

1.  Une nouvelle fenêtre s'ouvrira, c'est l'éditeur de code (l'EDI de LibreOffice Basic). Il contiendra un code par défaut comme celui-ci :
    ```basic
    REM  *****  BASIC  *****

    Sub Main

    End Sub
    ```
2.  **Supprimez entièrement** ce code par défaut.
3.  Ouvrez le fichier `CorrectionPonctuationAvancee.bas` que je vous ai fourni.
4.  **Copiez tout le contenu** de ce fichier.
5.  **Collez-le** dans la fenêtre de l'éditeur de macros que vous venez de vider.
6.  Le code complet, commençant par `REM ***** BASIC *****` et se terminant par `End Function`, doit maintenant remplir la fenêtre.

### Étape 4 : Enregistrer et fermer

1.  Cliquez sur l'icône de disquette (`Enregistrer`) en haut de l'éditeur de macros, ou appuyez sur `Ctrl + S` pour sauvegarder le code.
2.  Vous pouvez maintenant fermer la fenêtre de l'éditeur de macros.

**Votre macro est maintenant installée et prête à être utilisée !**

Vous pouvez l'exécuter en retournant dans `Outils` -> `Macros` -> `Exécuter la macro...`, puis en naviguant jusqu'à `Mes macros` -> `Standard` -> `CorrectionPonctuation` et en sélectionnant `CorrectionPonctuationAvancee`.
