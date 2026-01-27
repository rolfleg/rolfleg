# Instructions pour le Dictionnaire et les Tests

Ces instructions complètent le guide d'installation et vous expliquent comment gérer le dictionnaire de noms propres et comment vérifier le bon fonctionnement de la macro.

### Étape 1 : Créer et placer le fichier `DictionnaireNomsPropres.txt`

L'efficacité de la macro repose sur un dictionnaire de noms propres que vous devez créer. La macro est conçue pour le trouver et même le créer pour vous si elle ne le trouve pas, mais il est préférable de savoir où il se trouve.

1.  **Ouvrez l'Explorateur de fichiers de Windows.**
2.  Dans la barre d'adresse, copiez-collez le chemin suivant et appuyez sur `Entrée` :
    ```
    %APPDATA%\LibreOffice\4\user\Scripts\python\
    ```
    *   **Note :** Ce chemin est celui que vous avez spécifié dans votre demande. Si le dossier `python` n'existe pas dans `Scripts`, vous pouvez le créer manuellement.

3.  **Créez le fichier dictionnaire :**
    *   Faites un clic droit dans le dossier -> `Nouveau` -> `Document texte`.
    *   Nommez le fichier exactement `DictionnaireNomsPropres.txt`. Assurez-vous qu'il ne se termine pas par `.txt.txt`.

4.  **Remplissez le dictionnaire :**
    *   Ouvrez le fichier `DictionnaireNomsPropres.txt`.
    *   **Ajoutez un nom propre par ligne.** La casse est importante (ex: "Paris", pas "paris").
    *   Vous pouvez commencer par copier-coller le contenu du fichier `DictionnaireNomsPropres.txt` que je vous ai fourni :
        ```
        Paris
        Jean
        Microsoft
        ```
    *   Enregistrez et fermez le fichier.

### Étape 2 : Tester la macro avec les phrases d'exemple

1.  **Ouvrez un nouveau document vierge dans LibreOffice Writer.**
2.  Ouvrez le fichier `cas_de_test.md` que je vous ai fourni.
3.  **Testez chaque règle une par une :**
    *   Copiez une section "Phrase avant correction" du fichier de test.
    *   Collez-la dans votre document Writer.
    *   Exécutez la macro en allant dans `Outils` -> `Macros` -> `Exécuter la macro...`.
    *   Naviguez jusqu'à `Mes macros` -> `Standard` -> `CorrectionPonctuation` et sélectionnez `CorrectionPonctuationAvancee`, puis cliquez sur **Exécuter**.
4.  **Vérifiez le résultat :**
    *   Le texte dans votre document doit maintenant correspondre au bloc "Résultat attendu après correction" du fichier de test.
    *   Répétez ce processus pour les quatre règles afin de vous assurer que tout fonctionne parfaitement.

---

**Félicitations !** Votre macro de correction avancée est maintenant entièrement configurée et testée. Vous pouvez l'utiliser sur vos documents.
