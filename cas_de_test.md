# Cas de Test pour la Macro de Correction de Ponctuation

Utilisez les phrases suivantes pour tester la macro. Copiez-les dans un document LibreOffice Writer vierge et exécutez la macro `CorrectionPonctuationAvancee`.

**Contenu supposé du `DictionnaireNomsPropres.txt` :**
```
Paris
Jean
Microsoft
```

---

### Règle 1 : Ne pas mettre de point avant un nom propre

**Objectif :** Si un mot se terminant par une minuscule est suivi d'un mot qui EST dans le dictionnaire, le point doit être supprimé.

**Phrase avant correction :**
```
Le nouveau directeur. Jean est arrivé ce matin.
Nous allons visiter. Paris en avion.
Le logiciel. Microsoft est très utilisé.
```

**Résultat attendu après correction :**
```
Le nouveau directeur Jean est arrivé ce matin.
Nous allons visiter Paris en avion.
Le logiciel Microsoft est très utilisé.
```

---

### Règle 2 : Ne pas mettre de point avant une minuscule

**Objectif :** Si un mot se terminant par une minuscule est suivi d'un mot en minuscule, le point doit être supprimé.

**Phrase avant correction :**
```
C'est une belle journée. pour se promener.
Il faut finir ce travail. avant demain.
```

**Résultat attendu après correction :**
```
C'est une belle journée pour se promener.
Il faut finir ce travail avant demain.
```

---

### Règle 3 : Mettre un point avant une majuscule (qui n'est pas un nom propre)

**Objectif :** Si un mot se terminant par une minuscule est suivi d'un mot avec une majuscule qui N'EST PAS dans le dictionnaire, un point doit être ajouté.

**Phrase avant correction :**
```
Le temps est magnifique Aujourd'hui.
Il faut partir Vite.
```

**Résultat attendu après correction :**
```
Le temps est magnifique. Aujourd'hui.
Il faut partir. Vite.
```

---

### Règle 4 : Supprimer le point et le retour paragraphe avant une minuscule

**Objectif :** Si une phrase se termine par `mot.`, puis un retour à la ligne, et que la ligne suivante commence par une minuscule (et n'est pas un nom propre) ou un guillemet, le point et le retour à la ligne sont remplacés par un espace.

**Phrase avant correction :**
```
Ceci est la première partie de la phrase.
la suite est sur cette ligne.

Il a terminé son discours.
«Merci à tous» a-t-il dit.
```

**Résultat attendu après correction :**
```
Ceci est la première partie de la phrase la suite est sur cette ligne.

Il a terminé son discours «Merci à tous» a-t-il dit.
```
