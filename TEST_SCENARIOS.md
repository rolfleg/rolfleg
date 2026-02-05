# Scénarios de test pour la macro de ponctuation

Ce document répertorie les cas de test pour vérifier la logique de la macro PunctuationMacro.bas.

## Dictionnaire des noms propres (Exemple)
- Paris
- Marie
- Jean

## Règle 1 : Minuscule + [.] + Espace + Nom Propre
**Action attendue :** Supprimer le point si présent, un seul espace.

| Entrée | Sortie attendue | Commentaire |
| :--- | :--- | :--- |
| `le chat. Paris` | `le chat Paris` | Paris est dans le dictionnaire. |
| `le chat  Paris` | `le chat Paris` | Normalisation des espaces. |
| `le chat. Marie` | `le chat Marie` | Marie est dans le dictionnaire. |

## Règle 2 : Minuscule + [.] + Espace + Minuscule
**Action attendue :** Supprimer le point si présent, un seul espace.

| Entrée | Sortie attendue | Commentaire |
| :--- | :--- | :--- |
| `le chat. dort` | `le chat dort` | 'dort' commence par une minuscule. |
| `le chat dort` | `le chat dort` | Déjà correct, un seul espace. |
| `le chat  dort` | `le chat dort` | Normalisation des espaces. |

## Règle 3 : Minuscule + [Espace] + Majuscule (Hors Dico)
**Action attendue :** Ajouter un point, garder l'espace existant.

| Entrée | Sortie attendue | Commentaire |
| :--- | :--- | :--- |
| `le chat Dort` | `le chat. Dort` | 'Dort' n'est pas dans le dictionnaire. |
| `le chat. Dort` | `le chat. Dort` | Déjà un point, pas de changement. |
| `le chat  Dort` | `le chat.  Dort` | Garde l'espace double. |

## Règle 4 : Minuscule + [.] + Paragraphe + (Minuscule ou «) (Hors Dico)
**Action attendue :** Supprimer le point, supprimer le retour paragraphe, un seul espace.

| Entrée | Sortie attendue | Commentaire |
| :--- | :--- | :--- |
| `Fin.`<br>`suivant` | `Fin suivant` | Fusion de paragraphes. |
| `Fin`<br>`suivant` | `Fin suivant` | Fusion même sans point. |
| `Fin.`<br>`« début` | `Fin « début` | '«' déclenche la fusion. |
| `Fin.`<br>`Paris` | `Fin.`<br>`Paris` | Paris est dans le dico, pas de fusion. |
