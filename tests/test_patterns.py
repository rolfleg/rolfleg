import re

def apply_rules(text):
    # En Python, on utilise re.UNICODE par défaut en 3.x, mais on définit les plages explicitement
    low = "a-zàâäæçéèêëîïôöœùûüÿ"
    cap = "A-ZÀ-ÂÄÆÇÉÈÊËÎÏÔÖŒÙÛÜŸ"

    # 1. Nettoyage des espaces autour des sauts de ligne
    text = re.sub(r'[ \t]+\n', '\n', text)
    text = re.sub(r'\n[ \t]+', '\n', text)

    # 2. Règle 6 : minuscule + \n + minuscule -> espace
    text = re.sub(f'([{low}])\n([{low}])', r'\1 \2', text)

    # 3. Règle 7 : minuscule + \n + Majuscule -> ". "
    text = re.sub(f'([{low}])\n([{cap}])', r'\1. \2', text)

    # 4. Règle 4 : minuscule + "." + \n + (minuscule ou «) -> minuscule + " " + (minuscule ou «)
    # Note: re.sub utilise \. pour le point. Dans le code VBA on a utilisé [.]
    text = re.sub(f'([{low}])\.\n([{low}«])', r'\1 \2', text)

    # 5. Nettoyage des doubles espaces
    text = re.sub(r' +', ' ', text)

    return text

def test():
    test_cases = [
        ("un\nchat", "un chat"),                         # Règle 6
        ("il fait beau\naujourd'hui", "il fait beau aujourd'hui"),
        ("été\nincertain", "été incertain"),             # Règle 6 avec accent
        ("fin\nNouveau", "fin. Nouveau"),                # Règle 7
        ("poursuite\nLa suite", "poursuite. La suite"),  # Règle 7
        ("mot.\nsuite", "mot suite"),                    # Règle 4
        ("mot.\n« début", "mot « début"),                # Règle 4 avec guillemet
        ("espace \naprès", "espace après"),              # Nettoyage + Règle 6
        ("déjà.\n«", "déjà «"),                          # Règle 4 avec accent
        ("Une phrase.\nUne autre.", "Une phrase.\nUne autre."), # Ne doit pas toucher (Maj + . + \n + Maj)
    ]

    for inp, expected in test_cases:
        actual = apply_rules(inp)
        if actual == expected:
            print(f"OK: {repr(inp)} -> {repr(actual)}")
        else:
            print(f"FAIL: {repr(inp)} -> {repr(actual)}, expected {repr(expected)}")
            exit(1)

if __name__ == "__main__":
    test()
    print("Tous les tests ont réussi !")
