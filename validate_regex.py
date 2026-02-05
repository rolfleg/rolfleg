import re

# Mock dictionary
PROPER_NOUNS = {"Paris", "Marie", "Jean"}

def is_lowercase(s):
    if not s: return False
    return s.islower()

def is_in_dictionary(word):
    return word in PROPER_NOUNS

# Rules 1-3
# We search for the end of the first word and look ahead for the second.
PATTERN_123_PART1 = r"([a-zà-ÿ])(\.?)(\s+)"

def apply_rules_123(text):
    # Process from end to start to maintain indices
    matches = list(re.finditer(PATTERN_123_PART1, text))
    new_text = text
    for m in reversed(matches):
        g1, g2, g3 = m.groups()
        # Avoid matching across paragraphs for Rules 1-3
        if "\n" in g3: continue

        start, end = m.span()
        next_part = new_text[end:]
        word_match = re.match(r"([A-Za-zÀ-ÿ«]+)", next_part)
        if word_match:
            g4 = word_match.group(1)
            first_char = g4[0]

            sNew = None
            if is_lowercase(first_char):
                # Rule 2: lowercase -> one space, no period
                sNew = g1 + " "
            elif is_in_dictionary(g4):
                # Rule 1: proper noun -> one space, no period
                sNew = g1 + " "
            else:
                # Rule 3: uppercase not in dico -> add period, keep space
                sNew = g1 + "." + g3

            if sNew is not None and sNew != (g1 + g2 + g3):
                new_text = new_text[:start] + sNew + new_text[end:]
    return new_text

# Rule 4
PATTERN_4 = r"([a-zà-ÿ])(\.?)(\s*)\n"

def apply_rule_4(text):
    matches = list(re.finditer(PATTERN_4, text))
    new_text = text
    for m in reversed(matches):
        g1, g2, g3 = m.groups()
        start, end = m.span()
        next_part = new_text[end:]
        # Get first word of next paragraph
        word_match = re.match(r"([A-Za-zÀ-ÿ«]+)", next_part)
        if word_match:
            g4 = word_match.group(1)
            # Simplification: word for dico check
            word_for_dico = re.sub(r"[«]", "", g4)
            if not is_in_dictionary(word_for_dico) and (is_lowercase(g4[0]) or g4[0] == "«"):
                # Rule 4: merge
                new_text = new_text[:start] + g1 + " " + new_text[end:]
    return new_text

# Test Cases
tests = [
    # Rule 1
    ("le chat. Paris", "le chat Paris"),
    ("le chat  Paris", "le chat Paris"),
    # Rule 2
    ("le chat. dort", "le chat dort"),
    ("le chat  dort", "le chat dort"),
    # Rule 3
    ("le chat Dort", "le chat. Dort"),
    ("le chat. Dort", "le chat. Dort"),
    ("le chat  Dort", "le chat.  Dort"),
    # Rule 4
    ("Fin.\nsuivant", "Fin suivant"),
    ("Fin\nsuivant", "Fin suivant"),
    ("Fin.\n« début", "Fin « début"),
    ("Fin.\nParis", "Fin.\nParis"),
]

print("Running Regex Validation Tests...")
success_count = 0
for i, (inp, expected) in enumerate(tests):
    res1 = apply_rules_123(inp)
    result = apply_rule_4(res1)

    if result == expected:
        print(f"Test {i+1}: PASS")
        success_count += 1
    else:
        print(f"Test {i+1}: FAIL")
        print(f"  Input:    {repr(inp)}")
        print(f"  Expected: {repr(expected)}")
        print(f"  Actual:   {repr(result)}")

print(f"\nSummary: {success_count}/{len(tests)} tests passed.")
import sys
sys.exit(0 if success_count == len(tests) else 1)
