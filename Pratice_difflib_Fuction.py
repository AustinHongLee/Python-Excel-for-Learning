import difflib


words = [
    "check",
    "cheese",
    "chemical",
    "chemist",
    "chemistry",
    "cherish",
    "cherry",
    "chess",
    "chew",
    "cheek"
]
print(difflib.get_close_matches("che", words))