import difflib

words = ["apple", "banana", "cherry", "date", "applepie", "pineapple"]
print(difflib.get_close_matches("appel", words))