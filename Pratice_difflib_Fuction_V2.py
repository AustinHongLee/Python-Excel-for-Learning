# coding: utf-8
import difflib
import re


word_list = ["Pipe", "Valve", "Bolt", "Nut", "Elbow", "Flange"]

def closest_translation(target_word, reference_list):
    # 找到最接近的關鍵字
    closest_word = difflib.get_close_matches(target_word, reference_list, n=1, cutoff=0.7)
    if closest_word:
        return closest_word[0]
    else:
        return target_word

input_str = "Pipe, ASME B36.10; Beveled End | ASTM A53-B, Electric Resistance Welded (Ej =0.85); SCH/THK S-STD"
words = re.split(',|;|\|', input_str)
translated_words = [closest_translation(word.strip(), word_list) for word in words]
print(translated_words)


'''
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
dilllib.get_close_matches("關鍵字",變數通常為一個要給關鍵字判讀的列表)

'''
