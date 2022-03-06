from ast import Sub
from random import randint, sample
from matplotlib import style
import pandas as pd
from docx import Document
from docx.shared import Inches
from itertools import combinations



letters = ['a', 'b', 'c', 'd', 'e']
exercises = [[1, 2, 3], [4, 5, 6], [7, 3, 5], [1, 2, 7], [4, 6, 1]]
document = Document()
subsets = []
df = ''
    
def gen_problem(mod, sec, exercises):
    mods = df.loc[(df['Mod'] == mod)]
    secs = mods.loc[(df['Sec'] == sec)]
    problems = secs.loc[(df['Type'] == 'Opening')]
    problem = problems['Original'].iloc[0]
    document.add_paragraph(
        problem, style='List Number'
    )

    sub_sentences = secs.loc[(df['Type'] == 'Sentence')]

    for index in exercises:
        option = sub_sentences.loc[(df['Exercise'] == index)]['Original'].iloc[0]
        p1 = document.add_paragraph(f"{getRoman(index)}. {index}-{option}")
        p1.paragraph_format.left_indent = Inches(.25)
    p = document.add_paragraph('São corretas apenas as afirmativas:')

    answer()

    document.add_paragraph('Gabarito: ')


    keys = secs.loc[(df['Type'] == 'Key')]
    explanations = secs.loc[(df['Type'] == 'Explanation')]
    for index in range(1, 4):
        key = keys.loc[(df['Exercise'] == index)]['Original'].iloc[0]
        explanation = explanations.loc[(df['Exercise'] == index)]['Original'].iloc[0]
        p2 = document.add_paragraph(f"{getRoman(index)}. é {key} porque {explanation}")
        p2.paragraph_format.left_indent = Inches(.25)

    document.add_paragraph(f'Referência: Módulo {mod}, Seção {sec}')

def getSubsets():   
    stuff = [1, 2, 3]
    for L in range(0, len(stuff)+1):
        for subset in combinations(stuff, L):
            subsets.append(subset)
    
def answer():
    indexs = sample(range(0, 8), 5)
    for i in range(0, 5):
        solution = solution_str(subsets[indexs[i]])
        p3 = document.add_paragraph(f"{letters[i]}) {solution}")
        p3.paragraph_format.left_indent = Inches(.25)

def solution_str(tuple):
    n = len(tuple)
    result = ''

    if n == 0:
        result = 'Nenhuma delas'
    else:
        seperator = ''
        for i in range(0, n):
            if i == n-1:
                seperator = ' e '
            else: 
                seperator = ', '
            result = result + seperator + getRoman(tuple[i])
        result = result[2:]
    return result

def getRoman(number):
	num = [1, 4, 5, 9, 10, 40, 50, 90, 100, 400, 500, 900, 1000]
	sym = ["I", "IV", "V", "IX", "X", "XL", "L", "XC", "C", "CD", "D", "CM", "M"]
	i = 12
	result = ''
	while number:
		div = number // num[i]
		number %= num[i]

		while div:
			result = result + sym[i]
			div -= 1
		i -= 1
	return result

if __name__ == "__main__":
    # Data Extraction
    df = pd.read_excel('Input.xlsx')
    df.dropna(axis=1, how='all', inplace=True)
    df.dropna(axis=0, how='all', inplace=True)
    df.rename(columns={'Unnamed: 1': 'Mod', 'Unnamed: 2': 'Sec', 'Unnamed: 3': 'Exercise', 'Unnamed: 4': 'Type', 'Unnamed: 5': 'Original', 'Unnamed: 6': 'Mirror'}, inplace=True)
    df = df.iloc[1:]

    # Answer subsets
    getSubsets()

    # produce docx
    for mod in range(1, 5):
        for sec in range(1, 3):
            for exercise in exercises:
                gen_problem(mod, sec, exercise)

    document.save('demo.docx')