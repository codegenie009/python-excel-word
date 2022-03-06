from random import sample
from itertools import combinations, chain

stuff = [1, 2, 3]
subsets = []
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

turple = (1, 2)
n = len(turple)
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
		result = result + seperator + getRoman(turple[i])
	result = result[2:]
print(result)








