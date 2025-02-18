# fuzzy regex test

import regex as re

stringA = r'(?:custard){e<=2}' # permit up to 2 errors

# {i,d,s} can also be specified separately -> insertions, deletions, substitutions

stringB = 'c ustard'
stringC = 'c8stard'
stringD = 'custardx'
stringE = 'cusard'
stringF = 'cuuusard' # two insertions and 1 deletion

if re.match(stringA,stringB):
    print('StringB was a match')
else:
    print('Not sure what went wrong')

if re.match(stringA,stringF):
    print('StringF was a match')
else:
    print('This is actually the correct answer, it is not enough of a match according to my rules')