import DataCleanser as DC
from datetime import date

#print(DC.newDateColumnGenerator(10, date(1800, 1, 1), date(1950, 1, 1), "%d/%m/%Y"))

print(DC.getDateBoundariesInputFromUser('test'))

#print(DC.getDateBoundariesInputFromUser('test', True))

#print(DC.getDateBoundariesInputFromUser('test', True, date('1450')))      