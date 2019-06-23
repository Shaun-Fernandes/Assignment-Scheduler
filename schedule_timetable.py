import random
from time import time
from statistics import mean
from random import shuffle
import itertools

def checkDupCols(arr):
    for row in arr:
        seen = set()
        for x in row:
            if x in seen:
                return True
            if x is not None:
                seen.add(x)
    return False


names = [   [1, 2, 3, 4, 5],
            [1, 2, 3, 4, 7, 9],
            [1, 2, 3, 4, 6],
            [1, 2, 3, 4, 8, 11],    ]
newNames = [list(i) for i in itertools.zip_longest(*names[:4])]
randNo = 0

for x in names: print(x)
for x in newNames: print(x)
randNo


randList = []
timeList = []
maxTimeLimit = 0.05

for _ in range(10):
    timeExceded = True
    while timeExceded:
        timeExceded = False
        startTime = time()
        names = [   ['MM','GJ','KT','NO','CT','MB','SS','AS','EN'],
                    ['MB','GT','AS','EN','SS'],
                    ['MB','GT','AS','EN','SS'],
                    ['KT','MM','GT','SS','GJ','NO','MB','CT'],
                    ['GJ'],
                    ['MM','GJ','KT','NO','CT'],
                    ['MW','GT','CT','KT','TF','SS','JK','MB','EN'],
                    ['MB','TF','SS','MW','SF','EN','CS','JN','JK']    ]
        randNo = 0
        shuffle(names[0])
        for i in range(len(names)):
            newNames = [list(x) for x in itertools.zip_longest(*names[:i+1])]
            while checkDupCols(newNames) and done:
                shuffle(names[i])
                newNames = [list(x) for x in itertools.zip_longest(*names[:i+1])]
                randNo += 1
                if( time()-startTime > maxTimeLimit):
                    print("Time taken for iteration", _, "was  = ", time()-startTime,)
                    print("Restarting iteration")
                    timeExceded = True
                    break
            if timeExceded:
                break

        if not timeExceded:
            print("_ = ", _, "number of randomizes = ", randNo)
            print("time = ", time()-startTime)
            timeList.append(time()-startTime)
            randList.append(randNo)

mean (randList)
mean (timeList)
len(randList)
len(timeList)


for x in names: print(x)
for x in newNames: print(x)
randNo



i = 2
for x in names: print(x)
newNames = [list(x) for x in itertools.zip_longest(*names[:i+1])]
newNames
for x in newNames: print(x)
checkDupCols(newNames)

shuffle(names[3])
