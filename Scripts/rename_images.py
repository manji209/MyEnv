import os
import re

dir = "C:/Users/Dinh/Desktop/Website_Pictures/Long_Names"
str = '15'
'''
str2 = '5'
temps = 'thie15lslsl15lalal15'

indices2 = [i for i, a in enumerate(temps) if a == str2]


for filename in os.listdir(dir):
    occurrences = filename.count(str)
    indices = [i for i, a in enumerate(filename) if a == str]
    print(indices)
    if len(indices) == 1 and len(filename) > 13:
        index_2 = filename.index('.')
        newstr = filename[:indices[0]] + filename[index_2:]
        os.rename(os.path.join(dir, filename),
                  os.path.join(dir, newstr))
    elif len(indices) == 2 and len(filename) > 13:
        index_2 = filename.index('.')
        newstr = filename[:indices[1]] + filename[index_2:]
        os.rename(os.path.join(dir, filename),
                  os.path.join(dir, newstr))
    elif len(indices) == 3 and len(filename) > 13:
        index_2 = filename.index('.')
        newstr = filename[:indices[2]] + filename[index_2:]
        os.rename(os.path.join(dir, filename),
                  os.path.join(dir, newstr))


print('This is 2nd:', indices2)
'''

for filename in os.listdir(dir):
    indices = []
    for match in re.finditer(str, filename):
        indices.append(match.start())
    try:
        if len(indices) == 1 and len(filename) > 13:
            index_2 = filename.index('.')
            newstr = filename[:indices[0]] + filename[index_2:]
            os.rename(os.path.join(dir, filename),
                    os.path.join(dir, newstr))
        elif len(indices) == 2 and len(filename) > 13:
            index_2 = filename.index('.')
            newstr = filename[:indices[1]] + filename[index_2:]
            os.rename(os.path.join(dir, filename),
                    os.path.join(dir, newstr))
        elif len(indices) == 3 and len(filename) > 13:
            index_2 = filename.index('.')
            newstr = filename[:indices[2]] + filename[index_2:]
            os.rename(os.path.join(dir, filename),
                    os.path.join(dir, newstr))
    except:
        continue


