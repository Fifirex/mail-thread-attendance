# to load new and error data into the database file

import json
filename = 'database/data.json'

with open(filename, "r") as file:
    data = json.load(file)

alias = input(" Enter the AttrEr name : ")
if (alias != ":q"):
    name = input(" Enter the actual name : ")
    data[alias] = [name, 0, False] # [name, index on row, used Flag]

# sorting block
orderedData = {}
for i in sorted(data, key=str.lower):
    orderedData[i] = data[i]

# indexing block
j = 1
for i in orderedData:
    orderedData[i][1] = j
    j+=1

with open(filename, "w") as file:
    json.dump(orderedData, file)