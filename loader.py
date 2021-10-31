# to load new and error data into the database file
import json
filename = 'database/name-map.json'

with open(filename, "r") as file:
    data = json.load(file)

alias = input(" Enter the AttrEr name : ")
if (alias != ":q"):
    name = input(" Enter the actual name : ")
    data[alias] = [name, False] # [name, used Flag]

with open(filename, "w") as file:
    json.dump(data, file)