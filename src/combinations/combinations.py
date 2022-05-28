import sys
import logging
import numpy as np
import csv
from more_itertools import set_partitions

# Configure logs
logging.basicConfig(
    stream=sys.stderr, 
    level=logging.DEBUG,
    format='[%(asctime)s] {%(filename)s:%(lineno)d} %(levelname)s - %(message)s',
)

# Parser command line parameters
try:
    csvcities = sys.argv[1]
    csvdistance = sys.argv[2]
except IndexError:
    raise SystemExit(f"Usage: {sys.argv[0]} <cities.csv> <distance.csv>")

# The maximium number of allowed cities to generate the combinations
# due to performance issues.
MAX_CITIES = 10

def clusterization(citieslist, distance):
    distance = distance.copy()
    while len(cities) > MAX_CITIES:
        min = 999999999
        line = 0
        column = 0
        for l in range(len(cities)):
            for c in range(len(cities)):
                dist = distance[l, c]
                #print("Distância:", dist)
                if dist < min and dist != 0:
                    min = dist
                    line = l
                    column = c
        #print("Cidades:", cities, " " , len(cities))
        #print("Trash:", trash, " ", len(trash))
        #print("Clusters: ", len(citieslist))
        print("A menor distância é ", min, ", vou unir as cidades ", cities[line], " e ", cities[column])
        centrodemassa = ""
        outracidade = ""
        if trash[line] > trash [column]:
            print("A cidade ", cities[line], "é o centro de massa (", trash[line] ,") e irá representar o cluster")
            centrodemassa = cities[line]
            outracidade = cities[column]
            
            index_add = 0
            index_remove = 0
            index = 0
            for cl in citieslist:
                for ct in cl:
                    if ct == centrodemassa:
                        index_add = index
                        #cl.append(outracidade)
                    if ct == outracidade:
                        index_remove = index
                        #citieslist.pop(index)
                index = index + 1
            print(index_add)
            for a in citieslist[index_remove]:
                citieslist[index_add].append(a)
            citieslist.pop(index_remove)
            print(index_remove)
            cities.pop(column)
            trash.pop(column)
            distance = np.delete(distance, column, 0) #deleta linha
            distance = np.delete(distance, column, 1) #delete coluna
        else:
            print("A cidade ", cities[column], "é o centro de massa (", trash[column] ,") e irá representar o cluster")
            centrodemassa = cities[column]
            outracidade = cities[line]

            index_add = 0
            index_remove = 0
            index = 0
            for cl in citieslist:
                for ct in cl:
                    if ct == centrodemassa:
                        index_add = index
                        #cl.append(outracidade)
                    if ct == outracidade:
                        index_remove = index
                        #citieslist.pop(index)
                index = index + 1
            print(index_add)
            for a in citieslist[index_remove]:
                citieslist[index_add].append(a)
            citieslist.pop(index_remove)
            print(index_remove)

            cities.pop(line)
            trash.pop(line)
            distance = np.delete(distance, line, 0) #deleta linha
            distance = np.delete(distance, line, 1) #delete coluna
    return citieslist

# Read cities from CSV file
cities = list()
trash = list()
citieslist = list()
trashlist = list()

with open(csvcities, mode='r', encoding="utf8") as csv_file:
    csv_reader = csv.DictReader(csv_file, delimiter=';')
    line_count = 0
    for row in csv_reader:
        if line_count == 0:
            line_count += 1
        cities.append(row["city"])
        newcity = list()
        newcity.append(row["city"])
        citieslist.append(newcity)
        newtrash = list()
        newtrash.append(float(row["trash"]))
        trash.append(float(row["trash"]))
        trashlist.append(newtrash)
        line_count += 1

cities_temp = cities.copy()
trash_temp = trash.copy()

# Read distances from CSV file
distance = np.loadtxt(open(csvdistance, "rb"), delimiter=",", skiprows=0)

if len(citieslist) > MAX_CITIES:
    # Call clusterization to reduce the number of cities
    clusterization(citieslist, distance)
    for i in citieslist:
        print (i)

combinations = list()
combinations += list(set_partitions(citieslist))

logging.debug(len(combinations))

for i in combinations:
    print (i)
    for y in list(i):
        print(y, len(y))