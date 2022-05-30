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

# RSU
rsutrash = [0,25,75,150,250,350,700,1250,2500,5000]
capexRT1 = [0,
10952,
4689,
3061,
2346,
2051,
1638,
1576,
1394,
1359
]
capexRT2 = [0,
9861,
3813,
2305,
1701,
1483,
1174,
1071,
956,
852
]
capexRT3 = [0,
10998,
4894,
3365,
2848,
2651,
1802,
1668,
1600,
1530
]
capexRT4 = [0,
12934,
7226,
5724,
5037,
4674,
3947,
4031,
3570,
3589
]
opexRT1 = [0,
1886,
707,
389,
269,
234,
171,
142,
126,
105
]
opexRT2 = [0,
1557,
568,
327,
229,
196,
144,
116,
105,
86
]
opexRT3 = [0,
1665,
626,
364,
261,
257,
167,
139,
126,
107
]
opexRT4 = [0,
2828,
1049,
603,
423,
344,
240,
207,
179,
165
]


def clusterization(citieslist, distance):
    distance = distance.copy()
    while len(cities_temp) > MAX_CITIES:
        min = 999999999
        line = 0
        column = 0
        for l in range(len(cities_temp)):
            for c in range(len(cities_temp)):
                dist = distance[l, c]
                #print("Distância:", dist)
                if dist < min and dist != 0:
                    min = dist
                    line = l
                    column = c
        #print("Cidades:", cities, " " , len(cities))
        #print("Trash:", trash, " ", len(trash))
        #print("Clusters: ", len(citieslist))
        print("A menor distância é ", min, ", vou unir as cidades ", cities_temp[line], " e ", cities_temp[column])
        centrodemassa = ""
        outracidade = ""
        if trash_temp[line] > trash_temp [column]:
            print("A cidade ", cities_temp[line], "é o centro de massa (", trash_temp[line] ,") e irá representar o cluster")
            centrodemassa = cities_temp[line]
            outracidade = cities_temp[column]
            
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
            cities_temp.pop(column)
            trash_temp.pop(column)
            distance = np.delete(distance, column, 0) #deleta linha
            distance = np.delete(distance, column, 1) #delete coluna
        else:
            print("A cidade ", cities_temp[column], "é o centro de massa (", trash_temp[column] ,") e irá representar o cluster")
            centrodemassa = cities_temp[column]
            outracidade = cities_temp[line]

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

            cities_temp.pop(line)
            trash_temp.pop(line)
            distance = np.delete(distance, line, 0) #deleta linha
            distance = np.delete(distance, line, 1) #delete coluna
    return citieslist

def getCityTrash(city):
    index = cities.index(city)
    return trash[index]

def getDistanceBetweenCites(cityA, cityB):
    indexA = cities.index(cityA)
    indexB = cities.index(cityB)
    return distance[indexA][indexB]

def getSubTrash(arr):
    totalTrash = 0
    #print ("Sub-arranjo: ", arr, " Tamanho: ", len(arr))
    for i in arr:
        #print ("Cidade(s):", i)
        for c in i:
            #logging.debug("Pegando quantidade de lixo para a cidade ", c)
            totalTrash = totalTrash + getCityTrash(c)
    #print ("Total Lixo: ", totalTrash, "Faixa: ", getFaixa(totalTrash))
    return totalTrash

def getFaixa(sumTrash):
    for i in range(len(rsutrash)):
        if sumTrash > rsutrash[i] and sumTrash < rsutrash[i+1]:
            return i

def getSubCapex(range, trashSum):
    cpRT1 = capexRT1[range]-(trashSum*(capexRT1[range]-capexRT1[range+1]))/(rsutrash[range+1]-rsutrash[range])
    cpRT2 = capexRT2[range]-(trashSum*(capexRT2[range]-capexRT2[range+1]))/(rsutrash[range+1]-rsutrash[range])
    cpRT3 = capexRT3[range]-(trashSum*(capexRT3[range]-capexRT3[range+1]))/(rsutrash[range+1]-rsutrash[range])
    cpRT4 = capexRT4[range]-(trashSum*(capexRT4[range]-capexRT4[range+1]))/(rsutrash[range+1]-rsutrash[range])
    return (cpRT1 + cpRT2 + cpRT3 + cpRT4)/4

def getSubOpex(range, trashSum):
    opRT1 = opexRT1[range]-(trashSum*(opexRT1[range]-opexRT1[range+1]))/(rsutrash[range+1]-rsutrash[range])
    opRT2 = opexRT2[range]-(trashSum*(opexRT2[range]-opexRT2[range+1]))/(rsutrash[range+1]-rsutrash[range])
    opRT3 = opexRT3[range]-(trashSum*(opexRT3[range]-opexRT3[range+1]))/(rsutrash[range+1]-rsutrash[range])
    opRT4 = opexRT4[range]-(trashSum*(opexRT4[range]-opexRT4[range+1]))/(rsutrash[range+1]-rsutrash[range])
    return (opRT1 + opRT2 + opRT3 + opRT4)/4

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
    logging.debug("Quantidade de cidadesde superior a ", MAX_CITIES, ", o algoritmo irá clusterizar as cidades.")
    # Call clusterization to reduce the number of cities
    clusterization(citieslist, distance)

logging.debug("Cálculando combinaçãoes...")
combinations = list()
combinations += list(set_partitions(citieslist))

logging.debug(len(combinations))

for i in combinations:
    trashArray = 0
    capexOpexArray = 0
    
    for y in i:
        trashSubArray = getSubTrash(y)
        capexSubArray = getSubCapex(getFaixa(trashSubArray), trashSubArray)
        opexSubArray = getSubOpex(getFaixa(trashSubArray), trashSubArray)
        capexOpexValue = ((capexSubArray+opexSubArray * 30.0) * trashSubArray * 312.0)/(trashSubArray * 312.0 * 30.0)
        trashArray = trashArray + trashSubArray
        capexOpexArray = (capexOpexValue * trashSubArray) + capexOpexArray
    finalValue = capexOpexArray/trashArray

    print ("Arranjo: ", i, "Valor Calculado: ", finalValue)