import sys
import logging
import numpy as np
import csv
import copy
import itertools
from more_itertools import set_partitions

# Configure logs
logging.basicConfig(
    stream=sys.stderr, 
    #level=logging.DEBUG,
    level=logging.INFO,
    format='[%(asctime)s] {%(filename)s:%(lineno)d} %(levelname)s - %(message)s',
)

# Parser command line parameters
try:
    csvcities = sys.argv[1]
    csvdistance = sys.argv[2]
    reportfile = sys.argv[3]
except IndexError:
    raise SystemExit(f"Usage: {sys.argv[0]} <cities.csv> <distance.csv>")

# The maximium number of allowed cities to generate the combinations
# due to performance issues.
MAX_CITIES = 10

# Custo de Movimentação
CUST_MOV_RESIDUOS = 1
CUST_MOV_REJEITOS = 0.7

# Threshold lixo
THRESHOLD_TRASH = 75.0

f = open(reportfile, "w")
tool = open("tool.csv", "w")

f.write("============= PARAMETROS ============= \n")
f.write("Máximo de cidades: " + repr(MAX_CITIES) + "\n")
f.write("Quantidade de lixo mínimo para um sub-arranjo: " + repr(THRESHOLD_TRASH) + "\n")
f.write("Custo Movimentação Resíduos: " + repr(CUST_MOV_RESIDUOS) + "\n")
f.write("Custo Movimentação Rejeitos: " + repr(CUST_MOV_REJEITOS) + "\n\n\n")

# RSU
rsutrash = [0,25,75,150,250,350,700,1250,2500,5000]
fator = [0.75,0.80,0.85,0.95,1,1,1,1,1,1]
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

# Read cities from CSV file
cities = list()
trash = list()
utvr = list()
aterro = list()
citieslist = list()
trashlist = list()
aterros_only = list()
utvrs_only = list()
custo_convencional = list()
custo_transbordo = list()
custo_pos_transbordo = list()

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
        for l in citieslist:
            if cities_temp[line] in l:
                trash_line = getSubTrash(l)
            if cities_temp[column] in l:
                trash_column = getSubTrash(l)
        centrodemassa = ""
        outracidade = ""

        logging.debug("A menor distância é  %d vou unir as cidades %s (%f) e %s (%f)", min, cities_temp[line], trash_line, cities_temp[column], trash_column)
        if trash_line > trash_column:
        #if trash_temp[line] > trash_temp [column]:
            logging.debug("A cidade %s é o centro de massa (%f) e irá representar o cluster", cities_temp[line], trash_line)
            #print("A cidade ", cities_temp[line], "é o centro de massa (", trash_temp[line] ,") e irá representar o cluster")
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
            #print(index_add)
            for a in citieslist[index_remove]:
                citieslist[index_add].append(a)
            citieslist.pop(index_remove)
            #print(index_remove)
            cities_temp.pop(column)
            trash_temp.pop(column)
            distance = np.delete(distance, column, 0) #deleta linha
            distance = np.delete(distance, column, 1) #delete coluna
        else:
            #print("A cidade ", cities_temp[column], "é o centro de massa (", trash_temp[column] ,") e irá representar o cluster")
            logging.debug("A cidade %s é o centro de massa (%f) e irá representar o cluster", cities_temp[column], trash_column)
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
            #print(index_add)
            for a in citieslist[index_remove]:
                citieslist[index_add].append(a)
            citieslist.pop(index_remove)
            #print(index_remove)

            cities_temp.pop(line)
            trash_temp.pop(line)
            distance = np.delete(distance, line, 0) #deleta linha
            distance = np.delete(distance, line, 1) #delete coluna
    #time.sleep(10000)
    return citieslist

def getCityTrash(city):
    index = cities.index(city)
    return trash[index]

def getCitycusto_convencional(city):
    index = cities.index(city)
    return custo_convencional[index]

def getCitycusto_transbordo(city):
    index = cities.index(city)
    return custo_transbordo[index]

def getCitycusto_pos_transbordo(city):
    index = cities.index(city)
    return custo_pos_transbordo[index]

def getDistanceBetweenCites(cityA, cityB):
    indexA = cities.index(cityA)
    indexB = cities.index(cityB)
    return distance[indexA][indexB]

def getSubTrash(arr):
    totalTrash = 0
    #print ("Sub-arranjo: ", arr, " Tamanho: ", len(arr))
    for i in arr:
        #print ("Cidade(s):", i)
        #for c in i:
            #logging.debug("Pegando quantidade de lixo para a cidade ", c)
        totalTrash = totalTrash + getCityTrash(i)
    #print ("Total Lixo: ", totalTrash, "Faixa: ", getFaixa(totalTrash))
    return totalTrash

def getFaixa(sumTrash):
    for i in range(len(rsutrash)):
        if sumTrash > rsutrash[i] and sumTrash < rsutrash[i+1]:
            return i

def removeArraysWithoutUTVR(combinations):
    comb = combinations.copy()
    for c in range(len(comb)):
        #print("Comb: ", comb[c])
        for sub in comb[c]:
            find = False
            for cluster in sub:
                for city in cluster:
                    if utvr[cities.index(city)] == "sim":
                        find = True
            if not find:
                #print("Combinação inválida: ", comb[c])
                combinations.remove(comb[c])
                break

def removeArraysTrashThreshold(combinations):
    threshold = THRESHOLD_TRASH
    comb = combinations.copy()
    for c in range(len(comb)):
        #print("Comb: ", comb[c])
        for sub in comb[c]:
            lixo = 0
            for cluster in sub:
                #print("Sub: ", sub)
                lixo = lixo + getSubTrash(cluster)

            #print("Lixo: ", lixo)    
            if lixo < threshold:
                combinations.remove(comb[c])
                break

def funccentrodemassa(cluster, cities, trash):
    max = 0
    for i in cluster:
        index = cities.index(i)
        if trash[index] > max:
            max = trash[index]
            c_centrodemassa = cities[index]
    return c_centrodemassa

def inboundoutbound(subarray, isCentralized):
    data = []
    for utvr_city in subarray:
        entry = {}
        sum_inbound = 0
        if utvr_city in utvrs_only:
            logging.debug("%s é uma UTVR...", utvr_city)
            entry["sub-arranjo"] = subarray
            entry["utvr"] = utvr_city
            for other_city in subarray:
                logging.debug("A distância de %s para %s é de %f. O lixo produzido por %s é %f", utvr_city, other_city, (getDistanceBetweenCites(utvr_city,other_city) * CUST_MOV_RESIDUOS), other_city, getCityTrash(other_city))
                #sum_inbound = sum_inbound + (getDistanceBetweenCites(utvr_city,other_city) * CUST_MOV_RESIDUOS)
                #sum_inbound = sum_inbound + ((getDistanceBetweenCites(utvr_city,other_city) * CUST_MOV_RESIDUOS) * getCityTrash(other_city))
                #sum_inbound = sum_inbound + ((getDistanceBetweenCites(utvr_city,other_city) * getCitycusto_pos_transbordo(other_city)) + getCitycusto_convencional(other_city) + getCitycusto_transbordo(other_city)) * getCityTrash(other_city)
                #sum_inbound = sum_inbound + ((getCitycusto_convencional(other_city) * getCityTrash(other_city)) + (getCitycusto_transbordo(other_city) * getCityTrash(other_city)) + (getCitycusto_pos_transbordo(other_city) * getDistanceBetweenCites(utvr_city,other_city) * getCityTrash(other_city))) * getCityTrash(other_city)
                sum_inbound = sum_inbound + ((getCitycusto_convencional(other_city)) + (getCitycusto_transbordo(other_city)) + (getCitycusto_pos_transbordo(other_city) * getDistanceBetweenCites(utvr_city,other_city))) * getCityTrash(other_city)
                
                if isCentralized and utvr_city == "Presidente Prudente":
                    logging.debug("Inbound = ((custo convencional %f) + (custo transbordo %f) + (pos transbordo %f * distancia %f)) * %f \n", 
                        getCitycusto_convencional(other_city), 
                         getCitycusto_transbordo(other_city), 
                         getCitycusto_pos_transbordo(other_city), getDistanceBetweenCites(utvr_city,other_city),
                         getCityTrash(other_city))
                    logging.debug("Inbound atual = %f", sum_inbound)
            if isCentralized and utvr_city == "Presidente Prudente":
                logging.debug("Inbound Final = %f / %f = %f \n", sum_inbound, getSubTrash(subarray),  sum_inbound / getSubTrash(subarray))
            sum_inbound = sum_inbound / getSubTrash(subarray)
            entry["inbound"] = sum_inbound
            #print("Inbound: ", entry["inbound"])
            for a in aterros_only:
                e = copy.deepcopy(entry)
                sum_outbound = 0
                sum_outbound = sum_outbound + (getDistanceBetweenCites(utvr_city,a) * (CUST_MOV_REJEITOS * getCitycusto_pos_transbordo(utvr_city))) * 0.5
                e["aterro"] = a
                e["outbound"] = sum_outbound
                e["total"] = sum_inbound + sum_outbound
                
                logging.debug("Adicionando: %s", e)
                data.append(e)
    
    data = sorted(data, key = lambda k: (k["total"]))
    #for d in data:
    #    print(d)
    #print()
    #print("========> Estou selecionando o dado: ", data[0])
    #print(len(data))
    #print()
    if isCentralized:
        #Retorna a UTVR sendo o centro de massa, não o mais eficaz
        cmassa = funccentrodemassa(subarray, cities, trash)
        print("CENTRALIZADO", subarray)
        print("CENTRO DE MASSA", cmassa)
        for d in data:
            if d["utvr"] == cmassa:
                return d
    else:
        return data[0]

def getSubCapex(range, trashSum):
    

    cpRT1 = capexRT1[range]*fator[range] + ((capexRT1[range]*fator[range]-capexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    cpRT2 = capexRT2[range]*fator[range] + ((capexRT2[range]*fator[range]-capexRT2[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    cpRT3 = capexRT3[range]*fator[range] + ((capexRT3[range]*fator[range]-capexRT3[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    cpRT4 = capexRT4[range]*fator[range] + ((capexRT4[range]*fator[range]-capexRT4[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    cpRT5 = capexRT1[range]*fator[range] + ((capexRT1[range]*fator[range]-capexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    
    if trashSum <= 75:
        return ((cpRT1 + cpRT2)/2)
    if trashSum <= 150:
        return (cpRT1 + cpRT2 + cpRT3)/3
    if trashSum <= 250:
        return (cpRT1 + cpRT2 + cpRT3 + cpRT5)/4
    else:
        return (cpRT1 + cpRT2 + cpRT3 + cpRT4 + cpRT5)/5

def getSubOpex(range, trashSum):
    opRT1 = opexRT1[range]*fator[range] + ((opexRT1[range]*fator[range]-opexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    opRT2 = opexRT2[range]*fator[range] + ((opexRT2[range]*fator[range]-opexRT2[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    opRT3 = opexRT3[range]*fator[range] + ((opexRT3[range]*fator[range]-opexRT3[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    opRT4 = opexRT4[range]*fator[range] + ((opexRT4[range]*fator[range]-opexRT4[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    opRT5 = opexRT1[range]*fator[range] + ((opexRT1[range]*fator[range]-opexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    
    if trashSum <= 75:
        return ((opRT1 + opRT2)/2)
    if trashSum <= 150:
        return (opRT1 + opRT2 + opRT3)/3
    if trashSum <= 250:
        return (opRT1 + opRT2 + opRT3 + opRT5)/4
    else:
        return (opRT1 + opRT2 + opRT3 + opRT4 + opRT5)/5

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
        utvr.append(row["utvr"])
        if row["utvr"] == "sim":
            utvrs_only.append(row["city"])
        aterro.append(row["aterro"])
        if row["aterro"] == "sim":
            aterros_only.append(row["city"])
        custo_convencional.append(float(row["custo_convencional"]))
        custo_transbordo.append(float(row["custo_transbordo"]))
        custo_pos_transbordo.append(float(row["custo_pos_transbordo"]))
        trashlist.append(newtrash)
        line_count += 1

cities_temp = cities.copy()
trash_temp = trash.copy()

# Read distances from CSV file
distance = np.loadtxt(open(csvdistance, "rb"), delimiter=",", skiprows=0)

if len(citieslist) > MAX_CITIES:
    logging.debug("Quantidade de cidades superior a %d o algoritmo irá clusterizar as cidades.", MAX_CITIES)
    # Call clusterization to reduce the number of cities
    clusterization(citieslist, distance)

f.write("============= ATERROS ============= \n")
for a in aterros_only:
    f.write(a + "\n")
f.write("\n\n\n")

f.write("============= CLUSTERS ============= \n")
i = 1 
for cl in citieslist:
    print(cl)
    f.write(repr(i) + ".\t" + repr(cl) + "\n")
    i = i + 1

f.write("\n\n\n============= ESTATÍSTICAS ============= \n")
logging.info("Cálculando combinaçãoes...")
combinations = list()
combinations += list(set_partitions(citieslist))
logging.info("Quantidade de combinações: %d", len(combinations))
f.write("Quantidade de combinações: " + repr(len(combinations)) + "\n")

logging.info("Removendo combinaçãoes cujo sub-arranjo não possui uma UTVR...")
removeArraysWithoutUTVR(combinations)
logging.info("Quantidade de combinações após a remoção: %d", len(combinations))
f.write("Quantidade de combinações (desconsiderando arranjos com sub-arranjos sem UTVR): " + repr(len(combinations)) + "\n")

logging.info("Removendo combinaçãoes cujo sub-arranjo não possui a quantidade de lixo necessária...")
removeArraysTrashThreshold(combinations)
logging.info("Quantidade de combinações após a remoção: %d", len(combinations))
f.write("Quantidade de combinações (desconsiderando arranjos com sub-arranjos que não somam a quantidade de lixo produzida mínima): " + repr(len(combinations)) + "\n\n\n")

logging.info("Cálculando valores (inbound, tecnologia e outbound) por combinação...")

new_comb = list()
for c in combinations:
    xcomb = list()
    for sub in c:
        subarray = list()
        for cluster in sub:   
            for city in cluster:
                subarray.append(city)
        xcomb.append(subarray)
    new_comb.append(xcomb)

data = []
current = 0
for i in new_comb:
    if current % 10000 == 0:
        print(current)
        logging.debug(current)

    trashArray = 0
    capexOpexArray = 0
    inboundArray = 0
    outboundArray = 0
    #print("Arranjo: ", i)
    logging.debug("Arranjo: %s", i)
    new = {}

    centralizado = False
    if len(i) == 1:
        centralizado = True

    sub = list()
    for y in i:
        logging.debug("Sub-arranjo: %s", y)
        trashSubArray = getSubTrash(y)
        capexSubArray = getSubCapex(getFaixa(trashSubArray), trashSubArray)
        opexSubArray = getSubOpex(getFaixa(trashSubArray), trashSubArray)
        capexOpexValue = ((capexSubArray+opexSubArray * 30.0) * trashSubArray * 313.0)/(trashSubArray * 313.0 * 30.0)
        trashArray = trashArray + trashSubArray
        capexOpexArray = (capexOpexValue * trashSubArray) + capexOpexArray
        rsinout = inboundoutbound(y, centralizado)
        rsinout["capex"] =  0#capexSubArray
        rsinout["opex"] = 0#opexSubArray
        rsinout["tecnologia"] = 0#capexOpexValue
        rsinout["capex"] =  capexSubArray
        rsinout["opex"] = opexSubArray
        rsinout["tecnologia"] = capexOpexValue
        
        #print("IN OUT: ", rsinout)
        inboundArray = inboundArray + (rsinout["inbound"] * trashSubArray)
        outboundArray = outboundArray + (rsinout["outbound"] * trashSubArray)
        rsinout["lixo"] = trashSubArray
        rsinout["total"] = capexOpexValue + rsinout["inbound"] + rsinout["outbound"]
        sub.append(rsinout)
        
    cpopfinalValue = 0 #capexOpexArray/trashArray
    cpopfinalValue = capexOpexArray/trashArray
    
    new["arranjo"] = i
    new["sub"] = sub
    new["capexopex"] = 0 #cpopfinalValue
    new["capexopex"] = cpopfinalValue
    new["lixo-array"] = trashArray
    new["inbound"] = inboundArray/trashArray
    new["outbound"] = outboundArray/trashArray
    new["total"] = cpopfinalValue + (inboundArray/trashArray) + (outboundArray/trashArray)
    data.append(new)
    #print ("Arranjo: ", i, "Valor Calculado: ", finalValue)
    #break

    current = current + 1;

logging.info("Ordenando combinações...")
data = sorted(data, key = lambda k: (k["total"]))

f.write("\n\n============= ARRANJO CENTRALIZADO ============= \n")
for d in data:
    if len(d["arranjo"]) == 1:
        f.write(repr(d["arranjo"]) + "\n")
        f.write("- Lixo: " + repr(d["lixo-array"]) + "\n")
        f.write("- Custo Total: " + repr(d["total"]) + "\n")
        print("\t Inbound", d["inbound"])
        f.write("-- Inbound: " + repr(d["inbound"]) + "\n")
        print("\t Tecnologia", d["capexopex"])
        f.write("-- Tecnologia: " + repr(d["capexopex"]) + "\n")
        print("\t Outbound", d["outbound"])
        f.write("-- Outbound: " + repr(d["outbound"]) + "\n\n")
        f.write("-- Sub-arranjos:\n")
        for x in range(len(d["sub"])):
            print("Sub-arranjo: ", d["sub"][x])
            f.write("\t" + repr(d["sub"][x]["sub-arranjo"]) + "\n")
            f.write("\t-- UTVR: " + repr(d["sub"][x]["utvr"]) + "\n")
            f.write("\t-- Aterro: " + repr(d["sub"][x]["aterro"]) + "\n")
            f.write("\t-- Lixo: " + repr(d["sub"][x]["lixo"]) + "\n")
            f.write("\t-- Total: " + repr(d["sub"][x]["total"]) + "\n")
            f.write("\t-- Inbound: " + repr(d["sub"][x]["inbound"]) + "\n")
            f.write("\t-- Tecnologia: " + repr(d["sub"][x]["tecnologia"]) + "\n")
            f.write("\t\t-- Capex: " + repr(d["sub"][x]["capex"]) + "\n")
            f.write("\t\t-- Opex: " + repr(d["sub"][x]["opex"]) + "\n")
            f.write("\t-- Outbound: " + repr(d["sub"][x]["outbound"]) + "\n\n")
            break

f.write("\n\n============= TOP 5 ARRANJOS ============= \n")
for i in range(5):
    #print(data[i])
    tool.write(repr(data[i]["arranjo"]) + ";Sumário;" + repr(data[i]["lixo-array"]) + ";" + repr(data[i]["inbound"])  + ";" + repr(data[i]["outbound"]) + "\n")


    f.write(repr(i+1) + ".\t" + repr(data[i]["arranjo"]) + "\n")
    f.write("- Lixo: " + repr(data[i]["lixo-array"]) + "\n")
    f.write("- Custo Total: " + repr(data[i]["total"]) + "\n")
    print(i, " - Arranjo: ", data[i]["arranjo"], " ", data[i]["total"])
    print("\t Inbound", data[i]["inbound"])
    f.write("-- Inbound: " + repr(data[i]["inbound"]) + "\n")
    print("\t Tecnologia", data[i]["capexopex"])
    f.write("-- Tecnologia: " + repr(data[i]["capexopex"]) + "\n")
    print("\t Outbound", data[i]["outbound"])
    f.write("-- Outbound: " + repr(data[i]["outbound"]) + "\n\n")
    f.write("-- Sub-arranjos:\n")
    for x in range(len(data[i]["sub"])):
        tool.write(repr(data[i]["arranjo"]) + ";" + repr(data[i]["sub"][x]["sub-arranjo"]) + ";" + repr(data[i]["sub"][x]["lixo"]) + ";" + repr(data[i]["sub"][x]["inbound"])  + ";" + repr(data[i]["sub"][x]["outbound"]) + "\n")


        print("Sub-arranjo: ", data[i]["sub"][x])
        f.write("\t" + repr(data[i]["sub"][x]["sub-arranjo"]) + "\n")
        f.write("\t-- UTVR: " + repr(data[i]["sub"][x]["utvr"]) + "\n")
        f.write("\t-- Aterro: " + repr(data[i]["sub"][x]["aterro"]) + "\n")
        f.write("\t-- Lixo: " + repr(data[i]["sub"][x]["lixo"]) + "\n")
        f.write("\t-- Total: " + repr(data[i]["sub"][x]["total"]) + "\n")
        f.write("\t-- Inbound: " + repr(data[i]["sub"][x]["inbound"]) + "\n")
        f.write("\t-- Tecnologia: " + repr(data[i]["sub"][x]["tecnologia"]) + "\n")
        f.write("\t\t-- Capex: " + repr(data[i]["sub"][x]["capex"]) + "\n")
        f.write("\t\t-- Opex: " + repr(data[i]["sub"][x]["opex"]) + "\n")
        f.write("\t-- Outbound: " + repr(data[i]["sub"][x]["outbound"]) + "\n\n")

    f.write("-----------------------------------------------------------------\n\n")

# Close report and tool file
f.close()
tool.close