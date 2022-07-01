import sys
import logging
import numpy as np
import csv
import copy
from more_itertools import set_partitions

def clusterization(data, distance, max):
    clusters = []
    for d in data:
        city = []
        city.append(d["name"])
        clusters.append(city)
    
    # Only reduce the number of cities if necessary
    if len(clusters) <= max:
        return clusters
    
    cities_temp = getCities(data)
    trash_temp = getTrash(data)
    trashlist = trash_temp.copy()
    while len(clusters) > max:
        min = 999999999
        line = 0
        column = 0
        # Locate the shortest distance between two cities
        for l in range(len(cities_temp)):
            for c in range(len(cities_temp)):
                dist = distance[l, c]
                if dist < min and dist != 0:
                    min = dist
                    line = l
                    column = c

        # Calculate the current t/d for a given cluster
        for c in clusters:
            if cities_temp[line] in c:
                trash_line = getSubTrash(data, c)
            if cities_temp[column] in c:
                trash_column = getSubTrash(data, c)
        centrodemassa = ""
        outracidade = ""

        logging.debug("A menor distância é  %d vou unir as cidades %s (%f) e %s (%f)", min, cities_temp[line], trash_line, cities_temp[column], trash_column)
        if trash_line > trash_column:
            logging.debug("A cidade %s é o centro de massa (%f) e irá representar o cluster", cities_temp[line], trash_line)
            centrodemassa = cities_temp[line]
            outracidade = cities_temp[column]
            cities_temp.pop(column)
            trash_temp.pop(column)
            distance = np.delete(distance, column, 0) #deleta linha
            distance = np.delete(distance, column, 1) #delete coluna
        else:
            logging.debug("A cidade %s é o centro de massa (%f) e irá representar o cluster", cities_temp[column], trash_column)
            centrodemassa = cities_temp[column]
            outracidade = cities_temp[line]
            cities_temp.pop(line)
            trash_temp.pop(line)
            distance = np.delete(distance, line, 0) #deleta linha
            distance = np.delete(distance, line, 1) #delete coluna
            
        index_add = 0
        index_remove = 0
        index = 0
        for cl in clusters:
            for ct in cl:
                if ct == centrodemassa:
                    index_add = index
                if ct == outracidade:
                    index_remove = index
            index = index + 1

        for a in clusters[index_remove]:
            clusters[index_add].append(a)
        clusters.pop(index_remove)

        
    return clusters

def getCityAttribute(data, city, attribute):
    try:
        for d in data:
            if d["name"] == city:
                return d[attribute]
    except IndexError:
        raise SystemExit(f"City: {city} or attribute: {attribute} not found.")

def getDistanceBetweenCites(data, distance, cityA, cityB):
    for i in range(len(data)):
        if data[i]["name"] == cityA:
            indexA = i
        if data[i]["name"] == cityB:
            indexB = i
    return distance[indexA][indexB]

def getFaixa(sumTrash, rsutrash):
    for i in range(len(rsutrash)):
        if sumTrash > rsutrash[i] and sumTrash < rsutrash[i+1]:
            return i

def removeArraysWithoutUTVR(combinations, utvrs):
    comb = combinations.copy()
    for c in range(len(comb)):
        #print("Comb: ", comb[c])
        for sub in comb[c]:
            find = False
            for cluster in sub:
                for city in cluster:
                    if city in utvrs:
                        find = True
            if not find:
                #print("Combinação inválida: ", comb[c])
                combinations.remove(comb[c])
                break

def removeArraysTrashThreshold(data, combinations, threshold):
    comb = combinations.copy()
    sublist = list()
    trashtotallist = list()
    for c in range(len(comb)):
        if c % 1000 == 0:
            print(c, " ", len(sublist))
        for sub in comb[c]:
            lixo = 0
            if sub in sublist:
                lixo = trashtotallist[sublist.index(sub)]
            else:
                for cluster in sub:
                    lixo = lixo + getSubTrash(data, cluster)
                #sublist.append(sub)
                #trashtotallist.append(lixo)
            if lixo < threshold:
                combinations.remove(comb[c])
                break
    return combinations

def funccentrodemassa(data, cluster):
    max = 0
    for i in cluster:
        trash = getCityAttribute(data, i, "trash")
        if trash > max:
            max = trash
            c_centrodemassa = i
    return c_centrodemassa

def inboundoutbound(cdata, distance, subarray, isCentralized, utvrs_only, aterros_only):
    data = []
    CAPEX_INBOUND = 150
    CAPEX_OUTBOUND = 25
    for utvr_city in subarray:
        entry = {}
        sum_inbound = 0
        if utvr_city in utvrs_only:
            logging.debug("%s é uma UTVR...", utvr_city)
            entry["sub-arranjo"] = subarray
            entry["utvr"] = utvr_city
            for other_city in subarray:
                conventional_cost = getCityAttribute(cdata, other_city, "conventional-cost")
                transshipment_cost = getCityAttribute(cdata, other_city, "transshipment-cost")
                cost_post_transhipment = getCityAttribute(cdata, other_city, "cost-post-transhipment")
                trash = getCityAttribute(cdata, other_city, "trash")
                sum_inbound = sum_inbound + ((conventional_cost) + (transshipment_cost) + (cost_post_transhipment * getDistanceBetweenCites(cdata, distance, utvr_city, other_city))) * trash
                
                if isCentralized and utvr_city == "Presidente Prudente":
                    logging.debug("Inbound atual = %f", sum_inbound)
            if isCentralized and utvr_city == "Presidente Prudente":
                logging.debug("Inbound Final = %f / %f = %f \n", sum_inbound, getSubTrash(cdata, subarray),  sum_inbound / getSubTrash(cdata, subarray))
            sum_inbound = sum_inbound / getSubTrash(cdata, subarray)
            sum_inbound = (CAPEX_INBOUND/35.0 + sum_inbound) /  1.0
            entry["inbound"] = sum_inbound
            #print("Inbound: ", entry["inbound"])
            for a in aterros_only:
                e = copy.deepcopy(entry)
                sum_outbound = 0
                sum_outbound = sum_outbound + (getDistanceBetweenCites(cdata, distance, utvr_city,a) * (0.7 * getCityAttribute(cdata, utvr_city, "cost-post-transhipment"))) * 0.5
                e["aterro"] = a
                e["outbound"] = (CAPEX_OUTBOUND/35.0 + sum_outbound) / 1.0
                e["total"] = sum_inbound + sum_outbound
                
                logging.debug("Adicionando: %s", e)
                data.append(e)
    
    data = sorted(data, key = lambda k: (k["total"]))
    if isCentralized:
        #Retorna a UTVR sendo o centro de massa, não o mais eficaz
        cmassa = funccentrodemassa(cdata, subarray)
        print("CENTRALIZADO", subarray)
        print("CENTRO DE MASSA", cmassa)
        for d in data:
            if d["utvr"] == cmassa:
                return d
    else:
        return data[0]

def getSubCapex(range, trashSum, rsutrash):
    fator = [1,1,1,1,1,1,1,1,1,1]
    capexRT1 = [0,
    1668,
    925,
    733,
    650,
    612,
    2135,
    2133,
    1985,
    1974
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

    cpRT1 = capexRT1[range]*fator[range] + ((capexRT1[range]*fator[range]-capexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    cpRT2 = capexRT2[range]*fator[range] + ((capexRT2[range]*fator[range]-capexRT2[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    cpRT3 = capexRT3[range]*fator[range] + ((capexRT3[range]*fator[range]-capexRT3[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    cpRT4 = capexRT4[range]*fator[range] + ((capexRT4[range]*fator[range]-capexRT4[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    cpRT5 = capexRT1[range]*fator[range] + ((capexRT1[range]*fator[range]-capexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    
    if trashSum <= 75:
        return ((cpRT1))
    if trashSum <= 150:
        return (cpRT1)
    if trashSum <= 350:
        return (cpRT1)
    else:
        return (cpRT1)

def getSubOpex(range, trashSum, rsutrash):
    fator = [1,1,1,1,1,1,1,1,1,1]
    opexRT1 = [0,
    314,
    135,
    89,
    71,
    64,
    214,
    198,
    188,
    179
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

    opRT1 = opexRT1[range]*fator[range] + ((opexRT1[range]*fator[range]-opexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    opRT2 = opexRT2[range]*fator[range] + ((opexRT2[range]*fator[range]-opexRT2[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    opRT3 = opexRT3[range]*fator[range] + ((opexRT3[range]*fator[range]-opexRT3[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    opRT4 = opexRT4[range]*fator[range] + ((opexRT4[range]*fator[range]-opexRT4[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    opRT5 = opexRT1[range]*fator[range] + ((opexRT1[range]*fator[range]-opexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    
    if trashSum <= 75:
        return ((opRT1))
    if trashSum <= 150:
        return (opRT1)
    if trashSum <= 350:
        return (opRT1)
    else:
        return (opRT1)

def getCities(data):
    cities = list()
    for d in data:
        cities.append(d["name"])
    return cities

def getTrash(data):
    trash = list()
    for d in data:
        trash.append(d["trash"])
    return trash

def getLandfill(data):
    landfill = list()
    for d in data:
        if d["landfill"]:
            landfill.append(d["name"])
    return landfill

def getUTVR(data):
    utvr = list()
    for d in data:
        if d["utvr"]:
            utvr.append(d["name"])
    return utvr

def getSubTrash(data, cluster):
    total = 0
    for c in cluster:
        total = total + getCityAttribute(data, c, "trash")
    return total

def main():
    # Configure logs
    logging.basicConfig(
        stream=sys.stderr, 
        #level=logging.DEBUG,
        level=logging.INFO,
        format='[%(asctime)s] {%(filename)s:%(lineno)d} %(levelname)s - %(message)s',
    )

    # Parser command line parameters
    try:
        CSVCITIES = sys.argv[1]                 # Cities file
        CSVDISTANCE = sys.argv[2]               # Distance matrix file
        MAX_CITIES = int(sys.argv[3])           # The maximium number of allowed cities to generate the combinations due to performance issues.
        TRASH_THRESHOLD = float(sys.argv[4])    # The minimun of trash for a sub-array
        REPORTFILE = sys.argv[5]                # The report file name
        OUTPUTFILE = sys.argv[6]                # The output file name
    except IndexError:
        raise SystemExit(f"Usage: {sys.argv[0]} <cities.csv> <distance.csv> <max cities> <trash threshold> <report.txt> <output.csv>")

    # Output files
    report = open(REPORTFILE, "w")
    output = open(OUTPUTFILE, "w")

    # Print parameters in report file
    report.write("============= PARAMETROS ============= \n")
    report.write("Arquivo de cidades: " + repr(CSVCITIES) + "\n")
    report.write("Arquivo de distâncias: " + repr(CSVDISTANCE) + "\n")
    report.write("Máximo de cidades: " + repr(MAX_CITIES) + "\n")
    report.write("Quantidade de lixo mínimo para um sub-arranjo: " + repr(TRASH_THRESHOLD) + "\n\n\n")

    # RSU
    rsutrash = [0,25,75,150,250,350,700,1250,2500,5000]
    
    citiesdata = []
    clusters = []
    with open(CSVCITIES, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                line_count += 1
            city = {}
            city["name"] = row["Município"]
            city["trash"] = float(row["Lixo (t/d)"])
            if row["UTVR"] == "Sim":
                city["utvr"] = True
            else:
                city["utvr"] = False
            if row["Aterro Existente"] == "Sim":
                city["landfill"] = True
            else:
                city["landfill"] = False
            city["conventional-cost"] = float(row["Custo de Coleta Mista Convencional"])
            city["transshipment-cost"] = float(row["Custo de Coleta e Transbordo de Resíduos Mistos"])
            city["cost-post-transhipment"] = float(row["Custo de Transporte Pós Transbordo"])
            citiesdata.append(city)
            line_count += 1

    # Read distances from CSV file
    distance = np.loadtxt(open(CSVDISTANCE, "rb"), delimiter=",", skiprows=0)

    if len(citiesdata) > MAX_CITIES:
        logging.debug("Quantidade de cidades superior a %d o algoritmo irá clusterizar as cidades.", MAX_CITIES)
    # Call clusterization to reduce the number of cities or just to build a list of list
    clusters = clusterization(citiesdata, distance, MAX_CITIES)

    # Print landfills in the report file
    report.write("============= ATERROS ============= \n")
    landfill = getLandfill(citiesdata)
    for l in landfill:
        report.write(l + "\n")
    report.write("\n\n\n")

    # Print clusters in the report file
    report.write("============= CLUSTERS ============= \n")
    i = 1 
    for c in clusters:
        print(c)
        report.write(repr(i) + ".\t" + repr(c) + "\n")
        i += 1

    report.write("\n\n\n============= ESTATÍSTICAS ============= \n")
    logging.info("Cálculando combinaçãoes...")
    combinations = list()
    combinations += list(set_partitions(clusters))
    logging.info("Quantidade de combinações: %d", len(combinations))
    report.write("Quantidade de combinações: " + repr(len(combinations)) + "\n")

    utvrs = getUTVR(citiesdata)
    if len(utvrs) != len(citiesdata):
        logging.info("Removendo combinaçãoes cujo sub-arranjo não possui uma UTVR...")
        combinations = removeArraysWithoutUTVR(combinations)
        logging.info("Quantidade de combinações após a remoção: %d", len(combinations))
    report.write("Quantidade de combinações (desconsiderando arranjos com sub-arranjos sem UTVR): " + repr(len(combinations)) + "\n")

    if TRASH_THRESHOLD > 0.0:
        logging.info("Removendo combinaçãoes cujo sub-arranjo não possui a quantidade de lixo necessária...")
        combinations = removeArraysTrashThreshold(citiesdata, combinations, TRASH_THRESHOLD)
        logging.info("Quantidade de combinações após a remoção: %d", len(combinations))
    report.write("Quantidade de combinações (desconsiderando arranjos com sub-arranjos que não somam a quantidade de lixo produzida mínima): " + repr(len(combinations)) + "\n\n\n")

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
            trashSubArray = getSubTrash(citiesdata, y)
            capexSubArray = getSubCapex(getFaixa(trashSubArray, rsutrash), trashSubArray, rsutrash)
            opexSubArray = getSubOpex(getFaixa(trashSubArray, rsutrash), trashSubArray, rsutrash)
            capexOpexValue = (capexSubArray/35.0 + opexSubArray)/ 1.0
            trashArray = trashArray + trashSubArray
            capexOpexArray = (capexOpexValue * trashSubArray) + capexOpexArray
            rsinout = inboundoutbound(citiesdata, distance, y, centralizado, utvrs, landfill)
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

        current = current + 1

    logging.info("Ordenando combinações...")
    data = sorted(data, key = lambda k: (k["total"]))

    report.write("\n\n============= ARRANJO CENTRALIZADO ============= \n")
    for d in data:
        if len(d["arranjo"]) == 1:
            report.write(repr(d["arranjo"]) + "\n")
            report.write("- Lixo: " + repr(d["lixo-array"]) + "\n")
            report.write("- Custo Total: " + repr(d["total"]) + "\n")
            report.write("-- Inbound: " + repr(d["inbound"]) + "\n")
            report.write("-- Tecnologia: " + repr(d["capexopex"]) + "\n")
            report.write("-- Outbound: " + repr(d["outbound"]) + "\n\n")
            report.write("-- Sub-arranjos:\n")

            output.write(repr(d["arranjo"]) + ";Sumário;NA;NA;" + repr(d["total"]) + ";" + repr(d["lixo-array"]) + ";" + repr(d["capexopex"]) + ";" + repr(d["inbound"])  + ";" + repr(d["outbound"]) + "\n")


            for x in range(len(d["sub"])):
                output.write(repr(d["arranjo"]) + ";" + repr(d["sub"][x]["sub-arranjo"]) + ";" + repr(d["sub"][x]["aterro"]) + ";" + repr(d["sub"][x]["utvr"]) + ";" + repr(d["sub"][x]["total"]) + ";" + repr(d["sub"][x]["lixo"]) + ";" + repr(d["sub"][x]["tecnologia"]) + ";" + repr(d["sub"][x]["inbound"])  + ";" + repr(d["sub"][x]["outbound"]) + "\n")

                report.write("\t" + repr(d["sub"][x]["sub-arranjo"]) + "\n")
                report.write("\t-- UTVR: " + repr(d["sub"][x]["utvr"]) + "\n")
                report.write("\t-- Aterro: " + repr(d["sub"][x]["aterro"]) + "\n")
                report.write("\t-- Lixo: " + repr(d["sub"][x]["lixo"]) + "\n")
                report.write("\t-- Total: " + repr(d["sub"][x]["total"]) + "\n")
                report.write("\t-- Inbound: " + repr(d["sub"][x]["inbound"]) + "\n")
                report.write("\t-- Tecnologia: " + repr(d["sub"][x]["tecnologia"]) + "\n")
                report.write("\t\t-- Capex: " + repr(d["sub"][x]["capex"]) + "\n")
                report.write("\t\t-- Opex: " + repr(d["sub"][x]["opex"]) + "\n")
                report.write("\t-- Outbound: " + repr(d["sub"][x]["outbound"]) + "\n\n")
                break

    report.write("\n\n============= TOP 5 ARRANJOS ============= \n")
    for i in range(len(data)):
        if i % 1000 != 0:
            continue
        output.write(repr(data[i]["arranjo"]) + ";Sumário;NA;NA;" + repr(data[i]["total"]) + ";" + repr(data[i]["lixo-array"]) + ";" + repr(data[i]["capexopex"]) + ";" + repr(data[i]["inbound"])  + ";" + repr(data[i]["outbound"]) + "\n")


        report.write(repr(i+1) + ".\t" + repr(data[i]["arranjo"]) + "\n")
        report.write("- Lixo: " + repr(data[i]["lixo-array"]) + "\n")
        report.write("- Custo Total: " + repr(data[i]["total"]) + "\n")
        report.write("-- Inbound: " + repr(data[i]["inbound"]) + "\n")
        report.write("-- Tecnologia: " + repr(data[i]["capexopex"]) + "\n")
        report.write("-- Outbound: " + repr(data[i]["outbound"]) + "\n\n")
        report.write("-- Sub-arranjos:\n")
        for x in range(len(data[i]["sub"])):
            output.write(repr(data[i]["arranjo"]) + ";" + repr(data[i]["sub"][x]["sub-arranjo"]) + ";" + repr(data[i]["sub"][x]["aterro"]) + ";" + repr(data[i]["sub"][x]["utvr"]) + ";" + repr(data[i]["sub"][x]["total"]) + ";" + repr(data[i]["sub"][x]["lixo"]) + ";" + repr(data[i]["sub"][x]["tecnologia"]) + ";" + repr(data[i]["sub"][x]["inbound"])  + ";" + repr(data[i]["sub"][x]["outbound"]) + "\n")

            report.write("\t" + repr(data[i]["sub"][x]["sub-arranjo"]) + "\n")
            report.write("\t-- UTVR: " + repr(data[i]["sub"][x]["utvr"]) + "\n")
            report.write("\t-- Aterro: " + repr(data[i]["sub"][x]["aterro"]) + "\n")
            report.write("\t-- Lixo: " + repr(data[i]["sub"][x]["lixo"]) + "\n")
            report.write("\t-- Total: " + repr(data[i]["sub"][x]["total"]) + "\n")
            report.write("\t-- Inbound: " + repr(data[i]["sub"][x]["inbound"]) + "\n")
            report.write("\t-- Tecnologia: " + repr(data[i]["sub"][x]["tecnologia"]) + "\n")
            report.write("\t\t-- Capex: " + repr(data[i]["sub"][x]["capex"]) + "\n")
            report.write("\t\t-- Opex: " + repr(data[i]["sub"][x]["opex"]) + "\n")
            report.write("\t-- Outbound: " + repr(data[i]["sub"][x]["outbound"]) + "\n\n")

        report.write("-----------------------------------------------------------------\n\n")

    # Close report and output file
    report.close()
    output.close

if __name__ == "__main__":
    main()