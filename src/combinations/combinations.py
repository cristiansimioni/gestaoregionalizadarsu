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
        if sumTrash >= rsutrash[i] and sumTrash <= rsutrash[i+1]:
            return i

def removeArraysWithoutUTVR(combinations, utvrs):
    comb = combinations.copy()
    for c in range(len(comb)):
        if c % (len(comb)/10.0) == 0:
            logging.info("Progreso: %d%%", c/len(comb)*100)
        for sub in comb[c]:
            find = False
            for cluster in sub:
                for city in cluster:
                    if city in utvrs:
                        find = True
            if not find:
                combinations.remove(comb[c])
                break

def removeArraysTrashThreshold(data, combinations, threshold):
    comb = combinations.copy()
    sublist = list()
    trashtotallist = list()
    for c in range(len(comb)):
        if c % (len(comb)/10.0) == 0:
            logging.info("Progreso: %d%%", c/len(comb)*100)
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

def inboundoutbound(cdata, distance, subarray, isCentralized, utvrs_only, aterros_only, existentlandfill):
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
            dist = 999999
            for a in existentlandfill:
                distCities = getDistanceBetweenCites(cdata, distance, utvr_city, a)
                if distCities < dist:
                    dist = distCities
                    sum_outbound = 0
                    sum_outbound = sum_outbound + (distCities * (0.7 * getCityAttribute(cdata, utvr_city, "cost-post-transhipment"))) * 0.5
                    entry["aterro-existente"] = a
                    entry["outbound-existente"] = (CAPEX_OUTBOUND/35.0 + sum_outbound) / 1.0
        
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
        for d in data:
            if d["utvr"] == cmassa:
                return d
    else:
        return data[0]

def getSubCapex(range, trashSum, rsutrash):
    fator = [1,1,1,1,1,1,1,1,1,1]
    fator = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
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
    capexRT1 = [0, 38717, 20334, 14235, 11184, 10340, 8963, 7978, 7240, 6665, 6204, 5827, 5512, 5246, 5017, 4819, 4646, 4492, 4356, 4234, 4124, 4024, 3933, 3849, 3773, 3702, 3637, 3576, 3520, 3467, 3418, 3372, 3329, 3288, 3250, 3213, 3179, 3146, 3115, 3086, 3058, 3031, 3005, 2980, 2957, 2934, 2913, 2892, 2872, 2853, 2834, 2838, 2820, 2820, 2803, 2787, 2771, 2756, 2741, 2727, 2713, 2700, 2686, 2674, 2661, 2649, 2638, 2626, 2615, 2605, 2594, 2584, 2573, 2563, 2554, 2544, 2535, 2526, 2517, 2508, 2499, 2488, 2480, 2472, 2464, 2456, 2448, 2441, 2434, 2426, 2419, 2412, 2405, 2399, 2392, 2385, 2379, 2373, 2366, 2364, 2360, 2350, 2346, 2343, 2339, 2335, 2346, 2342, 2338, 2335, 2332, 2328, 2325, 2322, 2318, 2315, 2312, 2309, 2306, 2303, 2300, 2297, 2294, 2292, 2289, 2286, 2283, 2281, 2278, 2276, 2273, 2270, 2268, 2266, 2263, 2261, 2258, 2256, 2254, 2252, 2250, 2248, 2247, 2245, 2243, 2241, 2240, 2238, 2236, 2235, 2233, 2232, 2230, 2229, 2227, 2226, 2224, 2223, 2221, 2233, 2231, 2229, 2228, 2226, 2225, 2224, 2222, 2221, 2220, 2219, 2217, 2216, 2215, 2214, 2213, 2211, 2210, 2209, 2208, 2207, 2206, 2205, 2204, 2203, 2202, 2201, 2200, 2199, 2198, 2197, 2196, 2195, 2194, 2193, 2192, 2191, 2190, 2191, 2190, 2189, 2188, 2187, 2186, 2035, 2035, 2035, 2035, 2035, 2035, 2035, 2035, 2035, 2048, 2048, 2048, 2048, 2048, 2048, 2048, 2048, 2048, 2048, 2049, 2049, 2049, 2049, 2049, 2049, 2049, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2047, 2048, 2048, 2048, 2048, 2048, 2048, 2048, 2048, 2048, 2048, 2048, 2049, 2049, 2049, 2049, 2049, 2049, 2049, 2050, 2050, 2050, 2050, 2050, 2050, 2051, 2051, 2051, 2051, 2051, 2052, 2052, 2052, 2052, 2053, 2053, 2053, 2053, 2054, 2054, 2054, 2054, 2055, 2055, 2055, 2055, 2056, 2057, 2057, 2057, 2057, 2058, 2058, 2071, 2071, 2071, 2071, 2072, 2072, 2072, 2073, 2073, 2073, 2074, 2074, 2074, 2075, 2075, 2075, 2076, 2076, 2076, 2077, 2077, 2077, 2078, 2078, 2078, 2079, 2079, 2079, 2080, 2080, 2081, 2081, 2081, 2082, 2082, 2082, 2083, 2083, 2084, 2084, 2084, 2085, 2085, 2086, 2086, 2087, 2087, 2087, 2088, 2088, 2089, 2089, 2090, 2090, 2090, 2091, 2091, 2092, 2092, 2093, 2093, 2094, 2094, 2094, 2095, 2095, 2096, 2096, 2097, 2097, 2097, 2098, 2098, 2099, 2099, 2100, 2100, 2101, 2101, 2101, 2102, 2102, 2103, 2103, 2104, 2104, 2104, 2105, 2105, 2106, 2106, 2107, 2107, 2108, 2108, 2109, 2109, 2109, 2110, 2110, 2121, 2121, 2122, 2122, 1815, 1816, 1817, 1819, 1820, 1821, 1822, 1823, 1824, 1826, 1827, 1828, 1829, 1830, 1832, 1833, 1834, 1835, 1836, 1838, 1839, 1840, 1841, 1842, 1844, 1845, 1846, 1847, 1848, 1850, 1851, 1852, 1853, 1854, 1856, 1857, 1858, 1859, 1860, 1862, 1863, 1864, 1865, 1867, 1868, 1869, 1870, 1872, 1873, 1874, 1875, 1876, 1878, 1879, 1880, 1882, 1883, 1884, 1886, 1887, 1888, 1890, 1891, 1892, 1894, 1895, 1897, 1898, 1899, 1901, 1902, 1903, 1905, 1906, 1907, 1909, 1910, 1912, 1913, 1914, 1916, 1917, 1918, 1920, 1921, 1923, 1924, 1925, 1927, 1928, 1929, 1931, 1932, 1934, 1935, 1936]
    capexRT1 = [0, 22893, 11568, 7876, 6029, 7044, 6141, 5499, 5018, 4646, 4350, 4109, 3909, 3741, 3598, 3475, 3368, 3274, 3192, 3119, 3054, 2953, 2855, 2765, 2683, 2608, 2538, 2473, 2413, 2358, 2305, 2257, 2211, 2168, 2128, 2090, 2054, 2020, 1987, 1957, 1928, 1900, 1874, 1849, 1825, 1802, 1780, 1759, 1739, 1719, 1701, 1719, 1701, 1725, 1707, 1691, 1675, 1659, 1644, 1630, 1616, 1603, 1590, 1577, 1565, 1553, 1541, 1530, 1519, 1493, 1483, 1473, 1464, 1454, 1445, 1437, 1428, 1420, 1412, 1404, 1396, 1389, 1381, 1374, 1367, 1361, 1354, 1348, 1341, 1335, 1329, 1323, 1317, 1312, 1306, 1301, 1295, 1290, 1285, 1280, 1275, 1262, 1258, 1254, 1251, 1247, 1278, 1274, 1271, 1267, 1263, 1260, 1256, 1253, 1249, 1246, 1243, 1239, 1236, 1233, 1230, 1227, 1223, 1220, 1217, 1214, 1211, 1208, 1205, 1203, 1200, 1197, 1194, 1192, 1189, 1186, 1184, 1181, 1179, 1176, 1174, 1172, 1169, 1167, 1165, 1163, 1160, 1158, 1156, 1154, 1152, 1150, 1148, 1146, 1144, 1142, 1140, 1138, 1136, 1165, 1163, 1159, 1157, 1155, 1153, 1151, 1149, 1147, 1145, 1143, 1142, 1140, 1138, 1136, 1135, 1133, 1131, 1130, 1128, 1127, 1125, 1123, 1122, 1120, 1119, 1117, 1116, 1114, 1113, 1111, 1110, 1108, 1107, 1106, 1104, 1103, 1102, 1100, 1099, 1098, 1096, 1095, 1094, 1093, 1091, 1090, 1089, 1088, 1086, 1085, 1084, 1083, 1113, 1111, 1110, 1109, 1108, 1106, 1105, 1104, 1103, 1102, 1100, 1099, 1098, 1097, 1096, 1095, 1094, 1092, 1091, 1090, 1089, 1088, 1087, 1086, 1085, 1084, 1083, 1082, 1081, 1080, 1079, 1078, 1077, 1076, 1075, 1074, 1073, 1072, 1071, 1070, 1069, 1068, 1068, 1067, 1066, 1065, 1064, 1063, 1062, 1061, 1061, 1060, 1059, 1058, 1057, 1056, 1056, 1055, 1054, 1053, 1052, 1052, 1051, 1050, 1049, 1049, 1048, 1047, 1046, 1046, 1045, 1044, 1043, 1043, 1042, 1041, 1041, 1040, 1039, 1038, 1038, 1037, 1036, 1036, 1035, 1034, 1034, 1033, 1032, 1052, 1052, 1051, 1050, 1050, 1049, 1048, 1048, 1047, 1046, 1046, 1045, 1044, 1044, 1043, 1042, 1042, 1041, 1041, 1040, 1039, 1039, 1038, 1038, 1037, 1036, 1036, 1035, 1035, 1034, 1034, 1033, 1032, 1032, 1031, 1031, 1030, 1030, 1029, 1029, 1028, 1028, 1027, 1027, 1026, 1025, 1025, 1024, 1024, 1023, 1023, 1022, 1022, 1021, 1021, 1020, 1020, 1019, 1019, 1019, 1018, 1018, 1017, 1017, 1016, 1016, 1015, 1014, 1014, 1013, 1012, 1012, 1011, 1010, 1010, 1009, 1008, 1008, 1007, 1006, 1006, 1005, 1004, 1004, 1003, 1002, 1002, 1001, 1001, 1000, 999, 999, 998, 997, 997, 996, 996, 995, 994, 994, 1007, 1007, 1006, 1006, 1005, 1004, 1004, 1003, 1002, 1002, 1001, 1001, 1000, 999, 999, 998, 998, 997, 996, 996, 995, 995, 994, 993, 993, 992, 992, 991, 991, 990, 990, 989, 988, 988, 987, 987, 986, 986, 985, 985, 984, 984, 983, 982, 982, 981, 981, 980, 980, 979, 979, 978, 978, 977, 977, 976, 976, 975, 975, 974, 974, 973, 973, 972, 972, 971, 971, 970, 970, 970, 969, 969, 968, 968, 967, 967, 966, 966, 965, 965, 965, 964, 964, 963, 963, 962, 962, 961, 961, 961, 960, 960, 959, 959, 958, 958, 958, 957, 957, 956]
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
    #cpRT2 = capexRT2[range]*fator[range] + ((capexRT2[range]*fator[range]-capexRT2[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    #cpRT3 = capexRT3[range]*fator[range] + ((capexRT3[range]*fator[range]-capexRT3[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    #cpRT4 = capexRT4[range]*fator[range] + ((capexRT4[range]*fator[range]-capexRT4[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    #cpRT5 = capexRT1[range]*fator[range] + ((capexRT1[range]*fator[range]-capexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    
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
    fator = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
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
    opexRT1 = [0, 8837, 4460, 3012, 2288, 2017, 1701, 1475, 1306, 1174, 1069, 982, 911, 850, 792, 747, 708, 668, 637, 610, 585, 562, 542, 523, 506, 490, 475, 461, 449, 437, 426, 416, 406, 397, 388, 380, 373, 366, 359, 352, 346, 340, 335, 329, 324, 320, 315, 310, 306, 302, 298, 301, 297, 295, 291, 288, 285, 282, 279, 276, 273, 270, 267, 265, 262, 260, 257, 255, 253, 253, 251, 249, 247, 245, 243, 241, 239, 238, 236, 234, 233, 231, 229, 228, 226, 225, 223, 222, 220, 219, 218, 217, 215, 214, 213, 212, 211, 210, 208, 209, 208, 210, 209, 208, 207, 206, 206, 205, 204, 203, 203, 202, 201, 200, 199, 198, 198, 197, 196, 195, 195, 194, 193, 192, 192, 191, 190, 190, 189, 188, 188, 187, 187, 186, 185, 185, 184, 184, 184, 184, 183, 183, 182, 182, 181, 181, 180, 180, 179, 179, 178, 178, 177, 177, 176, 176, 175, 175, 175, 175, 174, 174, 173, 173, 173, 172, 172, 171, 171, 171, 170, 170, 169, 169, 169, 168, 168, 168, 167, 167, 167, 166, 166, 166, 165, 165, 165, 165, 164, 164, 164, 163, 163, 163, 163, 162, 162, 163, 163, 163, 162, 162, 162, 158, 158, 158, 158, 157, 157, 157, 157, 156, 156, 156, 156, 156, 155, 155, 155, 155, 155, 154, 154, 154, 154, 154, 153, 153, 153, 154, 154, 154, 154, 153, 153, 153, 153, 153, 152, 152, 152, 152, 152, 152, 151, 151, 151, 151, 151, 151, 150, 150, 150, 150, 150, 150, 149, 149, 149, 149, 149, 149, 148, 148, 148, 148, 148, 148, 148, 147, 147, 147, 147, 147, 147, 147, 146, 146, 146, 146, 146, 146, 146, 146, 145, 145, 145, 145, 145, 145, 145, 145, 145, 144, 144, 145, 145, 144, 144, 144, 144, 150, 150, 150, 150, 150, 150, 150, 149, 149, 149, 149, 149, 149, 149, 149, 149, 148, 148, 148, 148, 148, 148, 148, 148, 148, 147, 147, 147, 147, 147, 147, 147, 147, 147, 146, 146, 146, 146, 146, 146, 146, 146, 146, 146, 146, 145, 145, 145, 145, 145, 145, 145, 145, 145, 145, 145, 144, 144, 144, 144, 144, 144, 144, 144, 144, 144, 144, 144, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 141, 141, 141, 141, 141, 145, 145, 144, 144, 139, 139, 139, 139, 139, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 137, 137, 137, 137, 137, 137, 137, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 135, 135, 135, 135, 135]
    opexRT1 = [0, 3417, 4460, 3012, 2288, 2017, 1701, 1475, 1306, 1174, 1069, 982, 911, 850, 792, 747, 708, 668, 637, 610, 585, 562, 542, 523, 506, 490, 475, 461, 449, 437, 426, 416, 406, 397, 388, 380, 373, 366, 359, 352, 346, 340, 335, 329, 324, 320, 315, 310, 306, 302, 298, 301, 297, 295, 291, 288, 285, 282, 279, 276, 273, 270, 267, 265, 262, 260, 257, 255, 253, 253, 251, 249, 247, 245, 243, 241, 239, 238, 236, 234, 233, 231, 229, 228, 226, 225, 223, 222, 220, 219, 218, 217, 215, 214, 213, 212, 211, 210, 208, 209, 208, 210, 209, 208, 207, 206, 206, 205, 204, 203, 203, 202, 201, 200, 199, 198, 198, 197, 196, 195, 195, 194, 193, 192, 192, 191, 190, 190, 189, 188, 188, 187, 187, 186, 185, 185, 184, 184, 184, 184, 183, 183, 182, 182, 181, 181, 180, 180, 179, 179, 178, 178, 177, 177, 176, 176, 175, 175, 175, 175, 174, 174, 173, 173, 173, 172, 172, 171, 171, 171, 170, 170, 169, 169, 169, 168, 168, 168, 167, 167, 167, 166, 166, 166, 165, 165, 165, 165, 164, 164, 164, 163, 163, 163, 163, 162, 162, 163, 163, 163, 162, 162, 162, 158, 158, 158, 158, 157, 157, 157, 157, 156, 156, 156, 156, 156, 155, 155, 155, 155, 155, 154, 154, 154, 154, 154, 153, 153, 153, 154, 154, 154, 154, 153, 153, 153, 153, 153, 152, 152, 152, 152, 152, 152, 151, 151, 151, 151, 151, 151, 150, 150, 150, 150, 150, 150, 149, 149, 149, 149, 149, 149, 148, 148, 148, 148, 148, 148, 148, 147, 147, 147, 147, 147, 147, 147, 146, 146, 146, 146, 146, 146, 146, 146, 145, 145, 145, 145, 145, 145, 145, 145, 145, 144, 144, 145, 145, 144, 144, 144, 144, 150, 150, 150, 150, 150, 150, 150, 149, 149, 149, 149, 149, 149, 149, 149, 149, 148, 148, 148, 148, 148, 148, 148, 148, 148, 147, 147, 147, 147, 147, 147, 147, 147, 147, 146, 146, 146, 146, 146, 146, 146, 146, 146, 146, 146, 145, 145, 145, 145, 145, 145, 145, 145, 145, 145, 145, 144, 144, 144, 144, 144, 144, 144, 144, 144, 144, 144, 144, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 141, 141, 141, 141, 141, 145, 145, 144, 144, 139, 139, 139, 139, 139, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 137, 137, 137, 137, 137, 137, 137, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 135, 135, 135, 135, 135]
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
    #opRT2 = opexRT2[range]*fator[range] + ((opexRT2[range]*fator[range]-opexRT2[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    #opRT3 = opexRT3[range]*fator[range] + ((opexRT3[range]*fator[range]-opexRT3[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    #opRT4 = opexRT4[range]*fator[range] + ((opexRT4[range]*fator[range]-opexRT4[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    #opRT5 = opexRT1[range]*fator[range] + ((opexRT1[range]*fator[range]-opexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    
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

def getExistentLandfill(data):
    existentlandfill = list()
    for d in data:
        if d["existent-landfill"]:
            existentlandfill.append(d["name"])
    return existentlandfill

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
        MAX_CITIES = int(sys.argv[3])           # The maximium number of allowed cities to generate the combinations due to performance issues
        TRASH_THRESHOLD = float(sys.argv[4])    # The minimun of trash for a sub-array
        REPORTFILE = sys.argv[5]                # The report file name
        OUTPUTFILE = sys.argv[6]                # The output file name
    except IndexError:
        raise SystemExit(f"Usage: {sys.argv[0]} <cities.csv> <distance.csv> <max cities> <trash threshold> <report.txt> <output.csv>")

    # Output files
    report = open(REPORTFILE, "w")
    output = open(OUTPUTFILE, "w")

    # Print parameters in report file
    report.write("============= PARÂMETROS ============= \n")
    report.write("Arquivo de cidades: " + repr(CSVCITIES) + "\n")
    report.write("Arquivo de distâncias: " + repr(CSVDISTANCE) + "\n")
    report.write("Máximo de cidades: " + repr(MAX_CITIES) + "\n")
    report.write("Quantidade de lixo mínimo para um sub-arranjo: " + repr(TRASH_THRESHOLD) + "\n\n\n")

    # RSU
    #rsutrash = [0,25,75,150,250,350,700,1250,2500,5000]
    rsutrash = [0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100, 105, 110, 115, 120, 125, 130, 135, 140, 145, 150, 155, 160, 165, 170, 175, 180, 185, 190, 195, 200, 205, 210, 215, 220, 225, 230, 235, 240, 245, 250, 255, 260, 265, 270, 275, 280, 285, 290, 295, 300, 305, 310, 315, 320, 325, 330, 335, 340, 345, 350, 355, 360, 365, 370, 375, 380, 385, 390, 395, 400, 405, 410, 415, 420, 425, 430, 435, 440, 445, 450, 455, 460, 465, 470, 475, 480, 485, 490, 495, 500, 505, 510, 515, 520, 525, 530, 535, 540, 545, 550, 555, 560, 565, 570, 575, 580, 585, 590, 595, 600, 605, 610, 615, 620, 625, 630, 635, 640, 645, 650, 655, 660, 665, 670, 675, 680, 685, 690, 695, 700, 705, 710, 715, 720, 725, 730, 735, 740, 745, 750, 755, 760, 765, 770, 775, 780, 785, 790, 795, 800, 805, 810, 815, 820, 825, 830, 835, 840, 845, 850, 855, 860, 865, 870, 875, 880, 885, 890, 895, 900, 905, 910, 915, 920, 925, 930, 935, 940, 945, 950, 955, 960, 965, 970, 975, 980, 985, 990, 995, 1000, 1005, 1010, 1015, 1020, 1025, 1030, 1035, 1040, 1045, 1050, 1055, 1060, 1065, 1070, 1075, 1080, 1085, 1090, 1095, 1100, 1105, 1110, 1115, 1120, 1125, 1130, 1135, 1140, 1145, 1150, 1155, 1160, 1165, 1170, 1175, 1180, 1185, 1190, 1195, 1200, 1205, 1210, 1215, 1220, 1225, 1230, 1235, 1240, 1245, 1250, 1255, 1260, 1265, 1270, 1275, 1280, 1285, 1290, 1295, 1300, 1305, 1310, 1315, 1320, 1325, 1330, 1335, 1340, 1345, 1350, 1355, 1360, 1365, 1370, 1375, 1380, 1385, 1390, 1395, 1400, 1405, 1410, 1415, 1420, 1425, 1430, 1435, 1440, 1445, 1450, 1455, 1460, 1465, 1470, 1475, 1480, 1485, 1490, 1495, 1500, 1505, 1510, 1515, 1520, 1525, 1530, 1535, 1540, 1545, 1550, 1555, 1560, 1565, 1570, 1575, 1580, 1585, 1590, 1595, 1600, 1605, 1610, 1615, 1620, 1625, 1630, 1635, 1640, 1645, 1650, 1655, 1660, 1665, 1670, 1675, 1680, 1685, 1690, 1695, 1700, 1705, 1710, 1715, 1720, 1725, 1730, 1735, 1740, 1745, 1750, 1755, 1760, 1765, 1770, 1775, 1780, 1785, 1790, 1795, 1800, 1805, 1810, 1815, 1820, 1825, 1830, 1835, 1840, 1845, 1850, 1855, 1860, 1865, 1870, 1875, 1880, 1885, 1890, 1895, 1900, 1905, 1910, 1915, 1920, 1925, 1930, 1935, 1940, 1945, 1950, 1955, 1960, 1965, 1970, 1975, 1980, 1985, 1990, 1995, 2000, 2005, 2010, 2015, 2020, 2025, 2030, 2035, 2040, 2045, 2050, 2055, 2060, 2065, 2070, 2075, 2080, 2085, 2090, 2095, 2100, 2105, 2110, 2115, 2120, 2125, 2130, 2135, 2140, 2145, 2150, 2155, 2160, 2165, 2170, 2175, 2180, 2185, 2190, 2195, 2200, 2205, 2210, 2215, 2220, 2225, 2230, 2235, 2240, 2245, 2250, 2255, 2260, 2265, 2270, 2275, 2280, 2285, 2290, 2295, 2300, 2305, 2310, 2315, 2320, 2325, 2330, 2335, 2340, 2345, 2350, 2355, 2360, 2365, 2370, 2375, 2380, 2385, 2390, 2395, 2400, 2405, 2410, 2415, 2420, 2425, 2430, 2435, 2440, 2445, 2450, 2455, 2460, 2465, 2470, 2475, 2480, 2485, 2490, 2495, 2500]
    
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
            if row["Aterro Pontencial"] == "Sim":
                city["landfill"] = True
            else:
                city["landfill"] = False
            if row["Aterro Existente"] == "Sim":
                city["existent-landfill"] = True
            else:
                city["existent-landfill"] = False
            city["conventional-cost"] = float(row["Custo de Coleta Mista Convencional"])
            city["transshipment-cost"] = float(row["Custo de Coleta e Transbordo de Resíduos Mistos"])
            city["cost-post-transhipment"] = float(row["Custo de Transporte Pós Transbordo"])
            citiesdata.append(city)
            line_count += 1

    # Read distances from CSV file
    distance = np.loadtxt(open(CSVDISTANCE, "rb"), delimiter=",", skiprows=0)

    if len(citiesdata) > MAX_CITIES:
        logging.info("Gerando clusters...")
        logging.debug("Quantidade de cidades superior a %d o algoritmo irá clusterizar as cidades.", MAX_CITIES)
    # Call clusterization to reduce the number of cities or just to build a list of list
    clusters = clusterization(citiesdata, distance, MAX_CITIES)

    # Print landfills in the report file
    report.write("============= ATERROS POTENCIAIS ============= \n")
    landfill = getLandfill(citiesdata)
    for l in landfill:
        report.write(l + "\n")
    report.write("\n\n\n")

    report.write("============= ATERROS EXISTENTES ============= \n")
    existentlandfill = getExistentLandfill(citiesdata)
    for l in existentlandfill:
        report.write(l + "\n")
    report.write("\n\n\n")

    # Print clusters in the report file
    report.write("============= CLUSTERS ============= \n")
    i = 1 
    for c in clusters:
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
        if current % (len(new_comb)/10.0) == 0:
            logging.info("Progreso: %d%%", current/len(new_comb)*100)

        trashArray = 0
        capexOpexArray = 0
        inboundArray = 0
        outboundArray = 0
        outboundExistentLandfill = 0
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
            rsinout = inboundoutbound(citiesdata, distance, y, centralizado, utvrs, landfill, existentlandfill)
            rsinout["capex"] =  0#capexSubArray
            rsinout["opex"] = 0#opexSubArray
            rsinout["tecnologia"] = 0#capexOpexValue
            rsinout["capex"] =  capexSubArray
            rsinout["opex"] = opexSubArray
            rsinout["tecnologia"] = capexOpexValue
            
            inboundArray = inboundArray + (rsinout["inbound"] * trashSubArray)
            outboundArray = outboundArray + (rsinout["outbound"] * trashSubArray)
            outboundExistentLandfill = outboundExistentLandfill + (rsinout["outbound-existente"] * trashSubArray)
            rsinout["lixo"] = trashSubArray
            rsinout["total"] = capexOpexValue + rsinout["inbound"] + rsinout["outbound"]
            sub.append(rsinout)
            
        cpopfinalValue = 0 #capexOpexArray/trashArray
        cpopfinalValue = capexOpexArray/trashArray
        
        new["arranjo"] = i
        new["sub"] = sub
        new["capexopex"] = cpopfinalValue
        new["lixo-array"] = trashArray
        new["inbound"] = inboundArray/trashArray
        new["outbound"] = outboundArray/trashArray
        new["outbound-existente"] = outboundExistentLandfill/trashArray
        new["total"] = cpopfinalValue + (inboundArray/trashArray) + (outboundArray/trashArray)
        data.append(new)

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

            output.write(repr(d["arranjo"]) + ";Sumário;NA;NA;NA;" + repr(d["total"]) + ";" + repr(d["lixo-array"]) + ";" + repr(d["capexopex"]) + ";" + repr(d["inbound"]) + ";" + repr(d["outbound"]) + ";" + repr(d["outbound-existente"]) + "\n")


            for x in range(len(d["sub"])):
                output.write(repr(d["arranjo"]) + ";" + repr(d["sub"][x]["sub-arranjo"]) + ";" + repr(d["sub"][x]["aterro"]) + ";" + repr(d["sub"][x]["aterro-existente"]) + ";" + repr(d["sub"][x]["utvr"]) + ";" + repr(d["sub"][x]["total"]) + ";" + repr(d["sub"][x]["lixo"]) + ";" + repr(d["sub"][x]["tecnologia"]) + ";" + repr(d["sub"][x]["inbound"])  + ";" + repr(d["sub"][x]["outbound"]) + ";" + repr(d["sub"][x]["outbound-existente"]) + "\n")

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
        output.write(repr(data[i]["arranjo"]) + ";Sumário;NA;NA;NA;" + repr(data[i]["total"]) + ";" + repr(data[i]["lixo-array"]) + ";" + repr(data[i]["capexopex"]) + ";" + repr(data[i]["inbound"])  + ";" + repr(data[i]["outbound"]) + ";" + repr(data[i]["outbound-existente"]) + "\n")

        report.write(repr(i+1) + ".\t" + repr(data[i]["arranjo"]) + "\n")
        report.write("- Lixo: " + repr(data[i]["lixo-array"]) + "\n")
        report.write("- Custo Total: " + repr(data[i]["total"]) + "\n")
        report.write("-- Inbound: " + repr(data[i]["inbound"]) + "\n")
        report.write("-- Tecnologia: " + repr(data[i]["capexopex"]) + "\n")
        report.write("-- Outbound: " + repr(data[i]["outbound"]) + "\n\n")
        report.write("-- Sub-arranjos:\n")
        for x in range(len(data[i]["sub"])):
            output.write(repr(data[i]["arranjo"]) + ";" + repr(data[i]["sub"][x]["sub-arranjo"]) + ";" + repr(data[i]["sub"][x]["aterro"]) + ";" + repr(data[i]["sub"][x]["aterro-existente"]) + ";" + repr(data[i]["sub"][x]["utvr"]) + ";" + repr(data[i]["sub"][x]["total"]) + ";" + repr(data[i]["sub"][x]["lixo"]) + ";" + repr(data[i]["sub"][x]["tecnologia"]) + ";" + repr(data[i]["sub"][x]["inbound"])  + ";" + repr(data[i]["sub"][x]["outbound"]) + ";" + repr(data[i]["sub"][x]["outbound-existente"]) + "\n")

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