import sys
import logging
import numpy as np
import csv
import copy

# Global variables
ROUND = 10

def sorted_k_partitions(seq, k):
    """Returns a list of all unique k-partitions of `seq`.

    Each partition is a list of parts, and each part is a tuple.

    The parts in each individual partition will be sorted in shortlex
    order (i.e., by length first, then lexicographically).

    The overall list of partitions will then be sorted by the length
    of their first part, the length of their second part, ...,
    the length of their last part, and then lexicographically.
    """
    n = len(seq)
    groups = []  # a list of lists, currently empty

    def generate_partitions(i):
        if i >= n:
            yield list(map(tuple, groups))
        else:
            if n - i > k - len(groups):
                for group in groups:
                    group.append(seq[i])
                    yield from generate_partitions(i + 1)
                    group.pop()

            if len(groups) < k:
                groups.append([seq[i]])
                yield from generate_partitions(i + 1)
                groups.pop()

    result = generate_partitions(0)

    # Sort the parts in each partition in shortlex order
    #result = [sorted(ps, key = lambda p: (len(p), p)) for ps in result]
    # Sort partitions by the length of each part, then lexicographically.
    #result = sorted(result, key = lambda ps: (*map(len, ps), ps))

    return result

def clusterization(data, distance, max):
    clusters = []
    for d in data:
        city = []
        city.append(d)
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
                trash_line = getSubArrayTrash(data, c)
            if cities_temp[column] in c:
                trash_column = getSubArrayTrash(data, c)
        centrodemassa = ""
        outracidade = ""

        logging.debug("A menor distância é %f vou unir as cidades %s (%f) e %s (%f)", min, cities_temp[line], trash_line, cities_temp[column], trash_column)
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

def getDistanceBetweenCites(data, distance, cityA, cityB):
    return round(distance[data[cityA]["position"]][data[cityB]["position"]], ROUND)

def getSubArrayRSURange(sumTrash, rsutrash):
    for i in range(len(rsutrash)):
        if sumTrash >= rsutrash[i] and sumTrash <= rsutrash[i+1]:
            return i

def removeArraysWithoutUTVR(combinations, utvrs):
    comb = combinations.copy()
    for c in range(len(comb)):
        if c % (round(len(comb)+1/10.0)) == 0:
            logging.info("Progresso: %d%%", c/len(comb)*100)
        for sub in comb[c]:
            find = False
            for cluster in sub:
                for city in cluster:
                    if city in utvrs:
                        find = True
            if not find:
                combinations.remove(comb[c])
                break
    logging.info("Progresso: 100%")
    return combinations

def removeArraysTrashThreshold(data, combinations, threshold):
    comb = combinations.copy()
    for c in range(len(comb)):
        if c % (round(len(comb)+1/10.0)) == 0:
            logging.info("Progresso: %d%%", c/len(comb)*100)
        for sub in comb[c]:
            trash = 0
            for cluster in sub:
                for city in cluster:
                    trash = trash + data[city]["trash"]
            if trash < threshold:
                combinations.remove(comb[c])
                break        
    logging.info("Progresso: 100%")
    return combinations

def funccentrodemassa(data, cluster, utvrs):
    max = 0
    c_centrodemassa = ""
    for i in cluster:
        trash = data[i]["trash"]
        if trash > max and i in utvrs:
            max = trash
            c_centrodemassa = i
    return c_centrodemassa

def calculateInboundOutbound(cdata, distance, subarray, isCentralized, utvrs_only, aterros_only, existentlandfill, CAPEX_INBOUND, CAPEX_OUTBOUND, PAYMENT_PERIOD, MOVIMENTATION_COST, LANDFILL_DEVIATION, totalTrash, ADDITIONAL_COST):
    data = []
    subArrayTrash = getSubArrayTrash(cdata, subarray)
    for utvr_city in subarray:
        entry = {}
        sum_inbound = 0
        if utvr_city in utvrs_only:
            #logging.debug("%s é uma UTVR...", utvr_city)
            entry["sub-arranjo"] = subarray
            entry["utvr"] = utvr_city
            for other_city in subarray:
                conventional_cost = cdata[other_city]["conventional-cost"]
                transshipment_cost = cdata[other_city]["transshipment-cost"]
                cost_post_transhipment = cdata[other_city]["cost-post-transhipment"]
                trash = cdata[other_city]["trash"]
                sum_inbound = sum_inbound + ((conventional_cost) + (transshipment_cost) + (cost_post_transhipment * getDistanceBetweenCites(cdata, distance, other_city, utvr_city))) * trash
                #sum_co2 = sum_co2 + (1.24 * getDistanceBetweenCites(cdata, distance, utvr_city, other_city * trash))

            #sum_co2 = sum_co2 / getSubTrash(cdata, subarray)
            entry["inbound-sum"] = sum_inbound
            entry["inbound-show"] = round(sum_inbound/totalTrash, 3)

            if isCentralized:
                sum_inbound = round(sum_inbound * ADDITIONAL_COST / subArrayTrash, 3)
            else:
                sum_inbound = round(sum_inbound / subArrayTrash, 3)

            entry["inbound"] = sum_inbound

            dist = 999999
            for a in existentlandfill:
                distCities = getDistanceBetweenCites(cdata, distance, utvr_city, a)
                if distCities < dist:
                    dist = distCities
                    sum_outbound = 0
                    if isCentralized:
                        entry['outbound-existent-landfill'] = sum_outbound + round((distCities * (MOVIMENTATION_COST * cdata[utvr_city]["cost-post-transhipment"])) * LANDFILL_DEVIATION, ROUND)
                    else:
                        entry['outbound-existent-landfill'] = sum_outbound + round((distCities * (MOVIMENTATION_COST * cdata[utvr_city]["cost-post-transhipment"])) * LANDFILL_DEVIATION * ADDITIONAL_COST, ROUND)
                    entry["aterro-existente"] = a
        
            for a in aterros_only:
                e = copy.deepcopy(entry)
                sum_outbound = 0
                #sum_outbound = sum_outbound + (getDistanceBetweenCites(cdata, distance, utvr_city,a) * (0.7 * cdata[utvr_city]["cost-post-transhipment"])) * LANDFILL_DEVIATION
                if isCentralized:
                    e["outbound"] = round(((subArrayTrash * LANDFILL_DEVIATION)*(cdata[utvr_city]["cost-post-transhipment"]*getDistanceBetweenCites(cdata, distance, utvr_city,a)*MOVIMENTATION_COST))/subArrayTrash, ROUND)
                else:
                    e["outbound"] = round(((subArrayTrash * LANDFILL_DEVIATION)*(cdata[utvr_city]["cost-post-transhipment"]*getDistanceBetweenCites(cdata, distance, utvr_city,a)*MOVIMENTATION_COST)) * ADDITIONAL_COST/subArrayTrash, ROUND)
                e["aterro"] = a
                e["total"] = e["inbound"] + e["outbound"]
                
                #logging.debug("Adicionando: %s", e)
                data.append(e)
    
    data = sorted(data, key = lambda k: (k["total"]))

    if isCentralized:
        #Retorna a UTVR sendo o centro de massa, não o mais eficaz
        cmassa = funccentrodemassa(cdata, subarray, utvrs_only)
        for d in data:
            if d["utvr"] == cmassa:
                return d
    else:
        return data[0]

def getSubArrayCapex(range, trashSum, rsutrash):
    fator = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
    #capexRT1 = [0, 22893, 11568, 7876, 6029, 7044, 6141, 5499, 5018, 4646, 4350, 4109, 3909, 3741, 3598, 3475, 3368, 3274, 3192, 3119, 3054, 2953, 2855, 2765, 2683, 2608, 2538, 2473, 2413, 2358, 2305, 2257, 2211, 2168, 2128, 2090, 2054, 2020, 1987, 1957, 1928, 1900, 1874, 1849, 1825, 1802, 1780, 1759, 1739, 1719, 1701, 1719, 1701, 1725, 1707, 1691, 1675, 1659, 1644, 1630, 1616, 1603, 1590, 1577, 1565, 1553, 1541, 1530, 1519, 1493, 1483, 1473, 1464, 1454, 1445, 1437, 1428, 1420, 1412, 1404, 1396, 1389, 1381, 1374, 1367, 1361, 1354, 1348, 1341, 1335, 1329, 1323, 1317, 1312, 1306, 1301, 1295, 1290, 1285, 1280, 1275, 1262, 1258, 1254, 1251, 1247, 1278, 1274, 1271, 1267, 1263, 1260, 1256, 1253, 1249, 1246, 1243, 1239, 1236, 1233, 1230, 1227, 1223, 1220, 1217, 1214, 1211, 1208, 1205, 1203, 1200, 1197, 1194, 1192, 1189, 1186, 1184, 1181, 1179, 1176, 1174, 1172, 1169, 1167, 1165, 1163, 1160, 1158, 1156, 1154, 1152, 1150, 1148, 1146, 1144, 1142, 1140, 1138, 1136, 1165, 1163, 1159, 1157, 1155, 1153, 1151, 1149, 1147, 1145, 1143, 1142, 1140, 1138, 1136, 1135, 1133, 1131, 1130, 1128, 1127, 1125, 1123, 1122, 1120, 1119, 1117, 1116, 1114, 1113, 1111, 1110, 1108, 1107, 1106, 1104, 1103, 1102, 1100, 1099, 1098, 1096, 1095, 1094, 1093, 1091, 1090, 1089, 1088, 1086, 1085, 1084, 1083, 1113, 1111, 1110, 1109, 1108, 1106, 1105, 1104, 1103, 1102, 1100, 1099, 1098, 1097, 1096, 1095, 1094, 1092, 1091, 1090, 1089, 1088, 1087, 1086, 1085, 1084, 1083, 1082, 1081, 1080, 1079, 1078, 1077, 1076, 1075, 1074, 1073, 1072, 1071, 1070, 1069, 1068, 1068, 1067, 1066, 1065, 1064, 1063, 1062, 1061, 1061, 1060, 1059, 1058, 1057, 1056, 1056, 1055, 1054, 1053, 1052, 1052, 1051, 1050, 1049, 1049, 1048, 1047, 1046, 1046, 1045, 1044, 1043, 1043, 1042, 1041, 1041, 1040, 1039, 1038, 1038, 1037, 1036, 1036, 1035, 1034, 1034, 1033, 1032, 1052, 1052, 1051, 1050, 1050, 1049, 1048, 1048, 1047, 1046, 1046, 1045, 1044, 1044, 1043, 1042, 1042, 1041, 1041, 1040, 1039, 1039, 1038, 1038, 1037, 1036, 1036, 1035, 1035, 1034, 1034, 1033, 1032, 1032, 1031, 1031, 1030, 1030, 1029, 1029, 1028, 1028, 1027, 1027, 1026, 1025, 1025, 1024, 1024, 1023, 1023, 1022, 1022, 1021, 1021, 1020, 1020, 1019, 1019, 1019, 1018, 1018, 1017, 1017, 1016, 1016, 1015, 1014, 1014, 1013, 1012, 1012, 1011, 1010, 1010, 1009, 1008, 1008, 1007, 1006, 1006, 1005, 1004, 1004, 1003, 1002, 1002, 1001, 1001, 1000, 999, 999, 998, 997, 997, 996, 996, 995, 994, 994, 1007, 1007, 1006, 1006, 1005, 1004, 1004, 1003, 1002, 1002, 1001, 1001, 1000, 999, 999, 998, 998, 997, 996, 996, 995, 995, 994, 993, 993, 992, 992, 991, 991, 990, 990, 989, 988, 988, 987, 987, 986, 986, 985, 985, 984, 984, 983, 982, 982, 981, 981, 980, 980, 979, 979, 978, 978, 977, 977, 976, 976, 975, 975, 974, 974, 973, 973, 972, 972, 971, 971, 970, 970, 970, 969, 969, 968, 968, 967, 967, 966, 966, 965, 965, 965, 964, 964, 963, 963, 962, 962, 961, 961, 961, 960, 960, 959, 959, 958, 958, 958, 957, 957, 956]
    capexRT1 = [105300,	38717,	20334,	14235,	11184,	10340,	8963,	7978,	7240,	6665,	6204,	5827,	5512,	5246,	5017,	4819,	4646,	4492,	4356,	4234,	4124,	4024,	3933,	3849,	3773,	3702,	3637,	3576,	3520,	3467,	3418,	3372,	3329,	3288,	3250,	3213,	3179,	3146,	3115,	3086,	3058,	3031,	3005,	2980,	2957,	2934,	2913,	2892,	2872,	2853,	2834,	2838,	2820,	2820,	2803,	2787,	2771,	2756,	2741,	2727,	2713,	2700,	2686,	2674,	2661,	2649,	2638,	2626,	2615,	2605,	2594,	2584,	2573,	2563,	2554,	2544,	2535,	2526,	2517,	2508,	2499,	2488,	2480,	2472,	2464,	2456,	2448,	2441,	2434,	2426,	2419,	2412,	2405,	2399,	2392,	2385,	2379,	2373,	2366,	2364,	2360,	2350,	2346,	2343,	2339,	2335,	2346,	2342,	2338,	2335,	2332,	2328,	2325,	2322,	2318,	2315,	2312,	2309,	2306,	2303,	2300,	2297,	2294,	2292,	2289,	2286,	2283,	2281,	2278,	2276,	2273,	2270,	2268,	2266,	2263,	2261,	2258,	2256,	2254,	2252,	2250,	2248,	2247,	2245,	2243,	2241,	2240,	2238,	2236,	2235,	2233,	2232,	2230,	2229,	2227,	2226,	2224,	2223,	2221,	2233,	2231,	2229,	2228,	2226,	2225,	2224,	2222,	2221,	2220,	2219,	2217,	2216,	2215,	2214,	2213,	2211,	2210,	2209,	2208,	2207,	2206,	2205,	2204,	2203,	2202,	2201,	2200,	2199,	2198,	2197,	2196,	2195,	2194,	2193,	2192,	2191,	2190,	2191,	2190,	2189,	2188,	2187,	2186,	2035,	2035,	2035,	2035,	2035,	2035,	2035,	2035,	2035,	2048,	2048,	2048,	2048,	2048,	2048,	2048,	2048,	2048,	2048,	2049,	2049,	2049,	2049,	2049,	2049,	2049,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2047,	2048,	2048,	2048,	2048,	2048,	2048,	2048,	2048,	2048,	2048,	2048,	2049,	2049,	2049,	2049,	2049,	2049,	2049,	2050,	2050,	2050,	2050,	2050,	2050,	2051,	2051,	2051,	2051,	2051,	2052,	2052,	2052,	2052,	2053,	2053,	2053,	2053,	2054,	2054,	2054,	2054,	2055,	2055,	2055,	2055,	2056,	2057,	2057,	2057,	2057,	2058,	2058,	2071,	2071,	2071,	2071,	2072,	2072,	2072,	2073,	2073,	2073,	2074,	2074,	2074,	2075,	2075,	2075,	2076,	2076,	2076,	2077,	2077,	2077,	2078,	2078,	2078,	2079,	2079,	2079,	2080,	2080,	2081,	2081,	2081,	2082,	2082,	2082,	2083,	2083,	2084,	2084,	2084,	2085,	2085,	2086,	2086,	2087,	2087,	2087,	2088,	2088,	2089,	2089,	2090,	2090,	2090,	2091,	2091,	2092,	2092,	2093,	2093,	2094,	2094,	2094,	2095,	2095,	2096,	2096,	2097,	2097,	2097,	2098,	2098,	2099,	2099,	2100,	2100,	2101,	2101,	2101,	2102,	2102,	2103,	2103,	2104,	2104,	2104,	2105,	2105,	2106,	2106,	2107,	2107,	2108,	2108,	2109,	2109,	2109,	2110,	2110,	2121,	2121,	2122,	2122,	1815,	1816,	1817,	1819,	1820,	1821,	1822,	1823,	1824,	1826,	1827,	1828,	1829,	1830,	1832,	1833,	1834,	1835,	1836,	1838,	1839,	1840,	1841,	1842,	1844,	1845,	1846,	1847,	1848,	1850,	1851,	1852,	1853,	1854,	1856,	1857,	1858,	1859,	1860,	1862,	1863,	1864,	1865,	1867,	1868,	1869,	1870,	1872,	1873,	1874,	1875,	1876,	1878,	1879,	1880,	1882,	1883,	1884,	1886,	1887,	1888,	1890,	1891,	1892,	1894,	1895,	1897,	1898,	1899,	1901,	1902,	1903,	1905,	1906,	1907,	1909,	1910,	1912,	1913,	1914,	1916,	1917,	1918,	1920,	1921,	1923,	1924,	1925,	1927,	1928,	1929,	1931,	1932,	1934,	1935,	1936]
    cpRT1 = capexRT1[range]*fator[range] + ((capexRT1[range]*fator[range]-capexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    return(round(cpRT1, ROUND))

def getSubArrayOpex(range, trashSum, rsutrash):
    fator = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
    #opexRT1 = [0, 3417, 4460, 3012, 2288, 2017, 1701, 1475, 1306, 1174, 1069, 982, 911, 850, 792, 747, 708, 668, 637, 610, 585, 562, 542, 523, 506, 490, 475, 461, 449, 437, 426, 416, 406, 397, 388, 380, 373, 366, 359, 352, 346, 340, 335, 329, 324, 320, 315, 310, 306, 302, 298, 301, 297, 295, 291, 288, 285, 282, 279, 276, 273, 270, 267, 265, 262, 260, 257, 255, 253, 253, 251, 249, 247, 245, 243, 241, 239, 238, 236, 234, 233, 231, 229, 228, 226, 225, 223, 222, 220, 219, 218, 217, 215, 214, 213, 212, 211, 210, 208, 209, 208, 210, 209, 208, 207, 206, 206, 205, 204, 203, 203, 202, 201, 200, 199, 198, 198, 197, 196, 195, 195, 194, 193, 192, 192, 191, 190, 190, 189, 188, 188, 187, 187, 186, 185, 185, 184, 184, 184, 184, 183, 183, 182, 182, 181, 181, 180, 180, 179, 179, 178, 178, 177, 177, 176, 176, 175, 175, 175, 175, 174, 174, 173, 173, 173, 172, 172, 171, 171, 171, 170, 170, 169, 169, 169, 168, 168, 168, 167, 167, 167, 166, 166, 166, 165, 165, 165, 165, 164, 164, 164, 163, 163, 163, 163, 162, 162, 163, 163, 163, 162, 162, 162, 158, 158, 158, 158, 157, 157, 157, 157, 156, 156, 156, 156, 156, 155, 155, 155, 155, 155, 154, 154, 154, 154, 154, 153, 153, 153, 154, 154, 154, 154, 153, 153, 153, 153, 153, 152, 152, 152, 152, 152, 152, 151, 151, 151, 151, 151, 151, 150, 150, 150, 150, 150, 150, 149, 149, 149, 149, 149, 149, 148, 148, 148, 148, 148, 148, 148, 147, 147, 147, 147, 147, 147, 147, 146, 146, 146, 146, 146, 146, 146, 146, 145, 145, 145, 145, 145, 145, 145, 145, 145, 144, 144, 145, 145, 144, 144, 144, 144, 150, 150, 150, 150, 150, 150, 150, 149, 149, 149, 149, 149, 149, 149, 149, 149, 148, 148, 148, 148, 148, 148, 148, 148, 148, 147, 147, 147, 147, 147, 147, 147, 147, 147, 146, 146, 146, 146, 146, 146, 146, 146, 146, 146, 146, 145, 145, 145, 145, 145, 145, 145, 145, 145, 145, 145, 144, 144, 144, 144, 144, 144, 144, 144, 144, 144, 144, 144, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 143, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 142, 141, 141, 141, 141, 141, 145, 145, 144, 144, 139, 139, 139, 139, 139, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 138, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 137, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 137, 137, 137, 137, 137, 137, 137, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 136, 135, 135, 135, 135, 135]
    opexRT1 = [25928,	8837,	4460,	3012,	2288,	2017,	1701,	1475,	1306,	1174,	1069,	982,	911,	850,	792,	747,	708,	668,	637,	610,	585,	562,	542,	523,	506,	490,	475,	461,	449,	437,	426,	416,	406,	397,	388,	380,	373,	366,	359,	352,	346,	340,	335,	329,	324,	320,	315,	310,	306,	302,	298,	301,	297,	295,	291,	288,	285,	282,	279,	276,	273,	270,	267,	265,	262,	260,	257,	255,	253,	253,	251,	249,	247,	245,	243,	241,	239,	238,	236,	234,	233,	231,	229,	228,	226,	225,	223,	222,	220,	219,	218,	217,	215,	214,	213,	212,	211,	210,	208,	209,	208,	210,	209,	208,	207,	206,	206,	205,	204,	203,	203,	202,	201,	200,	199,	198,	198,	197,	196,	195,	195,	194,	193,	192,	192,	191,	190,	190,	189,	188,	188,	187,	187,	186,	185,	185,	184,	184,	184,	184,	183,	183,	182,	182,	181,	181,	180,	180,	179,	179,	178,	178,	177,	177,	176,	176,	175,	175,	175,	175,	174,	174,	173,	173,	173,	172,	172,	171,	171,	171,	170,	170,	169,	169,	169,	168,	168,	168,	167,	167,	167,	166,	166,	166,	165,	165,	165,	165,	164,	164,	164,	163,	163,	163,	163,	162,	162,	163,	163,	163,	162,	162,	162,	158,	158,	158,	158,	157,	157,	157,	157,	156,	156,	156,	156,	156,	155,	155,	155,	155,	155,	154,	154,	154,	154,	154,	153,	153,	153,	154,	154,	154,	154,	153,	153,	153,	153,	153,	152,	152,	152,	152,	152,	152,	151,	151,	151,	151,	151,	151,	150,	150,	150,	150,	150,	150,	149,	149,	149,	149,	149,	149,	148,	148,	148,	148,	148,	148,	148,	147,	147,	147,	147,	147,	147,	147,	146,	146,	146,	146,	146,	146,	146,	146,	145,	145,	145,	145,	145,	145,	145,	145,	145,	144,	144,	145,	145,	144,	144,	144,	144,	150,	150,	150,	150,	150,	150,	150,	149,	149,	149,	149,	149,	149,	149,	149,	149,	148,	148,	148,	148,	148,	148,	148,	148,	148,	147,	147,	147,	147,	147,	147,	147,	147,	147,	146,	146,	146,	146,	146,	146,	146,	146,	146,	146,	146,	145,	145,	145,	145,	145,	145,	145,	145,	145,	145,	145,	144,	144,	144,	144,	144,	144,	144,	144,	144,	144,	144,	144,	143,	143,	143,	143,	143,	143,	143,	143,	143,	143,	143,	143,	143,	142,	142,	142,	142,	142,	142,	142,	142,	142,	142,	142,	142,	142,	142,	141,	141,	141,	141,	141,	145,	145,	144,	144,	139,	139,	139,	139,	139,	138,	138,	138,	138,	138,	138,	138,	138,	138,	138,	138,	138,	138,	138,	138,	138,	138,	138,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	137,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	137,	137,	137,	137,	137,	137,	137,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	136,	135,	135,	135,	135,	135]
    opRT1 = opexRT1[range]*fator[range] + ((opexRT1[range]*fator[range]-opexRT1[range+1]*fator[range+1]) * ((trashSum - rsutrash[range]) / (rsutrash[range]-rsutrash[range+1])))
    return(round(opRT1, ROUND))

def getCities(data):
    cities = list()
    for d in data:
        cities.append(d)
    return cities

def getTrash(data):
    trash = list()
    for d in data:
        trash.append(data[d]["trash"])
    return trash

def getLandfill(data):
    landfill = list()
    for d in data:
        if data[d]["landfill"]:
            landfill.append(data[d]['name'])
    return landfill

def getExistentLandfill(data):
    existentlandfill = list()
    for d in data:
        if data[d]["existent-landfill"]:
            existentlandfill.append(data[d]['name'])
    return existentlandfill

def getUTVR(data):
    utvr = list()
    for d in data:
        if data[d]["utvr"]:
            utvr.append(data[d]['name'])
    return utvr

def getSubArrayTrash(data, cluster):
    total = 0
    for c in cluster:
        total = total + data[c]["trash"]
    return round(total, ROUND)

def getSubArrayPopulation(data, cluster):
    total = 0
    for c in cluster:
        total = total + data[c]["population"]
    return total

def main():
    # Configure logs
    logging.basicConfig(
        stream=sys.stderr, 
        level=logging.INFO,
        format='[%(asctime)s] {%(filename)s:%(lineno)d} %(levelname)s - %(message)s',
    )

    # Parser command line parameters
    try:
        CSVCITIES = sys.argv[1]                 # Cities file
        CSVDISTANCE = sys.argv[2]               # Distance matrix file
        MAX_CITIES = int(sys.argv[3])           # The maximum number of allowed cities to generate the combinations due to performance issues
        TRASH_THRESHOLD = float(sys.argv[4])    # The minimun of trash for a sub-array
        CAPEX_INBOUND = float(sys.argv[5])      #
        CAPEX_OUTBOUND = float(sys.argv[6])     #
        PAYMENT_PERIOD = int(sys.argv[7])       # 
        MOVIMENTATION_COST = float(sys.argv[8]) #
        LANDFILL_DEVIATION = float(sys.argv[9]) #     
        REPORTFILE = sys.argv[10]               # The report file name
        OUTPUTFILE = sys.argv[11]               # The output file name
        if len(sys.argv) > 12:
            CSVRSU = sys.argv[12]
        else:
            CSVRSU = "rsu.csv"
    except IndexError:
        raise SystemExit(f"Usage: {sys.argv[0]} <cities.csv> <distance.csv> <max cities> <trash threshold> <capex inbound> <opex inbound> <paymnent period> <movimentation cost> <landfill deviation> <report.txt> <output.csv> <rsu.cvs>")

    # Static variables
    MAX_SUB_ARRAYS = 2                          # Max sub-arrays per array
    MAX_ARRAYS = 2000                           # Top # arrays that will be exported
    VERSION = "3.1.2"                           # Algorithm Version
    ADDITIONAL_COST = 1.25                      # Inbound for centralized array and outbound for non-centralized arrays

    # Output files
    report = open(REPORTFILE, "w")
    output = open(OUTPUTFILE, "w")

    # Print parameters in report file
    report.write("============= PARÂMETROS ============= \n")
    report.write("Arquivo de cidades: " + repr(CSVCITIES) + "\n")
    report.write("Arquivo de distâncias: " + repr(CSVDISTANCE) + "\n")
    report.write("Arquivo de RSU: " + repr(CSVRSU) + "\n")
    report.write("Máximo de cidades: " + repr(MAX_CITIES) + "\n")
    report.write("Quantidade de lixo mínimo para um sub-arranjo: " + repr(TRASH_THRESHOLD) + "\n")
    report.write("Capex Inbound: " + repr(CAPEX_INBOUND) + "\n")
    report.write("Capex Outbound: " + repr(CAPEX_OUTBOUND) + "\n")
    report.write("Prazo: " + repr(PAYMENT_PERIOD) + "\n")
    report.write("Desvio Aterro (%): " + repr(LANDFILL_DEVIATION) + "\n")
    report.write("Redução Custo de Transporte (%): " + repr(MOVIMENTATION_COST) + "\n")

    # RSU
    rsutrash = [0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100, 105, 110, 115, 120, 125, 130, 135, 140, 145, 150, 155, 160, 165, 170, 175, 180, 185, 190, 195, 200, 205, 210, 215, 220, 225, 230, 235, 240, 245, 250, 255, 260, 265, 270, 275, 280, 285, 290, 295, 300, 305, 310, 315, 320, 325, 330, 335, 340, 345, 350, 355, 360, 365, 370, 375, 380, 385, 390, 395, 400, 405, 410, 415, 420, 425, 430, 435, 440, 445, 450, 455, 460, 465, 470, 475, 480, 485, 490, 495, 500, 505, 510, 515, 520, 525, 530, 535, 540, 545, 550, 555, 560, 565, 570, 575, 580, 585, 590, 595, 600, 605, 610, 615, 620, 625, 630, 635, 640, 645, 650, 655, 660, 665, 670, 675, 680, 685, 690, 695, 700, 705, 710, 715, 720, 725, 730, 735, 740, 745, 750, 755, 760, 765, 770, 775, 780, 785, 790, 795, 800, 805, 810, 815, 820, 825, 830, 835, 840, 845, 850, 855, 860, 865, 870, 875, 880, 885, 890, 895, 900, 905, 910, 915, 920, 925, 930, 935, 940, 945, 950, 955, 960, 965, 970, 975, 980, 985, 990, 995, 1000, 1005, 1010, 1015, 1020, 1025, 1030, 1035, 1040, 1045, 1050, 1055, 1060, 1065, 1070, 1075, 1080, 1085, 1090, 1095, 1100, 1105, 1110, 1115, 1120, 1125, 1130, 1135, 1140, 1145, 1150, 1155, 1160, 1165, 1170, 1175, 1180, 1185, 1190, 1195, 1200, 1205, 1210, 1215, 1220, 1225, 1230, 1235, 1240, 1245, 1250, 1255, 1260, 1265, 1270, 1275, 1280, 1285, 1290, 1295, 1300, 1305, 1310, 1315, 1320, 1325, 1330, 1335, 1340, 1345, 1350, 1355, 1360, 1365, 1370, 1375, 1380, 1385, 1390, 1395, 1400, 1405, 1410, 1415, 1420, 1425, 1430, 1435, 1440, 1445, 1450, 1455, 1460, 1465, 1470, 1475, 1480, 1485, 1490, 1495, 1500, 1505, 1510, 1515, 1520, 1525, 1530, 1535, 1540, 1545, 1550, 1555, 1560, 1565, 1570, 1575, 1580, 1585, 1590, 1595, 1600, 1605, 1610, 1615, 1620, 1625, 1630, 1635, 1640, 1645, 1650, 1655, 1660, 1665, 1670, 1675, 1680, 1685, 1690, 1695, 1700, 1705, 1710, 1715, 1720, 1725, 1730, 1735, 1740, 1745, 1750, 1755, 1760, 1765, 1770, 1775, 1780, 1785, 1790, 1795, 1800, 1805, 1810, 1815, 1820, 1825, 1830, 1835, 1840, 1845, 1850, 1855, 1860, 1865, 1870, 1875, 1880, 1885, 1890, 1895, 1900, 1905, 1910, 1915, 1920, 1925, 1930, 1935, 1940, 1945, 1950, 1955, 1960, 1965, 1970, 1975, 1980, 1985, 1990, 1995, 2000, 2005, 2010, 2015, 2020, 2025, 2030, 2035, 2040, 2045, 2050, 2055, 2060, 2065, 2070, 2075, 2080, 2085, 2090, 2095, 2100, 2105, 2110, 2115, 2120, 2125, 2130, 2135, 2140, 2145, 2150, 2155, 2160, 2165, 2170, 2175, 2180, 2185, 2190, 2195, 2200, 2205, 2210, 2215, 2220, 2225, 2230, 2235, 2240, 2245, 2250, 2255, 2260, 2265, 2270, 2275, 2280, 2285, 2290, 2295, 2300, 2305, 2310, 2315, 2320, 2325, 2330, 2335, 2340, 2345, 2350, 2355, 2360, 2365, 2370, 2375, 2380, 2385, 2390, 2395, 2400, 2405, 2410, 2415, 2420, 2425, 2430, 2435, 2440, 2445, 2450, 2455, 2460, 2465, 2470, 2475, 2480, 2485, 2490, 2495, 2500]
    
    citiesdic = {}
    totalTrash = 0
    totalPopulation = 0
    with open(CSVCITIES, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            city = {}
            city["name"] = row["Município"]
            city["position"] = line_count
            city["population"] = round(float(row["População"]), ROUND)
            city["trash"] = round(float(row["Lixo (t/d)"]), ROUND)
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
            city["conventional-cost"] = round(float(row["Custo de Coleta Mista Convencional"]), ROUND)
            city["transshipment-cost"] = round(float(row["Custo de Coleta e Transbordo de Resíduos Mistos"]), ROUND)
            city["cost-post-transhipment"] = round(float(row["Custo de Transporte Pós Transbordo"]), ROUND)
            citiesdic[city["name"]] = city
            totalTrash = totalTrash + city["trash"]
            totalPopulation = totalPopulation + city["population"]
            line_count += 1

    # Read RSU file
    #reader = csv.DictReader(open(CSVRSU, mode='r', encoding="utf8"))
    #rsu = []
    #for r in reader:
    #    rsu.append(r)
    #    print(r)

    capexInbound  = round((CAPEX_INBOUND  * 1000000)/(totalTrash * 312 * PAYMENT_PERIOD), ROUND)
    capexOutbound = round((CAPEX_OUTBOUND * 1000000)/(totalTrash * 312 * PAYMENT_PERIOD), ROUND)

    report.write("Capex Inbound: " + repr(capexInbound) + "\n")
    report.write("Capex Outbound: " + repr(capexOutbound) + "\n\n\n")

    # Read distances from CSV file
    distance = np.loadtxt(open(CSVDISTANCE, "rb"), delimiter=",", skiprows=0)

    if len(citiesdic) > MAX_CITIES:
        logging.info("Gerando clusters...")
        logging.debug("Quantidade de cidades superior a %d o algoritmo irá clusterizar as cidades.", MAX_CITIES)
    # Call clusterization to reduce the number of cities or just to build a list of list
    clusters = []
    clusters = clusterization(citiesdic, distance, MAX_CITIES)

    # Print potencial landfills in the report file
    report.write("============= ATERROS POTENCIAIS ============= \n")
    landfill = getLandfill(citiesdic)
    for l in landfill:
        report.write(l + "\n")
    report.write("\n\n\n")

    # Print existent landfills in the report file
    report.write("============= ATERROS EXISTENTES ============= \n")
    existentlandfill = getExistentLandfill(citiesdic)
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
    logging.info("Gerando combinaçãoes...")
    combinations = list()
    i = 1
    while i <= MAX_SUB_ARRAYS and i <= MAX_CITIES:
        logging.info("Gerando combinaçãoes de tamanho: %d", i)
        combinations = combinations + list(sorted_k_partitions(clusters,i))
        i = i + 1
    logging.info("Quantidade de combinações: %d", len(combinations))
    report.write("Quantidade de combinações: " + repr(len(combinations)) + "\n")

    utvrs = getUTVR(citiesdic)
    if len(utvrs) != len(citiesdic):
        logging.info("Removendo combinaçãoes cujo sub-arranjo não possui uma UTVR...")
        combinations = removeArraysWithoutUTVR(combinations, utvrs)
        logging.info("Quantidade de combinações após a remoção: %d", len(combinations))
    report.write("Quantidade de combinações (desconsiderando arranjos com sub-arranjos sem UTVR): " + repr(len(combinations)) + "\n")

    if TRASH_THRESHOLD > 0.0:
        logging.info("Removendo combinaçãoes cujo sub-arranjo não possui a quantidade de lixo necessária...")
        combinations = removeArraysTrashThreshold(citiesdic, combinations, TRASH_THRESHOLD)
        logging.info("Quantidade de combinações após a remoção: %d", len(combinations))
    report.write("Quantidade de combinações (desconsiderando arranjos com sub-arranjos que não somam a quantidade de lixo produzida mínima): " + repr(len(combinations)) + "\n\n\n")

    logging.info("Cálculando valores (inbound, tecnologia e outbound) por combinação...")


    new_comb = list()
    #sub_arrays = set()
    for c in combinations:
        xcomb = list()
        for sub in c:
            subarray = list()
            for cluster in sub:   
                for city in cluster:
                    subarray.append(city)
            #sub_arrays.add(tuple(subarray))
            xcomb.append(subarray)
        new_comb.append(xcomb)
        
    #logging.info("Quantidade de subarranjos diferentes: %d", len(sub_arrays))

    data = []
    current = 0

    arrayCentralized = {}

    for i in new_comb:
        if current % 1000 == 0:
            logging.info("Progresso: %d de %d", current, len(new_comb))
        isCentralized = False
        if len(i) == 1:
            isCentralized = True

        arrayResult = {}
        subArrayResultList = list()

        # Array trash will be always the total, it's not necesary to calculate again. Same is valid for population
        arrayResult['trash'] = totalTrash
        arrayResult['population'] = totalPopulation
        # Initialize values for array
        arrayResult['inbound'] = 0
        arrayResult['inbound-sum'] = 0
        arrayResult['inbound-capex'] = 0
        arrayResult['inbound-show'] = 0
        arrayResult['inbound-show-sum'] = 0
        arrayResult['inbound-custo-incluindo-capex-nivel-arranjo'] = 0
        arrayResult['outbound'] = 0
        arrayResult['outbound-sum'] = 0
        arrayResult["outbound-custo-incluindo-capex-nivel-arranjo"] = 0
        arrayResult['outbound-show-capex'] = 0
        arrayResult['outbound-existent-landfill'] = 0
        arrayResult['outbound-existent-landfill-sum'] = 0
        arrayResult["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"] = 0
        arrayResult['outbound-show-capex-existent-landfill'] = 0
        arrayResult['technology'] = 0
        

        # Calculate inbound and outbound values without adding the capex
        for subArray in i:
            subArrayResult = {}
            subArrayResult = calculateInboundOutbound(citiesdic, distance, subArray, isCentralized, utvrs, landfill, existentlandfill, CAPEX_INBOUND, CAPEX_OUTBOUND, PAYMENT_PERIOD, MOVIMENTATION_COST, LANDFILL_DEVIATION, totalTrash, ADDITIONAL_COST)
            subArrayResult['trash'] = getSubArrayTrash(citiesdic, subArray)
            subArrayResult['rsu-range'] = getSubArrayRSURange(subArrayResult['trash'], rsutrash)
            subArrayResult['capex'] = getSubArrayCapex(subArrayResult['rsu-range'], subArrayResult['trash'], rsutrash)
            subArrayResult['opex'] = getSubArrayOpex(subArrayResult['rsu-range'], subArrayResult['trash'], rsutrash)
            subArrayResult['population'] = getSubArrayPopulation(citiesdic, subArray)
            subArrayResult['technology'] = (subArrayResult['capex']/PAYMENT_PERIOD + subArrayResult['opex'])
            subArrayResult['inbound-custo-incluindo-capex-nivel-arranjo'] = subArrayResult["inbound-show"] * subArrayResult["trash"] / totalTrash
            subArrayResult["inbound-custo-incluindo-capex-nivel-sub-arranjo"] = 0
            subArrayResult["total"] = subArrayResult['technology'] + subArrayResult["inbound"] + subArrayResult["outbound"]
            arrayResult['inbound-sum'] = arrayResult['inbound-sum'] + (subArrayResult["inbound"] * subArrayResult['trash'])
            arrayResult['inbound-show-sum'] = arrayResult['inbound-show-sum'] + (subArrayResult["inbound-show"] * subArrayResult['trash'])
            arrayResult['outbound-sum'] = arrayResult['outbound-sum'] + (subArrayResult["outbound"] * subArrayResult['trash'])
            arrayResult['outbound-existent-landfill-sum'] = arrayResult['outbound-existent-landfill-sum'] + (subArrayResult["outbound-existent-landfill"] * subArrayResult['trash'])
            subArrayResultList.append(subArrayResult)

        # Add capex value to inbound and outbound
        for subArray in subArrayResultList:
            if arrayResult['inbound-sum'] > 0:
                subArray['inbound-capex'] = subArray['inbound'] + capexInbound * ((subArray['inbound']*subArray['trash'])/arrayResult['inbound-sum'])
            else:
                subArray['inbound-capex'] = 0
            if arrayResult['inbound-show-sum'] > 0:
                t41 = (1000000/(subArray['trash']*312*PAYMENT_PERIOD))
                #t40 = (CAPEX_INBOUND * ((arrayResult['inbound-sum']-(subArray["inbound"] * subArray['trash']))/arrayResult['inbound-sum']))
                t40 = CAPEX_INBOUND * ((subArray['trash']*subArray['inbound'])/(arrayResult['inbound-sum']))
                if isCentralized:
                    t40 = CAPEX_INBOUND * ADDITIONAL_COST

                subArray["inbound-custo-incluindo-capex-nivel-sub-arranjo"] = subArray['inbound'] + (t40*t41)
                subArray["inbound-custo-incluindo-capex-nivel-arranjo"] = subArray['inbound-custo-incluindo-capex-nivel-sub-arranjo'] * subArray['trash'] / totalTrash
            else:
                subArray['inbound-custo-incluindo-capex-nivel-arranjo'] = 0    
            if arrayResult['outbound-sum'] > 0:

                t41 = (1000000/(subArray['trash']*312*PAYMENT_PERIOD))
                #t40 = (CAPEX_INBOUND * ((arrayResult['inbound-sum']-(subArray["inbound"] * subArray['trash']))/arrayResult['inbound-sum']))
                t40 = CAPEX_OUTBOUND * ((subArray['trash']*subArray['outbound'])/(arrayResult['outbound-sum']))

                if isCentralized:
                    subArray["outbound-custo-incluindo-capex-nivel-sub-arranjo"] = subArray["outbound"] + (t40*t41)
                else:
                    subArray["outbound-custo-incluindo-capex-nivel-sub-arranjo"] = subArray["outbound"] + (ADDITIONAL_COST*t40*t41)
                subArray["outbound-custo-incluindo-capex-nivel-arranjo"] = subArray["outbound-custo-incluindo-capex-nivel-sub-arranjo"] * subArray['trash'] / totalTrash
                #subArray["outbound-custo-incluindo-capex-nivel-arranjo"] = subArray['outbound'] + capexOutbound * (subArray['trash']/totalTrash)
                #subArray["outbound-custo-incluindo-capex-nivel-sub-arranjo"] = subArray["outbound-custo-incluindo-capex-nivel-arranjo"] * totalTrash / subArray['trash']

                t40 = CAPEX_OUTBOUND * ((subArray['trash']*subArray['outbound-existent-landfill'])/(arrayResult['outbound-existent-landfill-sum']))
                if isCentralized:
                    subArray["outbound-existent-landfill-custo-incluindo-capex-nivel-sub-arranjo"] = subArray["outbound-existent-landfill"] + (t40*t41)
                else:
                    subArray["outbound-existent-landfill-custo-incluindo-capex-nivel-sub-arranjo"] = subArray["outbound-existent-landfill"] + (ADDITIONAL_COST*t40*t41)
                subArray["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"] = subArray["outbound-existent-landfill-custo-incluindo-capex-nivel-sub-arranjo"] * subArray['trash'] / totalTrash


                #subArray["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"] = subArray['outbound-existent-landfill'] + capexOutbound * (subArray['trash']/totalTrash)
                #subArray["outbound-existent-landfill-custo-incluindo-capex-nivel-sub-arranjo"] = subArray["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"] * totalTrash / subArray['trash']
            else:
                subArray["outbound-custo-incluindo-capex-nivel-arranjo"] = 0
                subArray["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"] = 0

            subArray["total-capex"] = subArray['technology'] + subArray["inbound-custo-incluindo-capex-nivel-sub-arranjo"] + subArray["outbound-custo-incluindo-capex-nivel-sub-arranjo"]
            arrayResult['inbound'] = arrayResult['inbound'] + (subArray['inbound'] * subArray['trash'])
            arrayResult['inbound-show'] = arrayResult['inbound-show'] + subArray['inbound-show']
            arrayResult['inbound-custo-incluindo-capex-nivel-arranjo'] = arrayResult['inbound-custo-incluindo-capex-nivel-arranjo'] + subArray['inbound-custo-incluindo-capex-nivel-arranjo']
            arrayResult['outbound'] = arrayResult['outbound'] + (subArray['outbound'] * subArray['trash'])
            arrayResult['outbound-show-capex'] = arrayResult['outbound-show-capex'] + subArray["outbound-custo-incluindo-capex-nivel-arranjo"]
            arrayResult['outbound-show-capex-existent-landfill'] = arrayResult['outbound-show-capex-existent-landfill'] + subArray["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"]
            arrayResult['outbound-existent-landfill'] = arrayResult['outbound-existent-landfill'] + (subArray['outbound-existent-landfill'] * subArray['trash'])
            arrayResult['inbound-capex'] = arrayResult['inbound-capex'] + (subArray['inbound-capex'] * subArray['trash'])
            arrayResult["outbound-custo-incluindo-capex-nivel-arranjo"] = arrayResult["outbound-custo-incluindo-capex-nivel-arranjo"] + subArray["outbound-custo-incluindo-capex-nivel-arranjo"]
            arrayResult["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"] = arrayResult["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"] + subArray["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"]
            arrayResult['technology'] = arrayResult['technology'] + (subArray['technology'] * subArray['trash'])

        arrayResult["arranjo"] = i
        arrayResult["sub"] = subArrayResultList
        arrayResult["inbound"] = arrayResult['inbound']/arrayResult['trash']
        arrayResult["outbound"] = arrayResult['outbound']/arrayResult['trash']
        arrayResult["inbound-capex"] = arrayResult['inbound-capex']/arrayResult['trash']
        #arrayResult["outbound-custo-incluindo-capex-nivel-arranjo"] = arrayResult["outbound-custo-incluindo-capex-nivel-arranjo"]/arrayResult['trash']
        arrayResult["technology"] = arrayResult['technology']/arrayResult['trash']
        arrayResult['outbound-existent-landfill'] = arrayResult['outbound-existent-landfill']/arrayResult['trash']
        #arrayResult["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"] = arrayResult["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"]/arrayResult['trash']
        arrayResult["total"] = arrayResult["technology"] + arrayResult["inbound"] + arrayResult["outbound"]
        arrayResult["total-capex"] = arrayResult["technology"] + arrayResult["inbound-custo-incluindo-capex-nivel-arranjo"] + arrayResult["outbound-custo-incluindo-capex-nivel-arranjo"]
        
        if isCentralized:
            arrayCentralized = arrayResult

        # Verifica se a quantidade de resultados atual é maior que a quantidade desejada a ser analizada
        if len(data) >= MAX_ARRAYS:
            data = sorted(data, key = lambda k: (k["total-capex"]))
            last = data[-1]
            # Somente adiciona no array de data se um resultado melhor for encontrado
            if arrayResult["total-capex"] < last["total-capex"]:
                #print("Removendo ", last["total-capex"], " e inserindo ", arrayResult["total-capex"])
                data.pop()
                data.append(arrayResult)
        else:
            data.append(arrayResult)

        current = current + 1

    logging.info("Progresso: 100%")
    logging.info("Ordenando combinações...")
    data = sorted(data, key = lambda k: (k["total-capex"]))

    logging.info("Escrevendo relatórios...")
    report.write("\n\n============= ARRANJO CENTRALIZADO ============= \n")

    report.write(repr(arrayCentralized["arranjo"]) + "\n")
    report.write("- Lixo: " + repr(arrayCentralized["trash"]) + "\n")
    #report.write("- Custo Total: " + repr(arrayCentralized["total"]) + "\n")
    report.write("- Custo Total - Capex: " + repr(arrayCentralized["total-capex"]) + "\n")
    #report.write("-- Inbound (Old): " + repr(arrayCentralized["inbound"]) + "\n")
    #report.write("-- Inbound + Capex (Old): " + repr(arrayCentralized["inbound-capex"]) + "\n")
    #report.write("-- Inbound (custo opex): " + repr(arrayCentralized["inbound-show"]) + "\n")
    report.write("-- Inbound - Custo incluindo CAPEX (Nível Arranjo): " + repr(arrayCentralized["inbound-custo-incluindo-capex-nivel-arranjo"]) + "\n")
    report.write("-- Outbound - Custo incluindo CAPEX (Nível Arranjo): " + repr(arrayCentralized["outbound-custo-incluindo-capex-nivel-arranjo"]) + "\n")
    report.write("-- Outbound Aterro Existente - Custo incluindo CAPEX (Nível Arranjo): " + repr(arrayCentralized["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"]) + "\n")
    report.write("-- Tecnologia: " + repr(arrayCentralized['technology']) + "\n\n")
    #report.write("-- Outbound: " + repr(arrayCentralized["outbound"]) + "\n")
    #report.write("-- Outbound - Capex: " + repr(arrayCentralized["outbound-custo-incluindo-capex-nivel-arranjo"]) + "\n")
    #report.write("-- Outbound Aterro Existente: " + repr(arrayCentralized["outbound-existent-landfill"]) + "\n")
            
    report.write("-- Sub-arranjos:\n")

    output.write(repr(arrayCentralized["arranjo"]) + ";Sumário;NA;NA;NA;" + repr(arrayCentralized["population"]) + ";" + repr(arrayCentralized["total-capex"]) + ";" + repr(arrayCentralized["trash"]) + ";" + repr(arrayCentralized['technology']) + ";" + repr(arrayCentralized["inbound-custo-incluindo-capex-nivel-arranjo"]) + ";" + repr(arrayCentralized["outbound-custo-incluindo-capex-nivel-arranjo"]) + ";" + repr(arrayCentralized["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"]) + "\n")

    for x in range(len(arrayCentralized["sub"])):
        output.write(repr(arrayCentralized["arranjo"]) + ";" + repr(arrayCentralized["sub"][x]["sub-arranjo"]) + ";" + repr(arrayCentralized["sub"][x]["aterro"]) + ";" + repr(arrayCentralized["sub"][x]["aterro-existente"]) + ";" + repr(arrayCentralized["sub"][x]["utvr"]) + ";" + repr(arrayCentralized["sub"][x]['population']) + ";" + repr(arrayCentralized["sub"][x]["total-capex"]) + ";" + repr(arrayCentralized["sub"][x]["trash"]) + ";" + repr(arrayCentralized["sub"][x]['technology']) + ";" + repr(arrayCentralized["sub"][x]["inbound-custo-incluindo-capex-nivel-sub-arranjo"])  + ";" + repr(arrayCentralized["sub"][x]["outbound-custo-incluindo-capex-nivel-sub-arranjo"])  + ";" + repr(arrayCentralized["sub"][x]["outbound-existent-landfill-custo-incluindo-capex-nivel-sub-arranjo"]) + "\n")

        report.write("\t" + repr(arrayCentralized["sub"][x]["sub-arranjo"]) + "\n")
        report.write("\t-- UTVR: " + repr(arrayCentralized["sub"][x]["utvr"]) + "\n")
        report.write("\t-- Aterro: " + repr(arrayCentralized["sub"][x]["aterro"]) + "\n")
        report.write("\t-- Lixo: " + repr(arrayCentralized["sub"][x]['trash']) + "\n")
        report.write("\t-- População: " + repr(arrayCentralized["sub"][x]['population']) + "\n")
        #report.write("\t-- Total: " + repr(arrayCentralized["sub"][x]["total"]) + "\n")
        report.write("\t-- Total - Capex: " + repr(arrayCentralized["sub"][x]["total"]) + "\n")
        #report.write("\t-- Inbound: " + repr(arrayCentralized["sub"][x]["inbound"]) + "\n")
        #report.write("\t-- Inbound - Capex: " + repr(arrayCentralized["sub"][x]["inbound-capex"]) + "\n")
        report.write("\t-- Inbound - Custo OPEX: " + repr(arrayCentralized["sub"][x]["inbound"]) + "\n")
        report.write("\t-- Inbound - Custo incluindo CAPEX: " + repr(arrayCentralized["sub"][x]["inbound-custo-incluindo-capex-nivel-arranjo"]) + "\n")
        report.write("\t-- Inbound - Custo incluindo CAPEX (Nível Subarranjo): " + repr(arrayCentralized["sub"][x]["inbound-custo-incluindo-capex-nivel-sub-arranjo"]) + "\n")
        report.write("\t-- Outbound - Custo OPEX: " + repr(arrayCentralized["sub"][x]["outbound"]) + "\n")
        report.write("\t-- Outbound - Custo incluindo CAPEX: " + repr(arrayCentralized["sub"][x]["outbound-custo-incluindo-capex-nivel-arranjo"]) + "\n")
        report.write("\t-- Outbound - Custo incluindo CAPEX (Nível Subarranjo): " + repr(arrayCentralized["sub"][x]["outbound-custo-incluindo-capex-nivel-sub-arranjo"]) + "\n")
        report.write("\t-- Outbound Aterro Existente - Custo OPEX: " + repr(arrayCentralized["sub"][x]["outbound-existent-landfill"]) + "\n")
        report.write("\t-- Outbound Aterro Existente - Custo incluindo CAPEX: " + repr(arrayCentralized["sub"][x]["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"]) + "\n")
        report.write("\t-- Outbound Aterro Existente - Custo incluindo CAPEX (Nível Subarranjo): " + repr(arrayCentralized["sub"][x]["outbound-existent-landfill-custo-incluindo-capex-nivel-sub-arranjo"]) + "\n")
        report.write("\t-- Tecnologia: " + repr(arrayCentralized["sub"][x]['technology']) + "\n")
        report.write("\t\t-- Capex: " + repr(arrayCentralized["sub"][x]["capex"]) + "\n")
        report.write("\t\t-- Opex: " + repr(arrayCentralized["sub"][x]["opex"]) + "\n\n")

    report.write("\n\n============= TOP " + repr(MAX_ARRAYS) + " ARRANJOS ============= \n")
    for i in range(len(data)):
        if i >= MAX_ARRAYS:
            break

        output.write(repr(data[i]["arranjo"]) + ";Sumário;NA;NA;NA;" + repr(data[i]["population"]) + ";" + repr(data[i]["total-capex"]) + ";" + repr(data[i]["trash"]) + ";" + repr(data[i]['technology']) + ";" + repr(data[i]["inbound-custo-incluindo-capex-nivel-arranjo"])  + ";" + repr(data[i]["outbound-custo-incluindo-capex-nivel-arranjo"]) + ";" + repr(data[i]["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"]) + "\n")

        report.write(repr(i+1) + ".\t" + repr(data[i]["arranjo"]) + "\n")
        report.write("- Lixo: " + repr(data[i]["trash"]) + "\n")
        #report.write("- Custo Total: " + repr(data[i]["total"]) + "\n")
        report.write("- Custo Total + Capex: " + repr(data[i]["total-capex"]) + "\n")
        #report.write("-- Inbound (Old): " + repr(data[i]["inbound"]) + "\n")
        #report.write("-- Inbound + Capex (Old): " + repr(data[i]["inbound-capex"]) + "\n")
        #report.write("-- Inbound: " + repr(data[i]["inbound-show"]) + "\n")
        report.write("-- Inbound - Custo incluindo CAPEX (Nível Arranjo): " + repr(data[i]["inbound-custo-incluindo-capex-nivel-arranjo"]) + "\n")
        report.write("-- Outbound - Custo incluindo CAPEX (Nível Arranjo): " + repr(data[i]["outbound-custo-incluindo-capex-nivel-arranjo"]) + "\n")
        report.write("-- Outbound Aterro Existente - Custo incluindo CAPEX (Nível Arranjo): " + repr(data[i]["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"]) + "\n")
        report.write("-- Tecnologia: " + repr(data[i]['technology']) + "\n\n")
        #report.write("-- Outbound: " + repr(data[i]["outbound"]) + "\n")
        #report.write("-- Outbound + Capex: " + repr(data[i]["outbound-custo-incluindo-capex-nivel-arranjo"]) + "\n")
        #report.write("-- Outbound Aterro Existente: " + repr(data[i]["outbound-existent-landfill"]) + "\n")
        

        report.write("-- Sub-arranjos:\n")
        for x in range(len(data[i]["sub"])):
            output.write(repr(data[i]["arranjo"]) + ";" + repr(data[i]["sub"][x]["sub-arranjo"]) + ";" + repr(data[i]["sub"][x]["aterro"]) + ";" + repr(data[i]["sub"][x]["aterro-existente"]) + ";" + repr(data[i]["sub"][x]["utvr"]) + ";" + repr(data[i]["sub"][x]['population']) + ";" + repr(data[i]["sub"][x]["total-capex"]) + ";" + repr(data[i]["sub"][x]['trash']) + ";" + repr(data[i]["sub"][x]['technology']) + ";" + repr(data[i]["sub"][x]["inbound-custo-incluindo-capex-nivel-sub-arranjo"])  + ";" + repr(data[i]["sub"][x]["outbound-custo-incluindo-capex-nivel-sub-arranjo"]) + ";" + repr(data[i]["sub"][x]["outbound-existent-landfill-custo-incluindo-capex-nivel-sub-arranjo"]) + "\n")

            report.write("\t" + repr(data[i]["sub"][x]["sub-arranjo"]) + "\n")
            report.write("\t-- UTVR: " + repr(data[i]["sub"][x]["utvr"]) + "\n")
            report.write("\t-- Aterro: " + repr(data[i]["sub"][x]["aterro"]) + "\n")
            report.write("\t-- Lixo: " + repr(data[i]["sub"][x]['trash']) + "\n")
            report.write("\t-- População: " + repr(data[i]["sub"][x]['population']) + "\n")
            #report.write("\t-- Total: " + repr(data[i]["sub"][x]["total"]) + "\n")
            report.write("\t-- Total + Capex: " + repr(data[i]["sub"][x]["total-capex"]) + "\n")
            #report.write("\t-- Inbound: " + repr(data[i]["sub"][x]["inbound"]) + "\n")
            #report.write("\t-- Inbound + Capex: " + repr(data[i]["sub"][x]["inbound-capex"]) + "\n")
            report.write("\t-- Inbound - SUM: " + repr(data[i]["sub"][x]["inbound-sum"]) + "\n")
            report.write("\t-- Inbound - Custo OPEX: " + repr(data[i]["sub"][x]["inbound"]) + "\n")
            report.write("\t-- Inbound - Custo incluindo CAPEX: " + repr(data[i]["sub"][x]["inbound-custo-incluindo-capex-nivel-arranjo"]) + "\n")
            report.write("\t-- Inbound - Custo incluindo CAPEX (Nível Subarranjo): " + repr(data[i]["sub"][x]["inbound-custo-incluindo-capex-nivel-sub-arranjo"]) + "\n")
            
            
            report.write("\t-- Outbound - Custo OPEX: " + repr(data[i]["sub"][x]["outbound"]) + "\n")
            report.write("\t-- Outbound - Custo incluindo CAPEX: " + repr(data[i]["sub"][x]["outbound-custo-incluindo-capex-nivel-arranjo"]) + "\n")
            report.write("\t-- Outbound - Custo incluindo CAPEX (Nível Subarranjo): " + repr(data[i]["sub"][x]["outbound-custo-incluindo-capex-nivel-sub-arranjo"]) + "\n")
           
            report.write("\t-- Outbound Aterro Existente - Custo OPEX: " + repr(data[i]["sub"][x]["outbound-existent-landfill"]) + "\n")
            report.write("\t-- Outbound Aterro Existente - Custo incluindo CAPEX: " + repr(data[i]["sub"][x]["outbound-existent-landfill-custo-incluindo-capex-nivel-arranjo"]) + "\n")
            report.write("\t-- Outbound Aterro Existente - Custo incluindo CAPEX (Nível Subarranjo): " + repr(data[i]["sub"][x]["outbound-existent-landfill-custo-incluindo-capex-nivel-sub-arranjo"]) + "\n")

            report.write("\t-- Tecnologia: " + repr(data[i]["sub"][x]['technology']) + "\n")
            report.write("\t\t-- Capex: " + repr(data[i]["sub"][x]["capex"]) + "\n")
            report.write("\t\t-- Opex: " + repr(data[i]["sub"][x]["opex"]) + "\n\n")


        report.write("-----------------------------------------------------------------\n\n")

    # Close report and output file
    report.close()
    output.close

if __name__ == "__main__":
    main()