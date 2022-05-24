import sys
import logging
import numpy as np
import csv
from more_itertools import set_partitions

logging.basicConfig(stream=sys.stderr, level=logging.DEBUG)

try:
    arg = sys.argv[1]
except IndexError:
    raise SystemExit(f"Usage: {sys.argv[0]} <cities.csv> <distance.csv>")

# The maximium number of allowed cities to generate the combinations
# due to performance issues.
MAX_CITIES = 14


# Read cities from CSV file

# Read distances from CSV file