from constants import *


def find_column(ip, xlsx):
    row = "4"
    for c in range(31):
        cellID = COLUMN[c] + row
        if xlsx[cellID].value == ip:
            return COLUMN[c]


def check_status(status):
    if status == 'T':
        return 1
    if status == 'F':
        return 2


