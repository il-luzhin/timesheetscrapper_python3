'''

'''

# standard library
import pyodbc
import pandas as pd

# third part libraries
import config as cfg

# local repo

def find_crew_number(crew_name):
    mylist = (str.split(crew_name))
    query = ("""
        SELECT * FROM CrewNamesTable
        WHERE First Name = ?
        AND LastName = ?
    """, mylist[0],mylist[1])
    df_crewNum = pd.read_sql(query, cfg.conn)
    crew_num = df_crewNum[0]
    return crew_num

# find the number we need to start the new data with
def find_next_row_from_db(my_table, my_column):
    query = ("SELECT * FROM ?", my_table)
    df_hShift = pd.read_sql(query, cfg.conn)
    last_shift = df_hShift[my_column].max()
    # print(last_shift) # for testing
    new_shift = last_shift + 1
    return new_shift

def grabeventYR2(datestr):
    newdate = str.split(datestr, '/')
    yr_int = int(newdate[2])
    mos_int = int(newdate[1])
    yr_strt = 2011
    mos_strt = 9
    yr_diff = yr_int - yr_strt
    mos_diff = mos_int - mos_strt

    if yr_int == 2019:
        if mos_int < mos_strt:
            if mos_diff < 0:
                alphayr = ord('A') + yr_diff - 1
            else:
                alphayr = ord('A') + yr_diff
        else:
            if mos_diff < 0:
                alphayr = ord('A') + yr_diff
            else:
                alphayr = ord('A') + yr_diff +1
    elif yr_int > 2019:
        if mos_diff < 0:
            alphayr = ord('A') + yr_diff
        else:
            alphayr = ord('A') + yr_diff + 1
    else:
        if mos_diff < 0:
            alphayr = ord('A')+yr_diff-1
        else:
            alphayr = ord('A')+yr_diff

    return chr(alphayr)


# read from SQL and print to screen
def read2screen(query, conn):
    print("Read Method")
    cursor = conn.cursor()
    cursor.execute(query)
    for row in cursor:
        print(row)


# read from SQL without printing
def read2(query, conn):
    print("Read Method")
    cursor = conn.cursor()
    cursor.execute(query)

# general method
def general(query, conn):
    cursor = conn.cursor()
    cursor.execute(query)

# does your table exist?
def checkTableExists(dbcon, tablename):
    dbcur = dbcon.cursor()
    dbcur.execute("""
        SELECT COUNT(*)
        FROM information_schema.tables
        WHERE table_name = '{0}'
        """.format(tablename.replace('\'', '\'\'')))
    if dbcur.fetchone()[0] == 1:
        dbcur.close()
        return True

    dbcur.close()
    return False

def drop_table(self, table):
    self._exec(schema.DropTable(table))

if __name__ == '__main__':
    print()
    print('------------')
    print("METHOD CHECK")
    print('------------')
    print()
    query = "SELECT * FROM sys.tables"
    read(query, cfg.conn)


