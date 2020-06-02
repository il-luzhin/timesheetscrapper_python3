'''
This is the main file for several pieces of the timesheet scrapper
Please see each file for more info
'''

# imported from standard library
import pandas as pd


# imported from third party repos

# imported from local directories
import config as cfg
import myClasses as mycls

def main():

    # create the df
    cols_to_use = ["Shift", "HeadIDLetter", "HeadIDNumber",
                    "Date", "InTime", "OutTime",
                    "EventYrID", "EventID", "Reg",
                    "OT", "Double", "Acct",
                    "Blackscall", "MP"]

    df_scraped = pd.DataFrames(columns=cols_to_use)

    # create a list of the read_books
    read_list = os.listdir(cfg.dir)
    print(f"We need to read approximately {len(read_list)} files")
    print("Are you ready to read? RETURN for yes, CTRL+C for no.")
    input()

    # a loop to iterate through the read_list
    i = 0  # file list position number
    for i in range(len(read_list)):
        read_file = (cfg.dir + '\\' + read_list[i])
        print(read_file)
        read_book = xlrd.open_workbook(read_file)
        read_sheet = read_book.sheet_by_index(0)

    # is this actually a timesheet? And which one is it?
    if read_sheet.cell_value(7, 0) == 'SUNDAY':
        print("This timesheet was designed in 2011. Begin data scrape")
        # scrape(7) #!!!!FINSH THIS LINE!!!
    elif read_sheet.cell_value(19, 0) == 'SUNDAY':
        print("This timesheet was designed in 2015. Begin data scrape")
        tf15 = mycls.timesheet(15,2,19,68)
    elif read_sheet.cell_value(14, 0) == 'SUNDAY':
        print("This timesheet is a casual's timesheet. Begin data scrape")
        # scrape(14) #!!!!FINSH THIS LINE!!!

if __name__ == '__main__':
    print()
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print("               AC - TIMESHEET SCRAPPER APP")
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print()
    main()
    print()
    print()
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print("                        VICTORY!")
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print()