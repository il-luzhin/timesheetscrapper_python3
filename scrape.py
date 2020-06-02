'''
This file is for taking collected timesheets, scrapping them, and
placing the scraped data into a .xls for further processing.
'''

# imported from standard library
import os

# imported from third party repos
import xlrd


# imported from local directories
import config as cfg
import kris_fix as kf # added when kris broke the scrapper

# this method will write the time to the new .xls sheet
def write_time(timeread_sheet, timer_row,timer_col, timenew_sheet, timew_row):
    data = timeread_sheet.cell_value(timer_row, timer_col)
    if data == '':
        shift_in_tuple = (0, 0, 0, 0, 0, 0)
    else:
        shift_in_tuple = xlrd.xldate_as_tuple(data, 1)
    if shift_in_tuple[3] < 10:
        half1_time = f"{shift_in_tuple[3]}"
    else:
        half1_time = f"{shift_in_tuple[3]}"
    if shift_in_tuple[4] == 0:
        half2_time = f"{shift_in_tuple[4]}0"
    else:
        half2_time = f"{shift_in_tuple[4]}"
    time = half1_time + ":" + half2_time
    print(time)
    timenew_sheet.write(timew_row, timer_col + 2, time)

# this method will write the hours worked to the new .xls sheet
def write_hrs(hrsread_sheet, hrsr_row, hrsr_col, hrsnew_sheet, hrsw_row, hrsw_col):
    data = hrsread_sheet.cell_value(hrsr_row, hrsr_col)
    hrsnew_sheet.write(hrsw_row, hrsw_col, data)

def main():


                    # write time in
                    if head_num == 3:  # if it's kris, then...
                        kf.kf_format(w_row,r_row,new_sheet,read_sheet,2)
                    else:  # it's not kris, so....
                        write_time(read_sheet, r_row, 2, new_sheet, w_row)

                    # write time out - kris_fix2
                    if head_num == 3:
                        kf.kf_format(w_row, r_row, new_sheet, read_sheet, 3)
                    else:
                        write_time(read_sheet, r_row, 3, new_sheet, w_row)

                    # write reg time, ot, dt
                    w_col = 8
                    write_hrs(read_sheet, r_row, 4, new_sheet, w_row, w_col)
                    w_col = w_col + 1

                    write_hrs(read_sheet, r_row, 5, new_sheet, w_row, w_col)
                    w_col = w_col + 1

                    write_hrs(read_sheet, r_row, 6, new_sheet, w_row, w_col)

                    # write accounting code
                    data = read_sheet.cell_value(r_row, 8)
                    print(data)
                    acct_num = cfg.acct_codes[data]
                    new_sheet.write(w_row, 11, acct_num)

                    # write show data
                    data = 'J'  # this is year specific CHANGE THIS FOR YOUR NEEDS - WRITE SOMETHING BETTER
                    new_sheet.write(w_row, 6, data)

                    data = '0-310820'  # this is year specific
                    if acct_num != '6210-50-504' and acct_num != '6200-50-504':
                        new_sheet.write(w_row, 7, data)
                    # elif acct_num != '6200-50-504':
                    #   new_sheet.write(w_row, 7, data)
                    else:
                        new_sheet.write(w_row, 7, "")

                    # showcall true/false
                    data = read_sheet.cell_value(r_row, 9)
                    if data != '':
                        new_sheet.write(w_row, 12, 1)
                    else:
                        new_sheet.write(w_row, 12, 0)

                    # Meal Penalty true/false
                    data = read_sheet.cell_value(r_row, 7)
                    if data == 1:
                        new_sheet.write(w_row, 13, 1)
                    else:
                        new_sheet.write(w_row, 13, 0)

                    w_row = w_row + 1  # move along in the write_book

                else:
                    print("no data in cel E" + str((r_row) + 1))  # move on to the next time slot

    else: print("this is NOT a timesheet")
    print()
    print("All done! Time to save to...")
    print(f"{cfg.home}\Desktop\scraped_by_python.xls")
    new_book.save(f"{cfg.home}\Desktop\scraped_by_python.xls")
    # putting a hold on these two lines while I work on the post2db.py
    # print("Opening saved file")
    # os.system(f"start EXCEL.EXE {cfg.home}\Desktop\scraped_by_python.xls")
    print()

if __name__ == '__main__':
    print()
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print("         timesheetscrapper_python3.py file launched")
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print()
    main()
    print()
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print("                        VICTORY!")
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print()