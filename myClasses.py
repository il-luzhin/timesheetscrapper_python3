'''

'''

# imported from the standard library
import re
import xlrd

# imported from third party repos

# imported from local directories
import config as cfg

class timesheet:

    def __init__(self, name_row, name_column, start_data, end_data, col_data):
        self.name_row = name_row
        self.name_column = name_column
        self.start_data = start_data
        self.end_data = end_data
        self.col_data = col_data

    def scrape(self, df, scrape_sheet):

       mylist = []

        r_row = self.start_data  # r_row is now the read_book row
        r_col = self.col_data

        # A loop to iterate through the time slots one at a time
        for r_row in range(self.start_data, self.end_data):
            # Find the first slot with data
            if scrape_sheet.cell_type(r_row, 2) != 0:
                print("writing data")

                # write the HeadAlphaID
                data = 'z'
                mylist.append(data)

                # write persons employee number
                data = scrape_sheet.cell_value(self.name_row, self.name_column)
                my_dict = searchDict(cfg.dict_heads)
                for head_num in my_dict.search_for_match(data):
                    mylist.append(head_num)

                # write date loop
                i = self.start_data
                while i < 70:
                    if ((r_row >= i) and (r_row <= i + 6)):
                        data = scrape_sheet.cell_value(r_row, 0)
                        shift_date_tuple = xlrd.xldate_as_tuple(data, 1)
                        day = f"{shift_date_tuple[2]}"
                        month = f"{shift_date_tuple[1]}"
                        year = f"{shift_date_tuple[0]}"
                        shift_date = day + '/' + month + '/' + year
                        mylist.append(shift_date)
                        print(shift_date)
                        i += 7
                    else:
                        i += 7

                    # write time in
                    if head_num == 3:  # if it's kris, then...
                        data = scrape_sheet.cell_value(r_row, r_col)
                        kris_str = str(data)
                        count_int = len(kris_str)

                        if count_int == 1:
                            kris_tuple = (0, data, 0, 0)
                        elif count_int == 2:
                            kris_tuple = (kris_str[0], kris_str[1], 0, 0)
                        elif count_int == 3:
                            kris_tuple = (0, kris_str[0], kris_str[1], kris_str[2])
                        else:
                            kris_tuple = (kris_str[0], kris_str[1], kris_str[2], kris_str[3])

                        time = str(kris_tuple[0]) + str(kris_tuple[1]) + ":" + str(kris_tuple[2]) + str(kris_tuple[3])
                        print(time)
                        klnew_sheet.write(klw_row, klr_col + 2, time)
                    else:  # it's not kris, so....
                        write_time(read_sheet, r_row, 2, new_sheet, w_row)


class salariedHead:

    headCount = 0

    # the instance, 'self' in this case, is always teh first argument
    # in it, you define the class variables, then assign thjose variables to the instance variables
    def __init__(self, headAlphaID, headID, lastName, firstName, dir_path, has_left):
        self.headAlphaID = headAlphaID
        self.HeadID = headID
        self.lastName = lastName
        self. firstName = firstName
        self.dir_path = dir_path
        self.has_left = has_left
        salariedHead.headCount += 1 #this increments every time an employee is created

    #repr allows us to print the list with 'salariedHeads(my_var)'
    def __repr__(self):
        return f"salariedHead({self.HeadID}, '{self.lastName}', '{self.firstName}', '{self.dir_path}', {self.has_left})"

    #concat name
    def fullName(self):
        return f"{self.firstName} {self.lastName}"

    def email(self, domain):
        return self.firstName[0].lower() + self.lastName.lower() + domain

#test data
def testheads():
    Wi = salariedHead('z', 99, 'Wonka', 'Willy', 'N/A', False)
    print(Wi)
    print(Wi.fullName())
    print(Wi.email("@chocolatefactory.com"))
    print(salariedHead.headCount)

class searchDict(dict):

    def search_for_match(self, event):
        return (self[key] for key in self if re.match(key, event))




class timesheetloop(self):
    pass



#test salaried head class
if __name__ == '__main__':
    print()
    print('------------')
    print("METHOD CHECK")
    print('------------')
    print()
    testheads()
