# Scrape LDraw library and add parts into xlsx file
# blackbunt c() 2020-08-25
import os.path
import re
import urllib.request
import xlsxwriter


# pattern = "([^\\r]*\s)(.+?(?=\\r))"

def build_string(partnumber):
    ldraw_url = 'https://www.ldraw.org/library/official/parts/'
    return ldraw_url + str(partnumber) + ".dat"


def check_part_online(url, no):
    try:
        weburl = urllib.request.urlopen(str(url))
        print('Part no. ' + str(no) + ' exists.')
        return True
    # error handling
    except urllib.error.HTTPError as error:
        print(error)
        return False
    except urllib.error.URLError as error:
        print(error)
        return False
    print(urllib.request.get_method(url))


def get_partinfo(url, no):
    weburl = urllib.request.urlopen(url)
    data = weburl.read()
    data = str(data)
    # regex pattern for checking if file moved to new url
    corr_file_pattern = r"^(b\'0\s\~)([^\\r]*\s)(.+?(?=\\r))"

    if re.search(corr_file_pattern, data):
        corr_file = re.search(corr_file_pattern, data)
        # check if file moved
        if corr_file[2] is not None:
            if 'Moved to ' == str(corr_file[2]):
                new_part_no = corr_file[3]
                print("Part moved to " + new_part_no + ".")
                new_part = ["new", new_part_no]
                return new_part
            else:
                print("error")
    # if file not moved extract part name
    # regex pattern to match name of part
    else:
        pattern = r"([^b'0\s\\r]*\s)(.+?(?=\\r))"
        part_name = re.search(pattern, data)[0]

        # if file not moved part no is still the old one
        # return a list with the part name and the part id
        info = [no, part_name]
        return info


def name_workbook():
    # creates a excel file
    # checks for existing excel file
    # returns workbook
    while True:
        filename = input("Please enter Filename for Excel file: ")
        if not filename == 'exit':
            filename_str = str(filename) + ".xlsx"
            # check for typos
            input_name = input(filename_str + ". Is this correct? y/n ")
            if input_name == "y":
                # if file not exists, then create it
                if not os.path.isfile(filename_str):
                    workbook = xlsxwriter.Workbook(filename_str)
                    return workbook
                else:
                    print("File does already exist!")
                    # Overwriting prompt
                    while True:
                        input_file = input("Do you want to overwrite it? y/n ")
                        if input_file == "y":
                            workbook = xlsxwriter.Workbook(filename_str)
                            return workbook
                        elif input_file == "n":
                            break
                        else:
                            print("Wrong input!")
            elif input_name == "n":
                return None
            else:
                print("Wrong input!")

        else:
            # if input exit, then stop program
            print("Exiting...")
            exit()


# Greetz
print("Welcome to LDRAW Part Scraper!")
print("to exit, type 'exit'")
# create excel file
while True:
    excel_file = name_workbook()
    if excel_file is not None:
        break
# name worksheet
excel_sheet = excel_file.add_worksheet("Part List")
# write header
excel_sheet.write(0, 0, "part_id")
excel_sheet.write(0, 1, "name")
# set position in worksheet
row = 1
col = 0
# start part scraping
while True:
    part_no = input('Enter Part Number: ')
    if not part_no == 'exit':
        while True:
            part_info = []
            # generate part url string
            part_url = build_string(part_no)
            # if online, do magic
            if check_part_online(part_url, part_no):
                # get part infos
                part_info = get_partinfo(part_url, part_no)
                # if file not moved, write infos into worksheet
                if part_info[0] != "new":
                    print("Part no. " + str(part_info[0]) + " is a " + str(part_info[1]) + ".")
                    excel_sheet.write(row, col, part_info[0])
                    excel_sheet.write(row, col + 1, part_info[1])
                    row += 1
                    # return to input loop
                    break
                # if moved file detected, create a new part request
                elif part_info[0] == "new":
                    # generate new part url string
                    part_url = build_string(part_info[1])
                    # if online, do magic
                    if check_part_online(part_url, part_info[1]):
                        part_info = get_partinfo(part_url, part_info[1])
                        if part_info[0] != "new":
                            print("Part no. " + str(part_info[0]) + " is a " + str(part_info[1]) + ".")
                            excel_sheet.write(row, col, part_info[0])
                            excel_sheet.write(row, col + 1, part_info[1])
                            row += 1
                            # return to input loop
                            break
                            # if file offline return to input loop
                    else:
                        break

                else:
                    print("Database seems empty for this part...")
            # if file offline return to input loop
            else:
                break

    else:
        # save changes and exit
        excel_file.close()
        print("Exiting...")
        exit()
