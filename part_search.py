"""
Author: Austin Arrell
8/17/2020
V0.01

The purpose of this program is to query websites in order to determine
model numbers that a list of part numbers can fit into.
"""
import requests
import pandas as pd
from bs4 import BeautifulSoup as bs


part_list = []


def retrieve_part_numbers():
    """
    Grabs the list of part numbers from a file. For now it is a hard coded file
    but I would like to allow users to input a path to the file. Maybe even a
    folder that contains files.
    """
    try:
        file_with_part_numbers = open('Put_part_numbers_here.txt', 'r')
        for line in file_with_part_numbers:
            line = line.replace('\n','')
            part_list.append(line)
    except Exception as err:
        print("Part Numbers File Not Found! Creating... \n\nPlease paste part numbers in 'Put_part_numbers_here.txt' and restart program!")
        temp = open('Put_part_numbers_here.txt','w')

    pass


def format_model_number(model_numbers_to_format):
    """
    Takes a list of model numbers as a param and returns the list in an easy
    to read format.
    """
    formatted_model_numbers = []

    #cut the garbage away from our model numbers
    for model_number in model_numbers_to_format:
        model_number = model_number.upper()
        model_number = model_number.replace("-","").replace(" ","")
        model_number = model_number.replace("TONERCARTRIDGES,SUPPLIESANDPARTS","")
        model_number = model_number.replace("RICOH","")
        model_number = model_number.replace("LANIER","")
        model_number = model_number.replace("ALFICIO","")
        model_number = model_number.replace("AFICIO","")
        model_number = model_number.replace("SP","")
        model_number = model_number.replace("DN","")
        #cleaning data
        if not model_number.count("MP3003") and not  model_number.count("MP4503") and not  model_number.count("MPC80022"):
            formatted_model_numbers.append(model_number)

    #remove duplicates and return
    formatted_model_numbers = sorted(set(formatted_model_numbers))
    return(formatted_model_numbers)


def search_for_models(part_num):
    """
    Takes a part number as a param and returns a list of model numbers associated
    to it. This is done by searching precision roller's HTML.
    This is SPECIFIC to precision roller.
    """
    #list that will contain finalized part numbers
    part_numbers = []
    for i in range(len(part_num)):
        print("Searching for part:{}/{}...{}".format(i+1,len(part_num),str(part_num[i]).rjust(15,'.')))
        #Create a temporary list to store model numbers
        unformatted_model_numbers = []
        #Process part number into URL then search html with Beautiful Soup
        url = "https://www.precisionroller.com/search.php?q=" + part_num[i]
        request = requests.get(url)
        soup = bs(request.text,"html.parser")

        #Create a list to contain links that we find.
        links_from_url = soup.find_all("a")
        #Search the html for model numbers
        for link in links_from_url:
            if(str(link.get('title')).count('Toner Cartridges,')):
                if(str(link.get('title')).upper().count('RICOH') or str(link.get('title')).upper().count('LANIER') or str(link.get('title')).upper().count('ALFICIO')):
                    unformatted_model_numbers.append(link.get('title'))

        #Format our model numbers to be readable
        part_numbers.append(format_model_number(unformatted_model_numbers))

    return part_numbers

def search_for_part_description(url_to_search):
    """
    Takes a string as a param, searches a url for a description of the part
    Returns a list of descriptions. (Ill add this later)
    """
    pass


def export_to_xlsx(model_numbers_formatted,part_numbers):
    """
    Takes the finalized lists of model numbers and exports them into a nicely
    formatted excel document using pandas.
    Asks user what they want the new file to be called.
    """
    final_data_frame = pd.DataFrame({'Part Numbers':part_numbers},columns=['Part Numbers','Model Numbers'.ljust(150)])
    for i in range(len(model_numbers_formatted)):
        model_num_str = ""
        for model in model_numbers_formatted[i]:
            model_num_str=model_num_str+model+"/"
        final_data_frame.iat[i,1] = model_num_str

    path_to_exported_xlsx = "Untitled"
    userinput = input("\nSearch Complete!\n\nPlease name the excel document (do not include file extension):")
    if userinput:
        path_to_exported_xlsx = userinput

    #setup ExcelWriter object and use it to auto resize model_number column and then export
    writer = pd.ExcelWriter(path_to_exported_xlsx+".xlsx", engine='xlsxwriter')
    final_data_frame.to_excel(writer, sheet_name='Model Numbers')
    workbook = writer.book
    worksheet = writer.sheets['Model Numbers']

    #format column width and textwrap
    text_wrap = workbook.add_format({'text_wrap': True})
    worksheet.set_column(2,2, 150,text_wrap)
    worksheet.set_column(1,1,15)
    try:
        writer.save()
    except Exception as err:
        print("Error saving file! Maybe your filename is invalid?")

retrieve_part_numbers()
if(len(part_list)>0):
    export_to_xlsx(search_for_models(part_list), part_list)
else:
    print("Part list is empty! Please paste your part numbers 'Put_part_numbers_here.txt', save the file, and restart the program")
