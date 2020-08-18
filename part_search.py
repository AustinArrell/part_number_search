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


part_list = ["AE020266","AD027034","AE011131"]


def retrieve_part_numbers(filepath_to_parts_file):
    """
    Grabs the list of part numbers from a file. For now it is a hard coded file
    but I would like to allow users to input a path to the file. Maybe even a
    folder that contains files.
    """
    pass


def format_model_number(model_numbers_to_format):
    """
    Takes a list of model numbers as a param and returns the list in an easy
    to read format.
    """
    formatted_model_numbers = []

    #cut the garbage away from our model numbers
    for model_number in model_numbers_to_format:
        print(model_number)
        model_number = model_number.replace("-","").replace(" ","")
        model_number = model_number.replace("TonerCartridges,SuppliesandParts","")
        model_number = model_number.replace("Ricoh","")
        model_number = model_number.replace("Lanier","")
        model_number = model_number.replace("Alficio","")
        model_number = model_number.replace("Aficio","")
        model_number = model_number.replace("Savin","")
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
        print("Searching for part:{}/{}...{}".format(i,len(part_num),str(part_num[i]).rjust(15,'.')))
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
                unformatted_model_numbers.append(link.get('title'))

        #Format our model numbers to be readable
        part_numbers.append(format_model_number(unformatted_model_numbers))

    return part_numbers

def export_to_xlsx(model_numbers_formatted,part_numbers):
    """
    Takes the finalized lists of model numbers and exports them into a nicely
    formatted excel document using pandas.
    Asks user what they want the new file to be called.
    """
    print("\n")
    for i in range(len(model_numbers_formatted)):
        print("Model Numbers for part:{}".format(part_numbers[i]))
        for model in model_numbers_formatted[i]:
            print(model.rjust(20))
    pass


export_to_xlsx(search_for_models(part_list), part_list)
