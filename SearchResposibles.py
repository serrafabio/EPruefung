import requests_html  # requires urllib3
import requests
from constants import *
import sys
NUMBERS = ['0','1','2','3','4','5','6','7','8','9']


def cleanNameOfBuildingToRedable(location): # this function help, cleaning the name compared to the database and the website
    name, boolean = '', False
    for character in location:
        if(character == 'L'): boolean = True
        if(boolean):
            if(ord(character) != ord(' ')): name += character
            if(len(name) == 4): break
    return name

def cleanNameOfRoomToRedable(location): # this function help, cleaning the name compared to the database and the website
    name = ''
    for character in location:
        if NUMBERS.__contains__(character): name += character
    return name

def cleanNameOfRoomToRedable_2(location): # this function help, cleaning the name compared to the database and the website
    name, pointToStart = '', 0
    for c in range(len(location)-1,-1,-1):
        if (ord(location[c]) == ord(' ')): pointToStart = c+1; break
    for c in range(pointToStart,len(location)):
        if NUMBERS.__contains__(location[c]): name += location[c]
    return name

def CleanRepetitivePositions(value, lista, position): # Recursion Algorithm to clean repetitive names
    for j in range(position,-1,-1):
        if value == lista[j]: del lista[j]
    if position > 0:
        if position > len(lista)-1: position = len(lista)-1
        lista = CleanRepetitivePositions(lista[position], lista, position-1)
    return lista

def FindResponsibles():
    """
    The idea is to go to the link and give back the information required
    :return:  contacts: vector with several information, which is required to be cleaned
    """
    # Open the link
    with open(DEFAULT_FILE_WITH_WEBSITE, 'r', encoding='utf-8') as urlfile:
        for line in urlfile: url = str(line)
    try:
        response = requests.get(url)
        # Verify error
        response.raise_for_status()
    except Exception as err:
        return err  # return error if the link do not work anymore
    else:
        # Algorithm Here
        html, contacts = requests_html.HTML(html=response.text), []
        firstLines = html.find(HTML_MAIN)
        for line_1 in firstLines:
            for line_2 in line_1.find(HTML_SECTION):
                for line_3 in line_2.find(HTML_DIV):
                    for line_4 in line_3.find(HTML_TABLE):
                        for line_5 in line_4.find(HTML_TBODY):
                           for line_6 in line_5.find(HTML_TR):
                                for line_7 in line_6.find(HTML_TD):
                                    for line_8 in line_7.find(HTML_TD):
                                        contacts.append(line_8.text) # save the information of interest
        return contacts

def CleanContactsToNameandEmail(contacts):
    """
    Here we receives the contacts and clean them just to have the name, e-mail and location of the correspoding person.
    :param contacts: list with several information
    :return: clean_contacts: list of list, contains the name, email and location
    """

    sys.setrecursionlimit(1500) # Recursion Limit
    auxiliar, clean_contacts, name, cont = [], [], '', 1
    for c in contacts: name+= c +'\n' # First we save all in single line
    string, cont_1, cont_2, boolean = '', 1, 1, False
    for c in name: # then we separete into lines in one vector
        if c != '\n': string += c
        else: # save the data to the vector
            if string.rfind('...') != -1: string = string[:-3] + FINAL_EMAIL
            if string != '': auxiliar.append(string)
            string = ''
    for c in range(len(auxiliar)):
        if auxiliar[c][:3] == '+49':
            clean_contacts.append([auxiliar[c-1], auxiliar[c+1]])
    i = len(clean_contacts)-1
    if len(clean_contacts)>1:
        clean_contacts = CleanRepetitivePositions(clean_contacts[i], clean_contacts, i-1) # Clean repetitive names
    return clean_contacts # return

def FindResponsibleForEachRoom(locations, contacts):
    """
    The idea is to return all the contacts presented in th
    :param locations: here contains the building and room that we want to find the contacts
    :param contacts:  here contains the name, e-mail, telefon and location of the person
    :return: list with the name and telefon of the correspoding person
    """
    building, room, position = cleanNameOfBuildingToRedable(locations[0]), cleanNameOfRoomToRedable(locations[1]), []
    for i in range(len(contacts)):
        location_building, location_room = cleanNameOfBuildingToRedable(contacts[i][1]), cleanNameOfRoomToRedable_2(contacts[i][1])
        if (building == location_building) and (room == location_room):
            position.append([contacts[i][0], contacts[i][1]])
    return position
