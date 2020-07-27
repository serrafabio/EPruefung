import xml.etree.ElementTree as ET
from constants import *
import os

# this class is capable of creating file of importantion to the GMW
class GMW():
    def __init__(self, filepath, createTXT):


        self.filepath = filepath # File name
        if(createTXT):
            self.auxiliarFilePathRead = AUXILIAR_TXT_FILE_READ
            self.auxiliarFilePathWrite = AUXILIAR_TXT_FILE_WRITE
            self.createTXTFile()

        self.mydoc = ET.parse(self.filepath) # open file
        self.root = self.mydoc.getroot()

    def createTXTFile(self):
        """
        create a txt file, which will be edit by the software and then transformed once more in xml file
        :return: txt file exactly indentic of the xml file
        """
        writeFile = open(self.auxiliarFilePathRead, 'w', encoding= 'utf-8')
        with open(self.filepath, 'r', encoding='utf-8') as xmlFile:
            for lines in xmlFile:
                string, bool = '', False
                for i in range(len(lines)):
                    string += lines[i]
                    if(ord(lines[i]) == ord('/')):
                        bool = True
                    if(bool):
                        if (ord(lines[i]) == ord('>')):
                            string += '\n'
                            writeFile.write(string)
                            bool = False
                            string = ''
        xmlFile.close(); writeFile.close()

    def cleanName(self, string):
        """
        Auxiliar function to clean a string
        :param string: the string to clean
        :return:  afte the clean
        """
        for i in range(len(string)):
            if(ord(string[i]) == ord('}')):
                string = string[i+1:]
                break
        return string

    def readAllLocations(self):
        """
        The idea is to read the xml file and find the all locations, and with this locations atualize the Data Base
        :return: matrix (strings): contends the name of the of the building, floor and room availables in Grossen Metrawatt
        """
        try:
            locations = []
            for elements in self.root: # to find the floor
                for elem in elements:
                    for e in elem:
                        buildingLocation = self.cleanName(e.tag)
                        if(buildingLocation == READ_BUILDING_NAME): # insert the buildings
                            for buils in e:
                                for n in buils:
                                    if(self.cleanName(n.tag) == READ_ID_NAME):
                                        parameters = str(n.text)
                                        for item in buils:
                                            items = self.cleanName(item.tag)
                                            if(items == READ_ITEMS_NAME):
                                                for item_2 in item:  # Here we find the floor
                                                    for i_2 in item_2:
                                                        if(self.cleanName(i_2.tag) == READ_ID_NAME):
                                                            floors_name = str(i_2.text)
                                                            for n in item_2:
                                                                floorIdentenftification = self.cleanName(n.tag)
                                                                if (floorIdentenftification == READ_FLOOR_NAME):
                                                                    for rooms in n:  # Now we found the ID of the mensioned room
                                                                        for r in rooms:
                                                                            tag_name = self.cleanName(r.tag)
                                                                            if (tag_name == READ_ID_NAME_LOCATIONS):
                                                                                room_id_name = str(r.text)
                                                                            elif (tag_name == READ_ID_NAME):
                                                                                locations.append([parameters, floors_name,str(r.text), room_id_name]) # now we have all available inside the database of Metrawatt
            return locations # we return the name of the building, the floor name, the room name and room id, then we have all information to input inside the database of Grossen Metrawatt
        except: return False

    def getDataFromXMLFile(self, locationID):
        """
        A given ID of location is given, and the algorithm need to returns all the devices from this certain location
        :param locationID: string that contends the location of the
        :return: list that conteds the information about the device
        """
        try:
            devices, cont = [], 0
            for elements in self.root:  # to find the floo
                if(cont == 0):
                    for e in elements:
                        for element in e:
                            for element_2 in element:
                                for element_3 in element_2:
                                    if str(element_3.text) == locationID:
                                        sub_device, boolean = [], True
                                        for sub_element in element_2:
                                            name = self.cleanName(sub_element.tag)
                                            if(boolean):
                                                if(name == READ_NAME_NAME):
                                                    if (str(sub_element.text) == 'None'):sub_device.append('')
                                                    else:sub_device.append(str(sub_element.text))
                                                elif (name == READ_ID_NAME):
                                                    try:sub_device.append(int(sub_element.text))
                                                    except: boolean = False
                                                elif (name == READ_ROOMTYPE_NUMBER):
                                                    if(sub_element.text != None): sub_device.append(int(sub_element.text))
                                                    else: sub_device.append(OFFICE)
                                                elif (name == READ_COMPANY_NAME):
                                                    if (str(sub_element.text) == 'None'):sub_device.append('')
                                                    else:sub_device.append(str(sub_element.text))
                                                elif (name == READ_NOTE_NAME):
                                                    if(str(sub_element.text) == 'None'): sub_device.append('')
                                                    else: sub_device.append(str(sub_element.text))
                                                elif (name == READ_SERIALNUMBER):
                                                    if (str(sub_element.text) == 'None'):sub_device.append('')
                                                    else:sub_device.append(str(sub_element.text))
                                                elif (name == READ_LAST_TESTING_DATE): sub_device.append(str(sub_element.text)[:-9])
                                                elif (name == READ_NEXT_TESTING_DATE): sub_device.append(str(sub_element.text)[:-9])
                                        if(boolean): devices.append(sub_device)
                cont+=1
                if(cont >= 1): break
            if devices == list(): return None
            else: return devices
        except: False # If the file is wrong, False is returned

    def insertIntoXMLFile(self, location, name, manufacture, inv_num, note, room_type, id_device):
        """
        We insert the information into the file
        :param location: ID location to a room
        :param name: name of the device
        :param manufacture: manufaturer of the device
        :param inv_num: Serial number of the device
        :param note: notification about the device
        :param room_type: type of the room, where the device is located
        :param: id_device (string): id of the device
        :return: XML file update
        """
        def ConvertToReadableName(string):
            conversion, boolean = '', False
            for i in range(len(string)):
                if (boolean):
                    if ((string[i] == '>') or (string[i] == ' ')):
                        boolean = False;
                        break
                    conversion += string[i]
                if(string[i] == '<'):
                    boolean = True
            return conversion

        # first we need to find the id
        idFound, Slice ,readFile, writeFile = False, True, open(self.auxiliarFilePathRead, 'r', encoding='utf-8') , open(self.auxiliarFilePathWrite, 'w', encoding='utf-8') # Verify if the ID exist
        for elements in readFile: # to find the floor
            before = elements
            new = elements
            if(elements == ('<'+READ_ID_NAME+'>' + id_device +'</'+READ_ID_NAME+'>\n')):
                idFound = True # Id found
            if(idFound):
                namechar = ConvertToReadableName(elements)
                if(namechar == READ_MEASURMENT_NAME):
                    Slice = False
                elif(namechar == '/' + READ_MEASURMENT_NAME):
                    Slice = True
                if(Slice):
                    if (namechar == READ_NAME_NAME):
                        if (name != ''): new = '<'+READ_NAME_NAME+'>' + name + '</'+READ_NAME_NAME+'>\n'
                        else: new = '<'+READ_NAME_NAME+' i:nil="true"/>\n'
                    elif (namechar == READ_NOTE_NAME):
                        if (note != ''):new = '<'+READ_NOTE_NAME+'>' + note + '</'+READ_NOTE_NAME+'>\n'
                        else: new = '<'+READ_NOTE_NAME+ ' i:nil="true"/>\n'
                    elif (namechar == READ_LOCATION_NAME):
                        if (location != ''):new = '<'+READ_LOCATION_NAME+'>' + location + '</'+ READ_LOCATION_NAME +'>\n'
                        else:new = '<'+READ_LOCATION_NAME+ ' i:nil="true"/>\n'
                    elif (namechar == READ_COMPANY_NAME):
                        if (manufacture != ''):new = '<'+READ_COMPANY_NAME+'>' + manufacture + '</'+READ_COMPANY_NAME+'>\n'
                        else:new = '<'+ READ_COMPANY_NAME+ ' i:nil="true"/>\n'
                    elif (namechar == READ_SERIALNUMBER):
                        if (inv_num != ''): new = '<'+READ_SERIALNUMBER+'>' + inv_num + '</'+READ_SERIALNUMBER+'>\n'
                        else: new = '<'+READ_SERIALNUMBER+' i:nil="true"/>\n'
                    elif (namechar == READ_ROOMTYPE_NUMBER):
                        if (room_type != ''): new = '<'+READ_ROOMTYPE_NUMBER+'>' + room_type + '</'+READ_ROOMTYPE_NUMBER+'>\n'
                        else:new = '<'+READ_ROOMTYPE_NUMBER+' i:nil="true"/>\n'
                        idFound = False
            writeFile.write(elements.replace(before, new))
        readFile.close(), writeFile.close()
        # Atulize the old Read to the Write
        readFile, writeFile = open(self.auxiliarFilePathRead, 'w', encoding='utf-8') , open(self.auxiliarFilePathWrite, 'r', encoding='utf-8')
        for lines in writeFile:
            readFile.write(lines.replace(lines, lines))
        writeFile.close(), readFile.close()

    def TransformAuxiliarFileIntoDef(self):
        """
        This is the last function and we just transform txt file final to the real one
        :return: Files removed and File transforme
        """
        os.remove(self.filepath) # Remove Previous version
        os.remove(self.auxiliarFilePathRead) # Remove Read File
        os.rename(self.auxiliarFilePathWrite, self.filepath) # We just rename the txt to xml


