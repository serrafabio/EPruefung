"""
Here we will manipulate the Data Base. The idea is to have 2 Data Bases, the first is the official one, where the data
are that will be exported. And the second is where the current data will be manipulate till save.
@ Fabio serra Pereira
contact: serrafabio10@outlook.com
"""
import sqlite3
from sqlite3 import Error
from constants import *

class DB_Manipulation():
    def __init__(self):
        self.connection_official_DB = sqlite3.connect(NAME_OFFICIAL_DATA_BASE) # We open the data base
        self.connection_current_DB = sqlite3.connect(NAME_CURRENT_DATA_BASE)  # Finally we copy the DataBase for current Manipulation
        self.connection_official_DB.backup(self.connection_current_DB)  # Coping the Data Base
        self.cursor_officialDB = self.connection_official_DB.cursor() # Then we create the cursor, only used to save
        self.cursor_currentDB = self.connection_current_DB.cursor() # Creating the cursor to manipulation

    def load_building(self):
        """
        The goal is to return all building available
        :return: building_available (list(strings)): Name of the Buildings
        """
        self.cursor_currentDB.execute("SELECT building FROM Devices;"); building_available = self.cursor_currentDB.fetchall() # First we get only the string
        building_available = [building_available[0] for building_available in self.cursor_currentDB.execute("SELECT building FROM Devices;")] # Then we get all available
        building_available = list(dict.fromkeys(building_available)) # Remove repeated element
        self.connection_current_DB.commit()
        return building_available

    def load_floor(self, building_name):
        """
        The goal is to return the floors available for a certain building
        :return: floor_available (list(strings)): Name of the Floors available
        """
        self.cursor_currentDB.execute("SELECT floor FROM Devices WHERE building='%s';" %(building_name)); floor_available = self.cursor_currentDB.fetchall() # First we get only the string
        floor_available = [floor_available[0] for floor_available in self.cursor_currentDB.execute("SELECT floor FROM Devices WHERE building= '%s';" %(building_name))] # Then we get all available
        floor_available = list(dict.fromkeys(floor_available)) # Remove
        self.connection_current_DB.commit()
        return floor_available

    def load_rooms(self, building_name, floor_name):
        """
        The goal is to return all rooms available of a certain Floor of a Building
        :return: rooms_available (list(strings)): The rooms available
        """
        self.cursor_currentDB.execute("SELECT room FROM Devices WHERE building=:building_n AND floor=:floor_n;",{"building_n": building_name, "floor_n":floor_name}); room_available = self.cursor_currentDB.fetchall() # First we get only the string
        room_available =[room_available[0] for room_available in self.cursor_currentDB.execute("SELECT room FROM Devices WHERE building=:building_n AND floor=:floor_n;",{"building_n": building_name, "floor_n":floor_name})]# Then we get all available
        room_available = list(dict.fromkeys(room_available))
        self.connection_current_DB.commit()
        return room_available

    def load_room_id(self, building_name, floor_name, room_name):
        """
        The goal is to return the floors available for a certain building
        :return: floor_available (list(strings)): Name of the Floors available
        """
        self.cursor_currentDB.execute("SELECT room_id FROM Devices WHERE building=:building_n AND floor=:floor_n AND room=:room_n;", {"building_n": str(building_name), "floor_n": str(floor_name), "room_n":str(room_name)}); room_id_available = self.cursor_currentDB.fetchall()  # First we get only the string
        room_id_available = [room_id_available[0] for room_id_available in self.cursor_currentDB.execute("SELECT room_id FROM Devices WHERE building=:building_n AND floor=:floor_n AND room=:room_n;", {"building_n": str(building_name), "floor_n": str(floor_name), "room_n":str(room_name)})]  # Then we get all available
        room_id_available = list(dict.fromkeys(room_id_available))  # Remove
        self.connection_current_DB.commit()
        return room_id_available

    def load_devices(self, building_name, floor_name, room_name):
        """
        The goal is to return the devices of a selected room of selected floor of a selected Building
        :return:  devices (list(list())): Return the parameters of devices available for a certain room
        """
        devices = list()
        for row in self.cursor_currentDB.execute("SELECT * FROM Devices WHERE building=:b_name AND floor=:fl_name AND room =:r_name;", {"b_name":building_name, "fl_name":floor_name, "r_name":room_name}): devices.append(row) # Now we have access all ids available in the room
        self.connection_current_DB.commit()
        return devices

    def load_devices_via_Room_ID(self, room_id_name):
        """
        The goal is to return the devices of a selected room of selected floor of a selected Building
        :return:  devices (list(list())): Return the parameters of devices available for a certain room
        """
        devices = list()
        for row in self.cursor_currentDB.execute("SELECT * FROM Devices WHERE room_id =:r_name;", {"r_name":room_id_name}):
            devices.append(row) # Now we have access all ids available in the room
        self.connection_current_DB.commit()
        return devices

    def loadAll(self):
        """
        The goal is to return the all devices
        :return:  devices (list(list())): Return the parameters of devices available for a certain room
        """
        devices = list()
        for row in self.cursor_currentDB.execute("SELECT * FROM Devices;"):
            devices.append(row) # Now we have access all ids available in the room
        self.connection_current_DB.commit()
        return devices

    def insert_date_into_data_base(self, id, current_date, type_room, next_date_verification):
        """
        The goal is to insert the date of verification, the type of the room and next verification date to a certain id
        :return: The databank modified
        """
        self.cursor_currentDB.execute("UPDATE Devices SET verification_date= ?, room_type= ?, next_verification_date= ? WHERE id= ?;", (str(current_date), type_room, str(next_date_verification), id))
        self.connection_current_DB.commit()

    def insert_location_into_data_base(self, id, building, floor, room, room_id):
        """
        The goal is to insert the location into a element in data base
        :return: The databank modified
        """
        self.cursor_currentDB.execute("UPDATE Devices SET building= ?, floor = ?, room = ?, room_id = ? WHERE id= ?;", (str(building), str(floor), str(room), str(room_id), id))
        self.connection_current_DB.commit()

    def insertInformationIntoCurrent(self, id_name, name, company, inv_num, note):
        """
        The goal is to insert the information into a element in the data base
        :param
        :return: The databank modified
        """
        self.cursor_currentDB.execute("UPDATE Devices SET name= ?, company = ?, inv_num = ?, note = ? WHERE id= ?;", (str(name), str(company), str(inv_num) , str(note), id_name))
        self.connection_current_DB.commit()

    def inserNewLocation(self, locations):
        """
        Here we create a new Location. It is importanteto notice that
        :param locations: name of the room
        :return: The databank modified
        """
        for location in locations: # We insert and actualize the name of locations into the database
            exist = self.seeIfIDRoomexists(str(location[3]))
            if exist:
                for device in self.load_devices_via_Room_ID(str(location[3])): self.insert_location_into_data_base(device[0], location[0], location[1], location[2], location[3])
            else:
                exist_id, cont = True, -1
                while exist_id:
                    exist_id = self.seeIfIDexists(cont) # building, floor, room, room_id, type_room, date_now, next_date
                    if not exist_id: self.CreatenewID(cont, location[0], location[1], location[2], location[3] , 24 , '2020-6-12', '2022-6-12')
                    cont -= 1
        for device in self.loadAll(): # if there is no longer a room it will be deleted
            if not self.seeIfIDRoomexists(device[9]): self.deleteIDs(device[0])

    def CopyBackupToOfficial(self, name_file):
        """
        here we get a backup file and transform to the official DataBase
        :param name_file: The name of the file that we want to transform to official
        :return: Data Base Official changed and atualized
        """
        connection = sqlite3.connect(name_file)  # We open the data base of backUP
        connection.backup(self.connection_official_DB)  # Coping from the Backup to the official Data Base
        connection.commit()

    def CheckIFBackupISCompatible(self, name_file):
        """
        here we get a backup file and transform to the current DataBase
        :param name_file: The name of the file that we want to transform to current
        :return: Data Base Current changed and atualized
        """
        cursor = sqlite3.connect(name_file).cursor() # We open the data base of backUP
        cursor.execute("SELECT * FROM Devices;") # see if the table exist
        data = cursor.fetchall();
        if (len(data) > 0): return True  # it exist
        else: return False  # do not exist

    def CopyOfficialToBackUP(self, name_file):
        """
        here we get a backup file and transform to the official DataBase
        :param name_file: The name of the file that we want to transform in official
        :return: Data Base Official changed and atualized
        """
        connection = sqlite3.connect(name_file)  # We open the data base of backUP
        self.connection_official_DB.backup(connection)  # Coping from the official Data Base to the the Backup
        self.connection_official_DB.commit()

    def seeIfIDexists(self, id):
        """
        Here we see if the ID gave already exists.
        :param: Id that we want to see
        :return: the parameter referring to the ID
        :return: (boolean): saying if  not exist
        """
        self.cursor_currentDB.execute("SELECT * FROM Devices WHERE id =:id_name;",{"id_name": id})
        data = self.cursor_currentDB.fetchall();
        self.connection_current_DB.commit()
        if(len(data)>0): return True # it exist already
        else: return False # do not exist

    def seeIfIDRoomexists(self, id):
        """
        Here we see if the ID of Room gave already exists.
        :param: Room id that we want to see
        :return: the parameter referring to the ID
        :return: (boolean): saying if  not exist
        """
        self.cursor_currentDB.execute("SELECT * FROM Devices WHERE room_id =:id_name;",{"id_name": id})
        data = self.cursor_currentDB.fetchall();
        self.connection_current_DB.commit()
        if(len(data)>0): return True # it exist already
        else: return False # do not exist


    def CreatenewID(self, id, building, floor, room, room_id, type_room, date_now, next_date):
        """
        Here we create a new device
        :param id: the id of the new device
        :return: database atualized
        """
        self.cursor_currentDB.execute("INSERT INTO Devices(id, name, company, inv_num, note, verification_date, building, floor, room, room_id, room_type, next_verification_date) VALUES(?, '', '', '', '', ?, ?, ?, ?, ? , ? , ?);" , (id, str(date_now),str(building), str(floor), str(room), str(room_id) ,type_room, str(next_date)))
        self.connection_current_DB.commit()

    def deleteIDs(self, id):
        """
        Here we delete the ID that we don't want anymore
        :param id: Item that we don't want anymore
        :return: database atulized
        """
        self.cursor_currentDB.execute("DELETE FROM Devices WHERE id =:id_name;", {"id_name": id})
        self.connection_current_DB.commit()

    def deleteRooms(self, building, floor, room):
        """
        Here we delete the ids inside a room
        :param building: The name of the building wanted to delete
        :param floor: the name of the floor where we want to delete
        :param room: the room that we want to delete
        :return: Data base atualized
        """
        self.cursor_currentDB.execute("DELETE FROM Devices WHERE  building=:building_name and floor=:floor_name and room=:room_name;", {"building_name": building, "floor_name": floor, "room_name":room})
        self.connection_current_DB.commit()

    def UpdateLocation(self, old_name_building, oldfloor_name, new_name_building, new_name_floor, new_name_room):
        """
        We refactor the name of the locations
        :param old_name_building: name of old location
        :param old_name_floor: name of old location
        :param new_name_building: name of the new building
        :param new_name_floor: name of the new floor
        :param new_name_room: name of the new room
        :return: Data Base atualized
        """
        self.cursor_currentDB.execute("UPDATE Devices SET building= ?, floor = ?, room = ? WHERE building= ? AND floor = ?;",(str(new_name_building), str(new_name_floor), str(new_name_room), str(old_name_building), str(oldfloor_name)))
        self.connection_current_DB.commit()

    def SaveUndoAction(self, indexOfUndo):
        """
        This function save the previous database
        :param indexOfUndo: The current index of DataBase that we are working
        :return: Data Base saved and created
        """
        nameOffile = NAME_RECOVERY_DATA_BASE + str(indexOfUndo) + NAME_DB_END # we create the dataBase
        connection = sqlite3.connect(nameOffile)  # We open the data base of backUP
        self.connection_current_DB.backup(connection)  # Coping from the current BackUP to the Previous Data Base
        connection.commit()
        if(indexOfUndo >= LIMIT_UNDO_ACTIONS):
            os.remove(NAME_RECOVERY_DATA_BASE + str(indexOfUndo-LIMIT_UNDO_ACTIONS) + NAME_DB_END)
            return indexOfUndo-LIMIT_UNDO_ACTIONS
        else: return 0

    def RecoverUndoAction(self, indexOfUndo):
        """
        Here we recovery the previous action
        :param indexOfUndo: The current index of DataBase that we are working
        :return: Data Base uploaded
        """
        nameOffile = NAME_RECOVERY_DATA_BASE + str(indexOfUndo) + NAME_DB_END  # we get the name of Previous database
        connection = sqlite3.connect(nameOffile)  # We open the data base of backUP
        connection.backup(self.connection_current_DB)  # Coping from the previous BackUP to the current Data Base
        connection.commit()