import sqlite3


# C:\Users\EPruefung\Documents\Software_Development\venv\official_db.db
# Commands into Data Base:
# CREATE TABLE IF NOT EXISTS Devices(id INTEGER PRIMARY KEY, name TEXT, company TEXT, inv_num TEXT, note TEXT, verification_date TEXT, building TEXT, floor TEXT, room TEXT, room_id TEXT, room_type INTEGER, next_verification_date TEXT);
# INSERT INTO Devices (id, name, company, inv_num, note, verification_date, building, floor, room, room_id, room_type, next_verification_date) VALUES(1752, 'Desktop', 'Dell', '000000', '', '2020-05-01', 'L1|11 (ETA-Fabrik)', '1. OG', 'R107' , 'b80c5ee1-2dec-419b-b5bf-513880632202' ,24, '2022-05-01');
# if you have any problem execute these commands via sqlite3 in cmd
# C:\\Users\\EPruefung\\Desktop\\VDE-Sicherheitsprüfung\\Datenbank\\EPRUEFUNG BackupsBackUP_EPruefung-2020-5-3.db

auxiliar = 0
if auxiliar == 0:
    connection_official_DB = sqlite3.connect('official_db.db')
    cursor_officialDB = connection_official_DB.cursor()

    cursor_officialDB.execute("DROP TABLE Devices;"); # Clean Tables
    cursor_officialDB.execute("CREATE TABLE IF NOT EXISTS Devices(id INTEGER PRIMARY KEY, name TEXT, company TEXT, inv_num TEXT, note TEXT, verification_date TEXT, building TEXT, floor TEXT, room TEXT, room_id TEXT, room_type INTEGER, next_verification_date TEXT);"); # Clean Tables
    connection_official_DB.commit()
    cursor_officialDB.execute("INSERT INTO Devices (id, name, company, inv_num, note, verification_date, building, floor, room, room_id, room_type, next_verification_date) VALUES(1752, 'Desktop', 'Dell', '000000', '', '2020-05-01', 'L1|11 (ETA-Fabrik)', '1. OG', 'R107' , 'b80c5ee1-2dec-419b-b5bf-513880632202' ,24, '2022-05-01');"); # Clean Tables
    connection_official_DB.commit()
    cursor_officialDB.execute("SELECT * FROM Devices;")
    connection_official_DB.commit()
else:
    connection_official_DB = sqlite3.connect("undo\\recovery0.db")
    cursor_currentDB = connection_official_DB.cursor()
    cursor_currentDB.execute("SELECT building FROM Devices;");
    building_available = cursor_currentDB.fetchall()  # First we get only the string
    building_available = [building_available[0] for building_available in  cursor_currentDB.execute("SELECT building FROM Devices;")]  # Then we get all available
    building_available = list(dict.fromkeys(building_available))  # Remove repeated element
    print(building_available)
