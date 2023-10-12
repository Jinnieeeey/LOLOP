import pandas as pd
import numpy as np
from openpyxl import load_workbook

def input_reader():
    lst = []

    a = ""


    while(a != "."):
        a = input()
        aaa = '' .join((z for z in a if not z.isdigit()))
        lst.append(aaa.split())

    lst = sum(lst, [])
    lst[:] = (value for value in lst if value != "x")

    print(lst)

def file_reader():
    lst = ["x"]
    wb = load_workbook('LOL T.xlsx')
    ws = wb['History of Teams']
    not_allowed = ["Top", "Jungle", "Mid", "ADC", "Support", "Top Sub", "Jungle Sub", "Mid Sub", "ADC Sub", "Support Sub", "Head Coach", "Coach", "T", "E", "A", "M", "O", "R", "G", "N", "I", "Z", "D", "S", "B"
                   , "P", "U", "R", "C", "H", "", " ", "S1", "S2 Spring", "S2 Summer", "S3 Spring", "S3 Summer", "S4 Spring", "S4 Summer", "S5 Spring", "S5 Summer", "S6 Spring", "S6 Summer", "S7 Spring", 
                   "S7 Summer", "S8 Spring", "S8 Summer", "S9 Spring", "S9 Summer", "Demacia", "Jennis", "Adios", "Eagle United", "Aquila", "Fredeta Relmis", "Holy City", "Urkan", "South Demacia", "Shurelya", 
                   "Nomain", "Kingland", "Dominik City", "North Demacia", "Mount Demacia", "West Urkan", "Kalamanda", "Sterak", "Gargoyle", "Runann", "Westerley", "Alabaster", "Noxus", "Jericho City", "Black Rose", 
                   "Zaun City", "Darkin", "Gloria Swain", "Swana Gloria", "Diss County", "Colonized Ionia", "Colonized Placidium", "Dark Road", "Hades City", "Trath Union", "Hunter City", "Land of Skirmishers", 
                   "Cinderhulk", "Draktharr", "Warrior's House", "Wriggle", "Kitae", "Mardred", "Malmortius", "Guinsoo", "Kimir", "Basilich", "Piltover", "Nenu S United", "Kirchei", "Rylai", "Beograndin", "Reveck", 
                   "Bandle City", "Eleisa", "Zhonya", "Sudenphyran City", "Mercury City", "Nenu S City", "Bami", "Drawsmith", "Banshee", "Harmony City", "Aether", "Luden", "Zalie", "Arcana", "Shadow Isle", 
                   "Harrowing City", "Nest of the Dragon", "Ruined Kingdom", "Fire Union", "Forbidden Island", "East Shadow", "West Shadow", "Twisted Treeline", "Never Island", "Island of Death", "Rokfar", "Helia", 
                   "Zz'Rot", "Van Damn Island", "Liandry", "Nashor", "Vesani", "Isolde", "Bilgewater", "Blue Flame Island", "Phantom Island", "Nagakabouros", "Zeke", "Serpent City", "Statikk", "Conqueror's Ocean", 
                   "Titan City", "Guardian Island", "Freljord", "Avarosa", "Kaladoun", "Winter City", "Frozen City", "Frostguard", "Serpentine River", "Howling Abyss", "Grez", "Ironspike", "Ice Glen", 
                   "Winter's Claw", "Ionia", "Wooglet", "Ionian United", "Free Ionian United", "Liber Ioniamque", "Mount Targon", "Great Barrier", "Magni Obice", "Doran", "Kinkou", "Ionian Placidium", "Tiamat", 
                   "Shurima", "Fyrone", "Urtistan", "Shurima City", "Ancient Shurima", "Rabadon", "Randuin", "Ralsiji", "Targon", "Voodoo Lands", "Atreus", "Buhru", "Solari Council", "Rakkor Town", "Kumungu", 
                   "Sablestone", "Void", "Icathia", "Irem", "Youmuu", "Bubbling Bog", "Underground Road", "Tempest Flats", "Lavender Sea", "Xer'Sai Territory", "Zaun", "North Zaun", "The Gate", "South Zaun", 
                   "Morgon", "East Zaun", "West Zaun", "Entresol", "Bonscutt", "Bandle", "Civitas Parva", "Ruina River", "Meki", "Glade", "Atma", "Bandle Tree", "Humilis", "Porta", "Vastayan Lands", "Shimon", 
                   "Kiilash", "Lhotlan", "Marai", "Makara", "Skrad", "Juloah", "Vlotah", 0.0]
    wsdf = pd.DataFrame(ws.values)
    wsdf.to_numpy()
    ta = []
    shape = (len(wsdf), 1)
    ta = np.zeros(shape)
    da = np.append(wsdf, ta, axis = 1)
    for i in da:
        for j in i:
            #print(j)
            if j not in lst and j not in not_allowed:
                lst.append(j)
    print(lst)

file_reader()