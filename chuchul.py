from csv import excel
import urllib.request
import re
import pandas as pd

mystr = ""


fp = urllib.request.urlopen("")
mybytes = fp.read()
mystr = mystr + mybytes.decode("utf8")

fp.close()

f = open("this.txt", "w", encoding="utf8")
f1 = open("that.txt", "w", encoding="utf8")

check = False
final_str = []
split = [a for a in re.split(r'(\s|\=)', mystr.strip()) if a]



for i in split:
    if (i == 'Page</span></div><div'):
        check = True
    if (i =='"license-description">'):
        check = False
    if (check):
        if (i not in [" ", "", "\n", "\t", "\s", "="]):
            final_str.append(i)



ca = ['None', 'Aatrox', 'Ahri', 'Akali', 'Akshan', 'Alistar', 'Amumu', 'Anivia', 'Annie', 'Aphelios', 'Ashe', 'Aurelion', 'Azir', 'Bard', "Bel'Veth", 'Blitzcrank', 'Brand', 'Braum', 
      'Caitlyn', 'Camille', 'Cassiopeia', "Cho'Gath", 'Corki', 'Darius', 'Diana', 'Draven', 'Dr.', 'Ekko', 'Elise', 'Evelynn', 'Ezreal', 'Fiddlesticks', 'Fiora', 'Fizz', 'Galio', 
      'Gangplank', 'Garen', 'Gnar', 'Gragas', 'Graves', 'Gwen', 'Hecarim', 'Heimerdinger', 'Illaoi', 'Irelia', 'Ivern', 'Janna', 'Jarvan', 'Jax', 'Jayce', 'Jhin', 'Jinx', "Kai'Sa", 
      'Kalista', 'Karma', 'Karthus', 'Kassadin', 'Katarina', 'Kayle', 'Kayn', 'Kennen', "Kha'Zix", 'Kindred', 'Kled', "Kog'Maw", 'LeBlanc', 'Lee', 'Leona', 'Lillia', 'Lissandra', 'Lucian', 
      'Lulu', 'Lux', 'Malphite', 'Malzahar', 'Maokai', 'Miss', 'Mordekaiser', 'Morgana', 'Nami', 'Nasus', 'Nautilus', 'Neeko', 'Nidalee', 'Nilah', 'Nocturne', 'Nunu', 'Olaf', 
      'Orianna', 'Ornn', 'Pantheon', 'Poppy', 'Pyke', 'Qiyana', 'Quinn', 'Rakan', 'Rammus', "Rek'Sai", 'Rell', 'Renata', 'Renekton', 'Rengar', 'Riven', 'Rumble', 'Ryze', 'Samira', 'Sejuani', 
      'Senna', 'Seraphine', 'Sett', 'Shaco', 'Shen', 'Shyvana', 'Singed', 'Sion', 'Sivir', 'Skarner', 'Sona', 'Soraka', 'Swain', 'Sylas', 'Syndra', 'Tahm', 'Taliyah', 'Talon', 'Taric', 'Teemo', 
      'Thresh', 'Tristana', 'Trundle', 'Tryndamere', 'Twisted', 'Twitch', 'Udyr', 'Urgot', 'Varus', 'Vayne', 'Veigar', "Vel'Koz", 'Vex', 'Vi', 'Viego', 'Viktor', 'Vladimir', 'Volibear', 
      'Warwick', 'Wukong', 'Xayah', 'Xerath', 'Xin', 'Yasuo', 'Yone', 'Yorick', 'Yuumi', 'Zac', 'Zed', 'Zeri', 'Ziggs', 'Zilean', 'Zoe', 'Zyra', "K'Sante", 'Milio', 'Naafiri', 'Master'] 

final_list = []
c = 0

for i in final_str:
    if (i[1:-1] in ca and i != '"Mastery' and i[0] == '"'):
        final_list.append(i[1:-1])
    elif (i[1:] in ca and i != '"Mastery'and i[0] == '"'):
        final_list.append(i[1:])
    elif("kda" in i and i[15:-10] != '">KDA'):
        final_list.append(i[15:-10])
    if("Team1Towers" in i):
        final_list.append(final_str[c + 8][:-10])
    if("Team2Towers" in i):
        final_list.append(final_str[c + 8][:-10])
    if("Team1Inhibitors" in i):
        final_list.append(final_str[c + 8][0])
    if("Team2Inhibitors" in i):
        final_list.append(final_str[c + 8][0])
    if("Team1Dragons" in i):
        final_list.append(final_str[c + 8][0])
    if("Team2Dragons" in i):
        final_list.append(final_str[c + 8][0])
    if("Team1Barons" in i):
        final_list.append(final_str[c + 8][0])
    if("Team2Barons" in i):
        final_list.append(final_str[c + 8][0])
        '''
    if("Team1Towers" in i):
        final_list.append(final_str[c + 16][:-10])
    if("Team2Towers" in i):
        final_list.append(final_str[c + 16][:-10])
    if("Team1Inhibitors" in i):
        final_list.append(final_str[c + 16][0])
    if("Team2Inhibitors" in i):
        final_list.append(final_str[c + 16][0])
    if("Team1Dragons" in i):
        final_list.append(final_str[c + 16][0])
    if("Team2Dragons" in i):
        final_list.append(final_str[c + 16][0])
    if("Team1Barons" in i):
        final_list.append(final_str[c + 16][0])
    if("Team2Barons" in i):
        final_list.append(final_str[c + 16][0])
        '''

    c += 1

excel_file = [["", "", "", "", "", "", "", "", "", "", ""]]

counter = 0

ta1 = []
ta2 = []

for i in range(len(final_list)):
    if (final_list[i] == "None"):
        final_list[i] = "x"
    elif (final_list[i] == "Jarvan"):
        final_list[i] = "Jarvan IV"
    elif (final_list[i] == "Miss"):
        final_list[i] = "Miss Fortune"
    elif (final_list[i] == "Aurelion"):
        final_list[i] = "Aurelion Sol"
    elif (final_list[i] == "Lee"):
        final_list[i] = "Lee Sin"
    elif (final_list[i] == "Twisted"):
        final_list[i] = "Twisted Fate"
    elif (final_list[i] == "Tahm"):
        final_list[i] = "Tahm Kench"
    elif (final_list[i] == "Xin"):
        final_list[i] = "Xin Zhao"
    elif (final_list[i] == "Dr."):
        final_list[i] = "Dr. Mundo"
    elif (final_list[i] == "Renata"):
        final_list[i] = "Renata Glasc"
    elif (final_list[i] == "Master"):
        final_list[i] = "Master Yi"
    elif (final_list[i] == "<"):
        final_list[i] = 0
lw = ""
rw = ""
for i in final_str:
    f.write(i + "\n")
f.close()
#for i in final_list:
    #f1.write(i + "\n")
#f1.close()
excel_file.append([11, 10, 7, 8, 9, 6, 3, 4, 5, 2, 1])
excel_file.append(["", "", "", "", "", "", "", "", "", "", ""])
for i in range(int(len(final_list)/38)):
    if (int(final_list[counter + 25]) > int(final_list[counter + 34])):
        lw = "W"
        rw = "L"
    else:
        lw = "L"
        rw = "W"
    excel_file.append(["", lw, "", "", "", "vs", "", "", "", rw, ""])
    ta1 = final_list[counter + 1].split("/")
    ta2 = final_list[counter + 11].split("/")
    excel_file.append(["", final_list[counter], int(ta1[0]), int(ta1[1]), int(ta1[2]), "", int(ta2[0]), int(ta2[1]), int(ta2[2]), final_list[counter + 10], ""])
    ta1 = final_list[counter + 3].split("/")
    ta2 = final_list[counter + 13].split("/")
    excel_file.append(["", final_list[counter + 2], int(ta1[0]), int(ta1[1]), int(ta1[2]), "", int(ta2[0]), int(ta2[1]), int(ta2[2]), final_list[counter + 12], ""])
    ta1 = final_list[counter + 5].split("/")
    ta2 = final_list[counter + 15].split("/")
    excel_file.append(["", final_list[counter + 4], int(ta1[0]), int(ta1[1]), int(ta1[2]), "", int(ta2[0]), int(ta2[1]), int(ta2[2]), final_list[counter + 14], ""])
    ta1 = final_list[counter + 7].split("/")
    ta2 = final_list[counter + 17].split("/")
    excel_file.append(["", final_list[counter + 6], int(ta1[0]), int(ta1[1]), int(ta1[2]), "", int(ta2[0]), int(ta2[1]), int(ta2[2]), final_list[counter + 16], ""])
    ta1 = final_list[counter + 9].split("/")
    ta2 = final_list[counter + 19].split("/")
    excel_file.append(["", final_list[counter + 8], int(ta1[0]), int(ta1[1]), int(ta1[2]), "", int(ta2[0]), int(ta2[1]), int(ta2[2]), final_list[counter + 18], ""])
    excel_file.append(["Ban", "", "", "", "", "", "", "", "", "", "Ban"])
    excel_file.append([final_list[counter + 20], "", "", int(final_list[counter + 25]), "", "", "", int(final_list[counter + 34]), "", "", final_list[counter + 29]])
    excel_file.append([final_list[counter + 21], "", "", int(final_list[counter + 26]), "", "", "", int(final_list[counter + 35]), "", "", final_list[counter + 30]])
    excel_file.append([final_list[counter + 22], "", "", int(final_list[counter + 28]), "", "", "", int(final_list[counter + 37]), "", "", final_list[counter + 31]])
    excel_file.append([final_list[counter + 23], "", "", int(final_list[counter + 27]), "", "", "", int(final_list[counter + 36]), "", "", final_list[counter + 32]])
    excel_file.append([final_list[counter + 24], "", "", "", "", "", "", "", "", "", final_list[counter + 33]])
    excel_file.append(["", "", "", "", "", "", "", "", "", "", ""])
    counter += 38


df = pd.DataFrame(excel_file)
df.to_excel(excel_writer = "chuchu.xlsx")
