import pandas as pd
import datetime as dt
from datetime import timedelta, datetime, date
import pandas as pd
import sys

GLN = ZENSIERT XXXX
Lieferantenname = "Coca Cola"


path = r"C:\Users\Maximilian.Rasch\Desktop\ZENSIERT"

# Enddatum Mai 2024

df = pd.read_excel(path)

def Liefertage_generieren(year1,year2, weekday):
    if weekday == "Monday":
        e = 7
    elif weekday == "Tuesday":
        e = 8
    elif weekday == "Wednesday":
        e = 9
    elif weekday == "Thursday":
        e = 10
    elif weekday == "Friday":
        e = 11

    d = date(year1, 1, 1)                    # January 1st
    d += timedelta(days = e - d.weekday())  # First Sunday
    while d.year == year1 or d.year == year2:
        yield d
        d += timedelta(days = 7)

def Bestelltag_generieren(year1,year2, weekday):
    if weekday == "Wednesday":
        e = 9
    elif weekday == "Thursday":
        e = 10
    elif weekday == "Friday":
        e = 11
    elif weekday == "Monday":
        e = 7
    elif weekday == "Tuesday":
        e = 8

    d = date(year1, 1, 1)                    # January 1st
    d += timedelta(days = e - d.weekday())  # First Sunday
    while d.year == year1 or d.year == year2:
        yield d
        d += timedelta(days = 7)


# LT Montag BT Mittwoch
Liefertag_Montag = [day.strftime('%Y%m%d') for day in Liefertage_generieren(2022,2023, "Monday")][1:]
Liefertag_Dienstag = [day.strftime('%Y%m%d') for day in Liefertage_generieren(2022,2023, "Tuesday")][1:]
Liefertag_Mittwoch = [day.strftime('%Y%m%d') for day in Liefertage_generieren(2022,2023, "Wednesday")][1:]
Liefertag_Donnerstag = [day.strftime('%Y%m%d') for day in Liefertage_generieren(2022,2023, "Thursday")]
Liefertag_Freitag = [day.strftime('%Y%m%d') for day in Liefertage_generieren(2022,2023, "Friday")]

Bestelltag_Mittwoch = [day.strftime('%Y%m%d') for day in Bestelltag_generieren(2022,2023, "Wednesday")]
Bestelltag_Donnerstag = [day.strftime('%Y%m%d') for day in Bestelltag_generieren(2022,2023, "Thursday")]
Bestelltag_Freitag = [day.strftime('%Y%m%d') for day in Bestelltag_generieren(2022,2023, "Friday")]
Bestelltag_Montag = [day.strftime('%Y%m%d') for day in Bestelltag_generieren(2022,2023, "Monday")]
Bestelltag_Dienstag = [day.strftime('%Y%m%d') for day in Bestelltag_generieren(2022,2023, "Tuesday")]




def output():
    for index, row in df.iterrows():
        if row["Montag"] == "X":
            for day, day2 in zip(Bestelltag_Mittwoch, Liefertag_Montag):
                day = day
                day2 = day2
                print(str(GLN) + ";" + "\"" + Lieferantenname + "\"" + ";" + "0" + str(row["ILN"]) + ";" + str(day) + ";" + str(day2))
        if row["Dienstag"] == "X":
            for day, day2 in zip(Bestelltag_Donnerstag, Liefertag_Dienstag):
                day = day
                day2 = day2
                print(str(GLN) + ";" + "\"" + Lieferantenname + "\"" + ";" + "0" + str(row["ILN"]) + ";" + str(day) + ";" + str(day2))
        if row["Mittwoch"] == "X":
            for day, day2 in zip(Bestelltag_Freitag, Liefertag_Mittwoch):
                day = day
                day2 = day2
                print(str(GLN) + ";" + "\"" + Lieferantenname + "\"" + ";" + "0" + str(row["ILN"]) + ";" + str(day) + ";" + str(day2))
        if row["Donnerstag"] == "X":
            for day, day2 in zip(Bestelltag_Montag, Liefertag_Donnerstag):
                day = day
                day2 = day2
                print(str(GLN) + ";" + "\"" + Lieferantenname + "\"" + ";" + "0" + str(row["ILN"]) + ";" + str(day) + ";" + str(day2))
        if row["Freitag"] == "X":
            for day, day2 in zip(Bestelltag_Dienstag, Liefertag_Freitag):
                day = day
                day2 = day2
                print(str(GLN) + ";" + "\"" + Lieferantenname + "\"" + ";" + "0" + str(row["ILN"]) + ";" + str(day) + ";" + str(day2))


class StdoutRedirection:
    """Standard output redirection context manager"""

    def __init__(self, path):
        self._path = path

    def __enter__(self):
        sys.stdout = open(self._path, mode="w")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        sys.stdout.close()
        sys.stdout = sys.__stdout__

with StdoutRedirection(r"C:\Users\Maximilian.Rasch\Desktop\Test123.txt"):
   output()