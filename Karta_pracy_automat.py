from datetime import datetime
from pathlib import Path
from calendar import monthrange
import os
import pandas as pd

months = ['0', 'Styczeń', 'Luty', 'Marzec', 'Kwiecień', 'Maj', 'Czerwiec', 'Lipiec', 'Sierpień', 'Wrzesień', 'Październik', 'Listopad', 'Grudzień']

date = datetime.now()
days_in_month = monthrange(date.year, date.month)[1]


class BlankObj:
    def __repr__(self):
        return ""

    
def list_weekend():
    """
    Function to check days when is saturday and sunday in current month.

    Input:

    None

    Output:

    Lists of saturdays and sundays in current month
    """
    DATE = 0
    saturdays = []
    sundays = []
    for days in range(1, days_in_month + 1):
        DATE = datetime(date.year, date.month, days)
        if DATE.strftime("%A") == 'Saturday':
            saturdays.append(days)
            continue
        else:
            saturdays.append(BlankObj())
            continue
    for days in range(1, days_in_month + 1):
        DATE = datetime(date.year, date.month, days)
        if DATE.strftime("%A") == 'Sunday':
            sundays.append(days)
            continue
        else:
            sundays.append(BlankObj())
            continue
    return saturdays, sundays
    

def off_work_days():
    """
    Function that collect data from user if he hasn't been in work.
    
    Input:

    None

    Output:

    Lists of days when you haven't been in work with 3 categories
    """
    off_days = []
    ill_days = []
    learn_days = []
    print("""

    Dni, w których Cię nie było z powodu (CTRL + C, aby przerwać):

    1. Urlopu
    2. Choroby
    3. Szkolenia
    
    """)
    while True:
        try:
            choice = input('Wybierz kategorię, z powodu której nie byłeś w pracy: ')
            if(choice == '1'):
                    while True:
                        try:
                            u_days = int(input('Podaj dzień, w którym nie byłeś w pracy (CTRL + C, aby przerwać): '))
                        except ValueError:
                            print('Dni muszą być liczbami!')
                        except KeyboardInterrupt:
                            off_days.pop(-1)
                            break
                        finally:
                            off_days.append(u_days)
            elif(choice == '2'):
                    while True:
                        try:
                            i_days = int(input('Podaj dzień, w którym nie byłeś w pracy (CTRL + C, aby przerwać): '))
                        except ValueError:
                            print('Dni muszą być liczbami!')
                        except KeyboardInterrupt:
                            ill_days.pop(-1)
                            break
                        finally:
                            ill_days.append(i_days)
            elif(choice == '3'):
                    while True:
                        try:
                            l_days = int(input('Podaj dzień, w którym nie byłeś w pracy (CTRL + C, aby przerwać): '))
                        except ValueError:
                            print('Dni muszą być liczbami!')
                        except KeyboardInterrupt:
                            learn_days.pop(-1)
                            break
                        finally:
                            learn_days.append(l_days)
            else:
                print('Nie wybrałeś nic z powyższego menu')
        except KeyboardInterrupt:
            return off_days, ill_days, learn_days
            break


def work_days(off_days_list, ill_days_list, learn_days_list, saturdays_list, sundays_list):
    """
    Function that removes days that appeared in off_work_days() from work_days list.

    Input:

    list1 (off days), list2 (illness), list3 (learning), list4 (list of saturdays), list5 (list of sundays)

    Output:

    list with days you have been in work
    """
    work_days = [day for day in range(1, days_in_month + 1)]
    for off_days in off_days_list:
        if off_days in work_days:
            work_days.remove(off_days)
    for ill_days in ill_days_list:
        if ill_days in work_days:
            work_days.remove(ill_days)
    for learn_days in learn_days_list:
        if learn_days in work_days:
            work_days.remove(learn_days)
    for saturdays in saturdays_list:
        if saturdays in work_days:
            work_days.remove(saturdays)
    for sundays in sundays_list:
        if sundays in work_days:
            work_days.remove(sundays)
    return work_days

    
def translated_month():
    """
    Function to translate current month to Polish with index of it.
    
    Input:

    None
    
    Output:

    Current month translated to Polish with index of it (ex. 6. June == 6. Czerwiec) 
    """
    for month in months:
        if int(date.strftime("%m")[1:]) == months.index(month):
            return f'{months.index(month)}. {str(month)}'
            break


def check_path():
    """
    Function to check if path with good variables exists and if true then change working directory to it.

    Input:

    None

    Output:

    Checks if dir with current month exists and change script working directory to it
    """
    path = r'C:\temp'
    if os.path.isdir(path):
        os.chdir(path)
    else:
        print(f'Nie odnaleziono ścieżki: {path}')


def dataframes_to_excel(dataframes, spaces):
        writer = pd.ExcelWriter(f'{name}_{surname}_karta_czasu_pracy_{date.year}_0{date.month}.xlsx')

        ROW = 0

        for df in dataframes:
            df.to_excel(writer, sheet_name='Result', startrow=ROW, startcol=0)
            ROW += ROW + len(df.index) + spaces + 2
        writer.save()


def auto_work_card_excel(name, surname, function, place):
    """
    Function that collect data from previous functions and generate excel file with filled tables of work days, information etc.

    Input:

    name, surname, function, place, off_work_days, ill_days, learning_days

    Output:

    Excel file with filled tables of work days. Made with Pandas.
    """
    full_off_days = off_work_days()
    weekend = list_weekend()
    full_work_days = work_days(full_off_days[0], full_off_days[1], full_off_days[2], weekend[0], weekend[1])

    Name_Dataframe = pd.DataFrame({

    'Imię i nazwisko': [name + ' ' + surname]
    
    })

    Role_Dataframe = pd.DataFrame({

    'Pełniona funkcja': [function]
    
    })
    
    Main_Dataframe = pd.DataFrame({

    'Dzień': [day for day in range(1, days_in_month + 1)],
    'Miejsce świadczenia usług': [place if day not in full_off_days[0] or day not in full_off_days[1] or day not in full_off_days[2] else BlankObj() for day in range(1, days_in_month + 1)],
    'Opis zadań wykonywanych': ['Wsparcie działu IT' if day not in full_off_days[0] or day not in full_off_days[1] or day not in full_off_days[2] else BlankObj() for day in range(1, days_in_month + 1)],
    'N Ś': ['X' if day in weekend[1] else BlankObj() for day in range(1, days_in_month + 1)],
    'S': ['X' if day in weekend[0] else BlankObj() for day in range(1, days_in_month + 1)],
    'U': ['X'  if day in full_off_days[0] else BlankObj() for day in range(1, days_in_month + 1)],
    'Ch': ['X'  if day in full_off_days[1] else BlankObj() for day in range(1, days_in_month + 1)],
    'Sz': ['X'  if day in full_off_days[2] else BlankObj() for day in range(1, days_in_month + 1)],
    'P': ['X' if day in full_work_days else BlankObj() for day in range(1, days_in_month + 1)]
    
    })

    Dataframes = (Name_Dataframe, Role_Dataframe, Main_Dataframe)
    dataframes_to_excel(Dataframes, 1)
    
    
check_path()

while True:
    name = input('Podaj swoje imie: ')
    surname = input('Podaj swoje nazwisko: ')
    function = input('Podaj stanowisko: ')
    place = input('Podaj miejsce pracy: ')
    if name and surname and function and place:
        auto_work_card_excel(name, surname, function, place)
        print('Utworzenie pliku powiodło się')
        break
    else:
        print('Wszystkie pola muszą zostać wypełnione')






















