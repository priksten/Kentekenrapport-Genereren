from rdw.rdw import Rdw

import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showinfo

import pandas as pd

# Hier definieren we alle functies die we nodig hebben voor het programma
def create_car_info_dataframe(kenteken):
    """Deze functie neemt als een input het ingevoerde kenteken en als output een dataframe met auto-gegevens"""
    kenteken2 = format_kenteken(kenteken)
    car_info = get_kenteken_info(kenteken2)
    try:
        items = get_kenteken_keys(car_info)
    except AttributeError:
        print("Het door u ingevoerde kenteken kan niet worden gevonden")
    else:
        values = get_kenteken_values(car_info)
        data = make_dict(items, values)
        df = pd.DataFrame(data)
        print(df)

def create_tuple_list(kenteken):
    """Deze functie accepteert als invoer het door de gebruiker ingevoerde kenteken en retourneert een lijst met tuples"""
    kenteken2 = format_kenteken(kenteken)
    car_info = get_kenteken_info(kenteken2)
    items = get_kenteken_keys(car_info)
    values = get_kenteken_values(car_info)
    
    kenteken_data = []

    for n in range(0, len(items)):
            kenteken_data.append((items[n], values[n]))
    
    return kenteken_data

def create_category_tuple_list(keys, values):
    """Deze functie neemt als input de keys en de values en maakt hier tupels van"""
    category_data = []
    for n in range(0, len(keys)):
        category_data.append((keys[n], values[n]))
    return category_data

def create_category_dataframe(dict):
    "Deze functie maakt van de category_keys en de category_values een dataframe"
    df = pd.DataFrame(dict)
    return df

def create_treeview_table(kenteken_data, frame):
    """Deze functie neemt als input een naam, tkinter frame en een lijst met tuplesen een frame en retourneert vervolgens eeen treeview met de gewenste naam en de gewenste data erin"""
    columns = ('property', 'car_data')
    table = ttk.Treeview(frame, columns = columns, show = 'headings')

    table.heading('property', text = 'Property')
    table.heading('car_data', text = 'Car Data')

    for item in kenteken_data:
            table.insert('', tk.END, values = item)
    return table   


def make_dict(items, values):
    """Maakt een woordenboek met daarin de gewenste kenteken-info """
    data_car = {
        'Property' : items,
        'Car Data' : values
    }
    return data_car

def filter_info_category(car_info_keys, car_info_values, category_keys):
    """Deze functie haalt alle values die horen bij de keys van de aangegeven categorie en retourneert deze info"""
    category_keys_copy = category_keys[:]
        
    for key in category_keys:
        if key not in car_info_keys:
            category_keys_copy.remove(key)
    
    indices = []
    category_values = []

    for key in category_keys_copy:
        for i in range(len(car_info_keys)):
            if car_info_keys[i] == key:
                indices.append(i)
    
    for index in indices:
        value = car_info_values[index]
        category_values.append(value)
    
    category_dict = make_dict(category_keys_copy, category_values)

    return category_dict

def make_category_tuples(car_info_keys, car_info_values, category_keys):
    """Deze functie haalt alle values die horen bij de keys van de aangegeven categorie en retourneert deze info"""
    category_keys_copy = category_keys[:]
        
    for key in category_keys:
        if key not in car_info_keys:
            category_keys_copy.remove(key)
    
    indices = []
    category_values = []

    for key in category_keys_copy:
        for i in range(len(car_info_keys)):
            if car_info_keys[i] == key:
                indices.append(i)
    
    for index in indices:
        value = car_info_values[index]
        category_values.append(value)
    

    category_info = create_category_tuple_list(category_keys_copy, category_values)
    
    return category_info

def format_kenteken(kenteken):
    """Deze functie zet het door de gebruiker ingevoerde kenteken om in een juiste vorm"""
    kenteken1 = kenteken.replace("-","")
    kenteken2 = kenteken1.upper()
    return kenteken2

def get_category_keys(list_category_keys, properties, data):
    """Deze functie haalt uit de lijst met daarin alle keys, de keys die voor de ingevoerde categorie van belang zijn"""
    pass

def get_category_values(list_category_keys, data):
    """"""
    pass

def get_kenteken_info(kenteken):
    """Haalt de kentekengegevens op via de RDW API en retourneert deze (als deze gegevens gevonden worden) in een woordenboek"""
    car = Rdw()
    result = car.get_vehicle_data(kenteken)
    
    try:
        result_str = result[0]                        
    except IndexError:
        print("Het door u ingevoerde kenteken kan niet worden gevonden. Probeer opnieuw")
    else:
        return result_str      

def get_kenteken_keys(kenteken_info):
    """Deze functie haalt alle keys uit het woordenboek dat door get_kenteken_info wordt geretourneerd"""
    car_property = []
    for key in kenteken_info.keys():
        car_property.append(key)
    return car_property

def get_kenteken_values(kenteken_info):
    """De functie haalt alle values uit het woordenboek dat door get_kenteken_info wordt geretourneerd"""
    car_data = []
    for value in kenteken_info.values():
        car_data.append(value)
    return car_data

def make_dataframe(category_dict):
    """Deze functie neemt als invoer een dictionary, en zet deze om naar een dataframe"""
    dataframe = pd.DataFrame(category_dict)
    return dataframe

def show_data(kenteken_entry, show_info):
    kenteken = kenteken_entry.get()
    try:
        list = create_tuple_list(kenteken)
    except AttributeError:
        msg = "Het door u ingevoerde kenteken kan niet worden gevonden."
        showinfo(
            title = "Kenteken is niet gevonden", 
            message = msg
        )   
    else:
        car_info = create_treeview_table(list, show_info)
        show_treeview(show_info, car_info, 0, 0)

def show_treeview(frame, name, row, col):
    """Met deze functie plaats je een tabel met naam 'name' in row en col van het grid"""
    name.grid(row = int(row), column= int(col))
    scrollbar = ttk.Scrollbar(frame, orient = tk.VERTICAL, command = name.yview)
    name.configure(yscroll = scrollbar.set)
    scrollbar.grid(row = int(row), column = int(col) + 1)

def export_to_word(doc, category,dataframe):
        """Deze functie exporteert een gegeven dataframe naar Word"""
        doc.add_heading(category, level=1)
        t1 = doc.add_table(rows = dataframe.shape[0], cols = dataframe.shape[1])
        t1.style = 'TableGrid'

        for i in range(dataframe.shape[0]):
            for j in range(dataframe.shape[1]):
                cell = dataframe.iat[i,j]
                t1.cell(i,j).text = str(cell) 




