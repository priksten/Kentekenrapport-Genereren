from rdw.rdw import Rdw
import pandas as pd
import docx
from docx2pdf import convert

import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showinfo

import functions
from functions import *

# Variabelen ten behoeve van filterfunctie
# Filteren op categorie
# Lijst met daarin de keys die horen bij elke categorie
# Vervolgens kan de informatie per categorie gefilterd en weergegeven worden met behulp van de functies
#       - filter_info_category(car_info_keys, car_info_values, category_keys)
#       - make_dataframe(category_dict)
list_basiskenmerken_keys = ['kenteken', 'voertuigsoort', 'merk', 'handelsbenaming',
                                    'bruto_bpm', 'inrichting', 'aantal_zitplaatsen', 'eerste_kleur', 'tweede_kleur', 'catalogusprijs',
                                    'aantal_staanplaatsen', 'aantal_deuren', 'aantal_wielen', 'type', 'typegoedkeuringsnummer', 'variant', 'uitvoering',
                                    'volgnummer_wijziging_eu_typegoedkeuring', 'aantal_rolstoelplaatsen', 'rupsonderstelconfiguratiecode', 'subcategorie_nederland',
                                    'aerodyn_voorz']
list_registratie_keys = ['vervaldatum_apk', 'datum_tenaamstelling', 'datum_eerste_toelating', 'datum_eerste_tenaamstelling_in_nederland','wacht_op_keuren',
                                    'wam_verzekerd', 'europese_voertuigcategorie', 'europese_voertuigcategorie_toevoeging', 'europese_uitvoeringcategorie_toevoeging',
                                    'plaats_chassisnummer', 'export_indicator', 'openstaande_terugroepactie_indicator', 'vervaldatum_tachograaf', 'taxi_indicator',
                                    'jaar_laatste_registratie_tellerstand', 'tellerstandoordeel', 'code_toelichting_tellerstandoordeel', 'tenaamstellen_mogelijk', 
                                    'vervaldatum_apk_dt', 'datum_tenaamstelling_dt', 'datum_eerste_toelating_dt', 'datum_eerste_tenaamstelling_in_nederland_dt', 
                                    'vervaldatum_tachograaf_dt', 'registratie_datum_goedkeuring_afschrijvingsmoment_bpm', 
                                    'registratie_datum_goedkeuring_afschrijvingsmoment_bpm_dt']        
list_motor_keys = ['aantal_cilinders', 'cilinderinhoud', 'maximale_constructiesnelheid', 'oplegger_geremd', 'aanhangwagen_autonoom_geremd',
                                    'aanhangwagen_middenas_geremd', 'afwijkende_maximum_snelheid', 'type_gasinstallatie', 'maximum_ondersteunende_snelheid',
                                    'type_remsysteem_voertuig_code', 'gem_lading_wrde', 'massa_alt_aandr']
list_milieu_keys = ['zuinigheidsclassificatie']
list_matengewichten_keys = ['massa_ledig_voertuig', 'toegestane_maximum_massa_voertuig', 'massa_rijklaar', 'maximum_massa_trekken_ongeremd',
                                    'maximum_trekken_massa_geremd', 'laadvermogen', 'afstand_hart_koppeling_tot_achterzijde_voertuig', 'afstand_voorzijde_voertuig_tot_hart_koppeling',
                                    'lengte', 'breedte', 'technische_max_massa_voertuig', 'vermogen_massarijklaar', 'wielbasis', 'maximum_massa_samenstelling', 
                                    'maximum_last_onder_de_vooras_sen_tezamen_koppeling', 'wielbasis_voertuig_minimum', 'wielbasis_voertuig_maximum', 'lengte_voertuig_minimum',
                                    'lengte_voertuig_maximum', 'breedte_voertuig_minimum', 'breedte_voertuig_maximum', 'hoogte_voertuig', 'hoogte_voertuig_minimum',
                                    'hoogte_voertuig_maximum', 'massa_bedrijfsklaar_minimaal', 'massa_bedrijfsklaar_maximaal', 'technisch_toelaatbaar_massa_koppelpunt',
                                    'maximum_massa_technisch_maximaal', 'maximum_massa_technisch_minimaal', 'verticale_belasting_koppelpunt_getrokken_voertuig', 
                                    'verl_cab_ind']

# Functies ten behoeve van het filteren


# We maken eerst een basic GUI
root = tk.Tk()
root.title('Zoek uw kenteken')
root.geometry('1400x800')
# root.resizable(False, False)

# We maken een frame waarin de gebruiker het kenteken kan invoeren
invoer_kenteken = ttk.Frame(root, height = 200, width = 300, relief= 'flat', padding = 10)

# We maken een frame waarin de kentekengegevens worden weergegeven
show_info = ttk.Frame(root, width = 300, height = 600, relief = 'flat', padding = 10)

# Label, invoervak en button

instructie = ttk.Label(invoer_kenteken, text = "Voer uw kenteken in: ")
instructie.grid(row = 0, column = 0, padx = 5, pady = 5, sticky = tk.W)

kenteken_entry = tk.StringVar()

input_kenteken = ttk.Entry(invoer_kenteken, textvariable=kenteken_entry)
input_kenteken.grid(row = 1, column = 0, padx = 5, pady = 5, sticky = tk.W)

basis_treeview = ttk.Treeview()
basiskenmerken_Label = None
scrollbar_basis = None

reg_treeview = ttk.Treeview()
registratie_Label= None
scrollbar_reg = None

mot_treeview = ttk.Treeview()
motor_Label = None
scrollbar_mot = None

mil_treeview = ttk.Treeview()
milieu_Label = None
scrollbar_mil = None

mat_treeview = ttk.Treeview()
matenGewichten_Label = None
scrollbar_mat = None

def create_data():
    global list
    global kenteken_info
    global car_info_keys
    global car_info_values
    global kenteken_def
    global mat_treeview
    global basis_treeview
    global reg_treeview
    global mil_treeview
    global mat_treeview

    kenteken = kenteken_entry.get()
    kenteken_def = format_kenteken(kenteken)

    # Wis de gegevens in de treeview en schakel de checkboxes uit
    # Verwijder alle items in de treeview
    for item in basis_treeview.get_children():
        basis_treeview.delete(item)
    
    for item in reg_treeview.get_children():
        reg_treeview.delete(item)
    
    for item in mot_treeview.get_children():
        mot_treeview.delete(item)
    
    for item in mil_treeview.get_children():
        mil_treeview.delete(item)
    
    for item in mat_treeview.get_children():
        mat_treeview.delete(item)


    show_basiskenmerken.set(0)
    show_registratie.set(0)
    show_motor.set(0)
    show_milieu.set(0)
    show_matengewichten.set(0)


    try:
        list = create_tuple_list(kenteken_def)
    except AttributeError:
        msg = "Het door u ingevoerde kenteken kan niet worden gevonden."
        showinfo(
            title = "Kenteken is niet gevonden", 
            message = msg
        )   
    else:
        kenteken_info = functions.get_kenteken_info(kenteken_def)
        car_info_keys  = functions.get_kenteken_keys(kenteken_info)
        car_info_values = functions.get_kenteken_values(kenteken_info)
        msg = str("Het door uw ingevoerde kenteken is: " + kenteken_def)
        showinfo(
            title = "Kenteken gevonden",
            message = msg
        )    

def show_tree_basis():
    value_basis = int(show_basiskenmerken.get())
    global basiskenmerken_Label
    global basis_treeview
    global scrollbar_basis
    global basis_treeview

    if basis_treeview is not None:
        basis_treeview.grid_forget()

    if value_basis == 1:
                    basis_values = make_category_tuples(car_info_keys, car_info_values, list_basiskenmerken_keys)
                    basis_treeview = create_treeview_table(basis_values, show_info)
                    basiskenmerken_Label = ttk.Label(show_info, text = "Basiskenmerken")
                    basiskenmerken_Label.grid(row = 0, column = 0, sticky = tk.W) 
                    basis_treeview.grid(row = 1, column= 0)
                    scrollbar_basis = ttk.Scrollbar(show_info, orient = tk.VERTICAL, command = basis_treeview.yview)
                    basis_treeview.configure(yscroll = scrollbar_basis.set)
                    scrollbar_basis.grid(row = 1, column = 1)
                                
    if value_basis == 0:
                    basiskenmerken_Label.grid_forget()
                    basis_treeview.grid_forget()
                    scrollbar_basis.grid_forget()
        
def show_tree_reg():
    value_reg = int(show_registratie.get())
    global registratie_Label
    global reg_treeview
    global scrollbar_reg
    global reg_treeview

    if reg_treeview is not None:
        reg_treeview.grid_forget()

    if value_reg == 1:
                    reg_values = make_category_tuples(car_info_keys, car_info_values, list_registratie_keys)
                    reg_treeview = create_treeview_table(reg_values, show_info)
                    registratie_Label = ttk.Label(show_info, text = "Registratie")
                    registratie_Label.grid(row = 2, column = 0, sticky = tk.W) 
                    reg_treeview.grid(row = 3, column = 0)
                    scrollbar_reg = ttk.Scrollbar(show_info, orient = tk.VERTICAL, command = reg_treeview.yview)
                    reg_treeview.configure(yscroll = scrollbar_reg.set)
                    scrollbar_reg.grid(row = 3, column = 1)
            
    if value_reg == 0:
                    registratie_Label.grid_forget()
                    reg_treeview.grid_forget()
                    scrollbar_reg.grid_forget()
            

def show_tree_mot():
    value_mot = int(show_motor.get())
    global motor_Label
    global mot_treeview
    global scrollbar_mot
    global mot_treeview

    if mot_treeview is not None:
        mot_treeview.grid_forget()

    if value_mot == 1:
                    mot_values = make_category_tuples(car_info_keys, car_info_values, list_motor_keys)
                    mot_treeview = create_treeview_table(mot_values, show_info)
                    motor_Label = ttk.Label(show_info, text = "Motor")
                    motor_Label.grid(row = 4, column = 0, sticky = tk.W)               
                    mot_treeview.grid(row = 5, column = 0)
                    scrollbar_mot = ttk.Scrollbar(show_info, orient = tk.VERTICAL, command = mot_treeview.yview)
                    mot_treeview.configure(yscroll = scrollbar_mot.set)
                    scrollbar_mot.grid(row = 5, column = 1)
            
    if value_mot == 0:
                    motor_Label.grid_forget()
                    mot_treeview.grid_forget()
                    scrollbar_mot.grid_forget()

def show_tree_mil():
    value_mil = int(show_milieu.get())
    global milieu_Label
    global mil_treeview
    global scrollbar_mil
    global mil_treeview

    if mil_treeview is not None:
        mil_treeview.grid_forget()

    if value_mil == 1:
                    mil_values = make_category_tuples(car_info_keys, car_info_values, list_milieu_keys)
                    mil_treeview = create_treeview_table(mil_values, show_info)
                    milieu_Label = ttk.Label(show_info, text = "Milieu")
                    milieu_Label.grid(row = 0, column = 2, sticky = tk.W)  
                    mil_treeview.grid(row = 1, column = 2)
                    scrollbar_mil = ttk.Scrollbar(show_info, orient = tk.VERTICAL, command = mil_treeview.yview)
                    mil_treeview.configure(yscroll = scrollbar_mil.set)
                    scrollbar_mil.grid(row = 1, column = 3)
            
    if value_mil == 0:
                    milieu_Label.grid_forget()
                    mil_treeview.grid_forget()
                    scrollbar_mil.grid_forget()            
        
def show_tree_mat():
    value_mat = int(show_matengewichten.get())
    global matenGewichten_Label
    global mat_treeview
    global scrollbar_mat
    global mat_treeview

    if mat_treeview is not None:
        mat_treeview.grid_forget()

    if value_mat == 1:
                    mat_values = make_category_tuples(car_info_keys, car_info_values, list_matengewichten_keys)
                    mat_treeview = create_treeview_table(mat_values, show_info)
                    matenGewichten_Label = ttk.Label(show_info, text = "Maten en Gewichten")
                    matenGewichten_Label.grid(row = 2, column = 2, sticky = tk.W)                 
                    mat_treeview.grid(row = 3, column = 2)
                    scrollbar_mat = ttk.Scrollbar(show_info, orient = tk.VERTICAL, command = mat_treeview.yview)
                    mat_treeview.configure(yscroll = scrollbar_mat.set)
                    scrollbar_mat.grid(row = 3, column = 3)
            
    if value_mat == 0:
                    matenGewichten_Label.grid_forget()
                    mat_treeview.grid_forget()
                    scrollbar_mat.grid_forget() 

kenteken_button = ttk.Button(invoer_kenteken, text="Zoek kenteken!", command= create_data)
kenteken_button.grid(row = 1, column=1, padx = 5, pady = 5,sticky = tk.W)

# We maken het frame 'invoer_kenteken' zichtbaar
invoer_kenteken.grid(row = 0, column = 0, sticky= tk.W)
show_info.grid(row = 0, column = 1, rowspan = 6, sticky = tk.W)

# We maken een frame voor waarbinnen de filter-opties zullen worden aangeboden.
filter_info = ttk.Frame(root, height = 300, width = 300, relief = 'flat', padding= 10)

# We maken nu de onderdelen van het filtermenu:
#   - Label met tekst 'Welke gegevens wilt u over het kenteken zien?'
#   - Twee radio buttons: 'Alle' en 'Filteren per categorie'
#   - Gewone button met tekst 'OK' 
# Als de gebruiker ervoor kiest om een keuze te maken uit de categorieen, dan komt er een derde frame.
# In dit derde frame is een lijst met checkboxen zichtbaar. De tekst bij de checkbox zijn de vijf categorieen met informatie
# We maken eerst een frame. 
# Vervolgens plaatsen we in dit frame vijf checkboxen en een button met de tekst 'Laat zien'          
categorie_filter = ttk.Frame(root, height = 300, width = 300, relief = 'flat', padding= 10)
categorie_label = ttk.Label(categorie_filter, text= "Welke categorieen wilt u zien?")
categorie_label.grid(row = 0, column = 0, padx = 0, pady = 5, sticky = tk.W)

show_basiskenmerken = tk.StringVar()
show_registratie = tk.StringVar()
show_motor = tk.StringVar()
show_milieu = tk.StringVar()
show_matengewichten = tk.StringVar()

         
checkbox_1 = ttk.Checkbutton(categorie_filter, text="Basiskenmerken", command = show_tree_basis, variable = show_basiskenmerken, onvalue= 1, offvalue= 0)
checkbox_2 = ttk.Checkbutton(categorie_filter, text="Registratie", command = show_tree_reg, variable = show_registratie, onvalue= 1, offvalue= 0)
checkbox_3 = ttk.Checkbutton(categorie_filter, text="Motor", command = show_tree_mot, variable = show_motor, onvalue= 1, offvalue= 0)
checkbox_4 = ttk.Checkbutton(categorie_filter, text="Milieu", command = show_tree_mil, variable = show_milieu, onvalue= 1, offvalue= 0)
checkbox_5 = ttk.Checkbutton(categorie_filter, text="Maten en Gewichten", command = show_tree_mat, variable = show_matengewichten, onvalue= 1, offvalue= 0)

checkbox_1.grid(sticky = tk.W)
checkbox_2.grid(sticky = tk.W)
checkbox_3.grid(sticky = tk.W)
checkbox_4.grid(sticky = tk.W)
checkbox_5.grid(sticky = tk.W)

categorie_filter.grid(row = 1, column = 0, sticky = tk.W)

filter_info.grid(row = 1, column = 0, sticky = tk.W)

# We maken een optie waarmee de gegevens naar Word en/of PDF kunnen worden gestuurd
# Het vierde frame biedt de gebruiker de mogelijkheid om de kentekengegevens (met bijbehorende filtering)
# te exporteren naar Word of naar Pdf. Hiervoor zijn nodig:
#       - frame: export_info
#       - label('Opties voor het exporteren van de kentekengegevens')
#       - buttons('Exporteer als Word' en 'Exporteer als PDF' )
#       - functie export_data die aangeroepen wordt wanneer er op de knop wordt gedrukt

def export_data():
        """Deze functie kopieert de gegevens naar Word"""
        doc = docx.Document()
        heading = str('Kentekenrapport voor kenteken: ' + kenteken_def )
        doc.add_heading(heading, 0)
        try:
            value_basis = int(show_basiskenmerken.get())
        except ValueError:
            pass
        else:
            if value_basis == 1:
                basis_dict = filter_info_category(car_info_keys, car_info_values, list_basiskenmerken_keys)
                df_basis = make_dataframe(basis_dict)
                export_to_word(doc, 'Basiskenmerken', df_basis)
                print(df_basis)
        
        try:
            value_reg = int(show_registratie.get())        
        except ValueError:
            pass
        else:        
            if value_reg == 1:
                reg_dict = filter_info_category(car_info_keys, car_info_values, list_registratie_keys)
                df_reg = make_dataframe(reg_dict)
                export_to_word(doc, 'Registratie', df_reg)
                print(df_reg)

        try:
            value_mot = int(show_motor.get())
        except ValueError:
            pass
        else:
            if value_mot == 1:
                mot_dict = filter_info_category(car_info_keys, car_info_values, list_motor_keys)
                df_mot = make_dataframe(mot_dict)
                export_to_word(doc, 'Motor', df_mot)
                print(df_mot)
        
        try:
            value_mil = int(show_milieu.get())
        except ValueError:
            pass
        else:
            if value_mil == 1:
                mil_dict = filter_info_category(car_info_keys, car_info_values, list_milieu_keys)
                df_mil = make_dataframe(mil_dict)
                export_to_word(doc, 'Milieu', df_mil)
                print(df_mil)
        
        try:
            value_mat = int(show_matengewichten.get())
        except ValueError:
            pass
        else:
            if value_mat == 1:
                mat_dict = filter_info_category(car_info_keys, car_info_values, list_matengewichten_keys)
                df_mat = make_dataframe(mat_dict)
                export_to_word(doc, 'Maten en Gewichten', df_mat)
                print(df_mat)    

        name_file = str(kenteken_def + '.docx') 
        doc.save(name_file) 

        name_file_pdf = str(kenteken_def + '.pdf')
        convert(name_file, name_file_pdf)          
   
export_info = ttk.Frame(root, height = 200, width = 300, relief = 'flat', padding= 10)

export_label = ttk.Label(export_info, text = "Opties voor het exporteren van de kentekengegevens: ")
export_label.grid(row = 0, column = 0, padx = 0, pady = 5, sticky = tk.W)

button_word_pdf = ttk.Button(export_info, text = "Exporteer gegevens als Word en PDF", command = export_data)
button_word_pdf.grid(row = 1, column = 0, padx = 5, pady = 5,sticky = tk.W)


export_info.grid(row = 3, column = 0, sticky = tk.W)


root.mainloop()