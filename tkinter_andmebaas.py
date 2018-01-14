from tkinter import *
from tkinter import ttk
from PIL import Image, ImageTk
import openpyxl
import sqlite3
 
class ÕpilasteAB :
    #Kõik andmed alguses nulliks, et neid määrata
    # andmebaasi jaoks vajalik connection
    db_conn = 0
    # Cursor andmebaasi jaoks
    theCursor = 0
    # Hoiab õpilase ID-d
    curr_student = 0
    
    def setup_db(self):
        # Ava/Tee uus andmebaas
        self.db_conn = sqlite3.connect('UT_Andmebaas.db')
        self.theCursor = self.db_conn.cursor()
 
        # Try ja except kontrollib võimalikke erroreid
        try:
            self.db_conn.execute("CREATE TABLE if not exists Õpilased(ID INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, FName TEXT NOT NULL, LName TEXT NOT NULL, Klass TEXT NOT NULL, Spordiala TEXT NOT NULL, Sugu TEXT NOT NULL);")
 
            self.db_conn.commit()
 
        except sqlite3.OperationalError:
            print("ERROR : Tabelit ei olnud võimalik teha")
        #self.stud_from_excel() 
 
    def stud_submit(self):
        # Lisa õpilane andmebaasi
        self.db_conn.execute("INSERT INTO Õpilased (FName, LName, Klass, Spordiala, Sugu) " +
                             "VALUES ('" +
                             self.fn_entry_value.get() + "', '" +
                             self.ln_entry_value.get() + "', '" +
                             self.klass_entry_value.get() + "', '" +
                             self.spordiala_entry_value.get() + "', '" +
                             self.sugu_entry_value.get() + "')")

        
        # Tühjenda sisestamiseks mõeldud kastid
        self.fn_entry.delete(0, "end")
        self.ln_entry.delete(0, "end")
        self.klass_entry.delete(0, "end")
        self.spordiala_entry.delete(0, "end")
        self.sugu_entry.delete(0, "end")
        
        self.db_conn.commit()
        # Uuenda listboxi
        self.update_listbox()

         
    def update_listbox(self):
        # kustuta õpilane listboxist
        self.list_box.delete(0, END)
 
        # võta andmebaasist õpilane
        try:
            result = self.theCursor.execute("SELECT ID, FName, LName FROM Õpilased")
 
            for row in result:
                stud_id = row[0]
                stud_fname = row[1]
                stud_lname = row[2]
                # Pane muudetud õpilane listboxi
                self.list_box.insert(stud_id,
                                     stud_fname + " " +
                                     stud_lname)
 
        except sqlite3.OperationalError:
            print("Selline andmebaas ei eksisteeri")
 
        except:
            print("1: Ei saanud andmeid andmebaasist")
    
    def load_student(self, event=None):
        # Saa õpilase id
        lb_widget = event.widget
        index = str(lb_widget.curselection()[0] + 1)
 
        # Salvesta selle õpilase id
        self.curr_student = index
        try:
            result = self.theCursor.execute("SELECT ID, FName, LName, Klass, Spordiala, Sugu FROM Õpilased WHERE ID=" + index)
            for row in result:
                stud_id = row[0]
                stud_fname = row[1]
                stud_lname = row[2]
                stud_klass = row[3]
                stud_spordiala = row[4]
                stud_sugu = row[5]
                # pane vajutatud väärtused sisestamise kasti
                self.fn_entry_value.set(stud_fname)
                self.ln_entry_value.set(stud_lname)
                self.klass_entry_value.set(stud_klass)
                self.spordiala_entry_value.set(stud_spordiala)
                self.sugu_entry_value.set(stud_sugu)
        except sqlite3.OperationalError:
            print("Andmebaas ei eksisteeri")
        except:
            print("2 : Ei saanud andmeid andmebaasist")
 
    # Värskenda õpilaste listi
    def update_student(self, event=None):
        # uuenda andmeid
        try:
            self.db_conn.execute("UPDATE Õpilased SET FName='" +
                                self.fn_entry_value.get() +
                                "', LName='" +
                                self.ln_entry_value.get() +
                                "', Klass='" +
                                self.klass_entry_value.get() +
                                "', Spordiala='" +
                                self.spordiala_entry_value.get() +
                                "', Sugu='" +
                                self.sugu_entry_value.get() +
                                "' WHERE ID=" +
                                self.curr_student)
            self.db_conn.commit()
 
        except sqlite3.OperationalError:
            print("Andmebaasi ei olnud võimalik uuendada")
 
        # Tühjenda kastid
        self.fn_entry.delete(0, "end")
        self.ln_entry.delete(0, "end")
        self.klass_entry.delete(0, "end")
        self.spordiala_entry.delete(0, "end")
        self.sugu_entry.delete(0, "end")
 
        # Uuenda listboxi uute õpilastega
        self.update_listbox()
        
    def stud_from_excel(self):
        #Exceli failide töötlemine andmebaasi
        failinimi = self.xl_entry_value.get()
        wb = openpyxl.load_workbook(str(failinimi) + ".xlsx")
        wb_sheet = wb.active
        self.db_conn.execute("INSERT INTO Õpilased (FName, LName, Klass, Spordiala, Sugu) " +
                        "VALUES ('" +
                        wb_sheet["J10"].value + "', '" +
                        wb_sheet["J11"].value + "', '" +
                        wb_sheet["J12"].value + "', '" +
                        wb_sheet["J13"].value + "', '" +
                        wb_sheet["J14"].value + "')")
                        #wb_sheet["B3.."].value() ei tööta, sest funktsiooni ei saa
                        #siin kasutada, on vaja wb_sheet panna võrduma wb.active
                        #ehk praegune sheet, mis lahti excelis
        self.db_conn.commit()
        self.update_listbox()
        
    def stud_filter(self, ktg):
        #ktg- kategooria, mis tulbast väärtust otsida(value)
        value = self.flt_entry_value.get()
        self.db_conn.execute("SELECT " + str(ktg) + " FROM Õpilased WHERE "+ str(ktg) + " = " + str(value))
        print(ktg + " valitud")

        
    def open_filter_window(self):
        self.filter_window(root)
        
#---------------------------------------------------------------------
#---------------------------------------------------------------------
#----------------------GUI--------------------------------------------
#---------------------------------------------------------------------
#---------------------------------------------------------------------

    def __init__(self, root):
        root.title("JWG Andmebaas")
        root.geometry("500x600")
        root["bg"] = "white" #taust - valge
        
        
        # ----- ESIMENE RIDA -----
        # ----- ESIMENE RIDA -----
        # ----- ESIMENE RIDA -----
        # ----- ESIMENE RIDA -----
        # ----- ESIMENE RIDA -----
        
        scrollbar = Scrollbar(root)
 
        self.list_box = Listbox(root)
        self.list_box.bind('<<ListboxSelect>>', self.load_student)
        self.list_box.insert(1, "Õpilased...")
        self.list_box.grid(row=0, column=0, columnspan = 2, rowspan = 5, padx=10, pady=10, sticky = W+E+N+S)
 
        # andmebaas valmib funktsiooniga -> setup_db()
        self.setup_db()
        # Uuenda õpilaste listboxi
        self.update_listbox()
        
        
        fn_label = Label(root, text="Eesnimi", bg = "white")
        fn_label.grid(row=0, column=2, padx=10, pady=10, sticky=W)

        self.fn_entry_value = StringVar(root, value = "")
        self.fn_entry = ttk.Entry(root,
                                  textvariable=self.fn_entry_value)
        self.fn_entry.grid(row=0, column=3, padx=10, pady=10, sticky=W)
        
        # ----- TEINE RIDA -----
        #------ TEINE RIDA -----
        # ----- TEINE RIDA -----
        #------ TEINE RIDA -----
        # ----- TEINE RIDA -----
        #------ TEINE RIDA -----
        
        ln_label = Label(root, text="Perekonnanimi", bg = "white")
        ln_label.grid(row=1, column=2, padx=10, pady=10, sticky=W)
 
        # Siin on kasti sisestatud perekonnanime väärtus muutujana
        self.ln_entry_value = StringVar(root, value="")
        self.ln_entry = ttk.Entry(root,
                                  textvariable=self.ln_entry_value)
        self.ln_entry.grid(row=1, column=3, padx=10, pady=10, sticky=W)

        
        # ------KOLMAS RIDA-----
        # ------KOLMAS RIDA-----
        # ------KOLMAS RIDA-----
        # ------KOLMAS RIDA-----
        # ------KOLMAS RIDA-----
        
        klass_label = Label(root, text = "Klass", bg = "white")
        klass_label.grid(row = 2, column = 2, padx = 10, pady = 10, sticky = W)
        
        self.klass_entry_value = StringVar(root, value = "")
        self.klass_entry = ttk.Entry(root,
                                     textvariable = self.klass_entry_value)
        self.klass_entry.grid(row = 2, column = 3, padx = 10, pady = 10, sticky = E)
        
        # ----- NELJAS RIDA -----
        # ----- NELJAS RIDA -----
        # ----- NELJAS RIDA -----
        # ----- NELJAS RIDA -----
        # ----- NELJAS RIDA -----
        
        spordiala_label = Label(root, text = "Spordiala", bg = "white")
        spordiala_label.grid(row = 3, column = 2, padx= 10, pady= 10, sticky = W)
        
        self.spordiala_entry_value = StringVar(root, value = "")
        self.spordiala_entry = ttk.Entry(root,
                                         textvariable= self.spordiala_entry_value)
        self.spordiala_entry.grid(row = 3, column = 3, padx = 10, pady = 10, sticky = W)
        
        # ----- VIIES RIDA -----
        # ----- VIIES RIDA -----
        # ----- VIIES RIDA -----
        # ----- VIIES RIDA -----
        # ----- VIIES RIDA -----
 
        sugu_label = Label(root, text = "Sugu", bg = "white")
        sugu_label.grid(row = 4, column = 2, padx = 10, pady= 10, sticky = W)
        
        self.sugu_entry_value = StringVar(root, value = "")
        self.sugu_entry = ttk.Entry(root,
                                    textvariable = self.sugu_entry_value)
        self.sugu_entry.grid(row = 4, column = 3, padx = 10, pady= 10, sticky = E)
        
        # ----- KUUES RIDA -----
        # ----- KUUES RIDA -----
        # ----- KUUES RIDA -----
        # ----- KUUES RIDA -----
        # ----- KUUES RIDA -----
               
        #xl for excel
        xl_label = Label(root, text= "Exceli fail", bg = "white")
        xl_label.grid(row = 5, column = 0, padx = 10, pady = 10, sticky = W)
        
        self.xl_entry_value = StringVar(root, value = "")
        self.xl_entry = ttk.Entry(root,
                                  textvariable = self.xl_entry_value)
        self.xl_entry.grid(row = 5, column = 1, padx = 10, pady = 10, sticky = W)
        
        
        #Filtreerimise nupp ja entrybox
        flt_label = ttk.Button(root,
                               text = "Filtreeri",
                               command = lambda: self.open_filter_window())
        flt_label.grid(row = 5, column = 2,
                               padx = 10, pady = 10, sticky = W)
        
        self.flt_entry_value = StringVar(root, value = "")
        self.flt_entry = ttk.Entry(root,
                              textvariable = self.flt_entry_value)
        self.flt_entry.grid(row = 5, column = 3,
                               padx = 10, pady = 10, sticky = W)
        # ----- SEITSMES RIDA -----
        # ----- SEITSMES RIDA -----
        # ----- SEITSMES RIDA -----
        # ----- SEITSMES RIDA -----
        # ----- SEITSMES RIDA -----
        
        #Exceli nupp
        self.excel_button = ttk.Button(root,
                                       text = "Sisesta fail",
                                       command = lambda: self.stud_from_excel())
        self.excel_button.grid(row = 6, column = 0, columnspan = 2,
                               padx= 10, pady = 10, sticky = W + E)



        #Õpilase lisamise nupp
        self.submit_button = ttk.Button(root,
                            text="Lisa õpilane",
                            command=lambda: self.stud_submit())
        self.submit_button 
        self.submit_button.grid(row=6, column=2,
                                padx=10, pady=10, sticky=W)
        
        #Õpilase info muutmise nupp
        self.update_button = ttk.Button(root,
                            text="Muuda andmeid",
                            command=lambda: self.update_student())
        self.update_button.grid(row=6, column=3,
                                padx=10, pady=10,sticky = E)

        # ----- KAHEKSAS RIDA -----
        # ----- KAHEKSAS RIDA -----
        # ----- KAHEKSAS RIDA -----
        # ----- KAHEKSAS RIDA -----
        # ----- KAHEKSAS RIDA -----
    
    
    
    #-----------FILTER AKEN-------------
    #-----------FILTER AKEN-------------
    #-----------FILTER AKEN-------------
    #-----------FILTER AKEN-------------
    #-----------FILTER AKEN-------------
        
    def filter_window(self, root):
        Label(root, text= "filtreerimine..", bg = "white")
        aken_filter = Toplevel()
        aken_filter.title("Filtreeri")
        aken_filter.geometry("150x300")
        aken_filter["bg"] = "white"
        #aken_filter on uus root        
        
        self.filter_list_box = Listbox(aken_filter)
        
        #.insert(jrk number, text) lisab listboxi objekti
        self.filter_list_box.insert(0, "Eesnimi")
        self.filter_list_box.insert(1, "Perekonnanimi")
        self.filter_list_box.insert(2, "Klass")
        self.filter_list_box.insert(3, "Spordiala")
        self.filter_list_box.insert(4, "Sugu")
        self.filter_list_box.grid(row=0, column=0, columnspan = 2, rowspan = 5,
                           padx=10, pady=10, sticky = W+E)
        self.filter_list_box.bind("<<ListboxSelect>>", )
        
        #Filtreerimiseks vajalik nupp
        self.filter_window_button = ttk.Button(aken_filter, text = "Filtreeri",
                                          command = self.stud_filter(self.filter_list_box.curselection()))
        self.filter_window_button.grid(row = 5, column = 0, rowspan= 2,
                                       padx = 30, pady= 10, sticky = W + E)
        
# root - raam
root = Tk()
# Tee Andmebaas ekraanile
õpAB = ÕpilasteAB(root)
root.mainloop()
