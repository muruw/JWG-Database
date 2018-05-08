from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from PIL import Image, ImageTk
import openpyxl
import sqlite3


class ÕpilasteAB :
    #Kõik andmed alguses nulliks, et neid määrata
    # andmebaasi jaoks vajalik connection
    db_conn = 0
    # Cursor andmebaasi jaoks
    c = 0
    # Hoiab õpilase ID-d
    curr_student = 0
    
    def setup_db(self):
        # Ava/Tee uus andmebaas
        self.db_conn = sqlite3.connect('UT_Andmebaas.db')
        self.c = self.db_conn.cursor()
 
        # Try ja except kontrollib võimalikke erroreid
        try:
            self.db_conn.execute("CREATE TABLE if not exists Õpilased(ID INTEGER PRIMARY KEY AUTOINCREMENT, FName TEXT, "
                                 "LName TEXT, Klass INTEGER, Spordiala TEXT, Sugu TEXT, Length INTEGER, "
                                 "Club TEXT, Freq INTEGER, Trainer TEXT, Achiev TEXT);")
 
            self.db_conn.commit()
 
        except sqlite3.OperationalError:
            print("ERROR : Tabelit ei olnud võimalik teha")
        #self.stud_from_excel() 
 
    def stud_submit(self):
        # Lisa õpilane andmebaasi
        self.db_conn.execute("INSERT INTO Õpilased (FName, LName, Klass, Spordiala, Sugu, Length, Club, Freq, Trainer, Achiev) " +
                             "VALUES ('" +
                             self.fn_entry_value.get() + "', '" +
                             self.ln_entry_value.get() + "', '" +
                             self.klass_entry_value.get() + "', '" +
                             self.spordiala_entry_value.get() + "', '" +
                             self.sugu_entry_value.get() + "', '" +
                             self.length_entry_value.get() + "', '" +
                             self.club_entry_value.get() + "', '" +
                             self.freq_entry_value.get() + "', '" +
                             self.trainer_entry_value.get() + "', '" +
                             self.achiev_entry_value.get() + "')")

        
        # Tühjenda sisestamiseks mõeldud kastid
        self.fn_entry.delete(0, "end")
        self.ln_entry.delete(0, "end")
        self.klass_entry.delete(0, "end")
        self.spordiala_entry.delete(0, "end")
        self.sugu_entry.delete(0, "end")
        self.length_entry.delete(0, "end")
        self.club_entry.delete(0, "end")
        self.freq_entry.delete(0, "end")
        self.trainer_entry.delete(0, "end")
        self.achiev_entry.delete(0, "end")
        
        self.db_conn.commit()
        # Uuenda listboxi
        self.update_listbox()

         
    def update_listbox(self):
        # kustuta õpilane listboxist
        self.list_box.delete(0, END)
 
        # võta andmebaasist õpilane
        try:
            result = self.c.execute("SELECT ID, FName, LName FROM Õpilased")
 
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
        #try:
        result = self.c.execute("SELECT ID, FName, LName, Klass, Spordiala, Sugu, Length, Club, Freq, Trainer, Achiev FROM Õpilased WHERE ID=" + index)
        for row in result:
            stud_id = row[0]
            stud_fname = row[1]
            stud_lname = row[2]
            stud_klass = row[3]
            stud_spordiala = row[4]
            stud_sugu = row[5]
            stud_length = row[6]
            stud_club = row[7]
            stud_freq = row[8]
            stud_trainer = row[9]
            stud_achiev = row[10]
            
                # pane vajutatud õpilase väärtused sisestamise kasti
            self.fn_entry_value.set(stud_fname)
            self.ln_entry_value.set(stud_lname)
            self.klass_entry_value.set(stud_klass)
            self.spordiala_entry_value.set(stud_spordiala)
            self.sugu_entry_value.set(stud_sugu)
            self.length_entry_value.set(stud_length)
            self.club_entry_value.set(stud_club)
            self.freq_entry_value.set(stud_freq)
            self.trainer_entry_value.set(stud_trainer)
            self.achiev_entry_value.set(stud_achiev)
##        except sqlite3.OperationalError:
##            print("Õpilase kättesaamisel tekib viga")
##        except:
##            print("2 : Ei saanud andmeid andmebaasist")
 
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
                                "', Length='" +
                                self.length_entry_value.get() +
                                "', Club='" +
                                self.club_entry_value.get() +
                                "', Freq='" +
                                self.freq_entry_value.get() +
                                "', Trainer='" +
                                self.trainer_entry_value.get() +
                                "', Achiev='" +
                                self.achiev_entry_value.get() +
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
        self.length_entry.delete(0, "end")
        self.club_entry.delete(0, "end")
        self.freq_entry.delete(0, "end")
        self.trainer.delete(0, "end")
        self.achiev.delete(0, "end")
 
        # Uuenda listboxi uute õpilastega
        self.update_listbox()
        
    def stud_from_excel(self):
        #Exceli failide töötlemine andmebaasi
        fileName = filedialog.askopenfilename(filetypes = (("Excel file", "*.xlsx"), ("All files", "*.*")))
        #failinimi = self.xl_entry_value.get()
        wb = openpyxl.load_workbook(fileName, data_only = True)
        wb_sheet = wb.active

        data_list = []
        empty_string = ""
            
        for m in range(0, 10):
            if wb_sheet["J1"+str(m)].value:
                data_list.append(wb_sheet["J1"+str(m)].value)
            else:
                data_list.append(empty_string)
                
        
        self.db_conn.execute("INSERT INTO Õpilased (FName, LName, Klass, Spordiala, Sugu, Length, Club, Freq, Trainer, Achiev) "
                        "VALUES ('" +
                        data_list[0] + "', '" +
                        data_list[1] + "', '" +
                        data_list[2] + "', '" +
                        data_list[3] + "', '" +
                        data_list[4] + "', '" +
                        data_list[5] + "', '" +
                        data_list[6] + "', '" +
                        data_list[7] + "', '" +
                        data_list[8] + "', '" +
                        data_list[9] + "')")
                        #wb_sheet["B3.."].value() ei tööta, sest funktsiooni ei saa
                        #siin kasutada, on vaja wb_sheet panna võrduma wb.active
                        #ehk praegune sheet, mis lahti excelis
        self.db_conn.commit()
        self.update_listbox()        
        
    def search_student_command(self):
        #Puhastan kõik värvidest
        ÕpilasteAB.update_listbox(self)

        #Kasutaja otsitud õpilased värvitakse roheliseks
        #s_t - s_t 
        s_t = self.flt_entry_value.get()
        self.c.execute("SELECT ID FROM Õpilased WHERE FName = ? OR LName = ? OR Spordiala = ? OR Klass = ? OR Sugu = ? OR Length = ? OR Club = ? OR Freq = ? OR Trainer = ? OR Achiev = ?",
                       (s_t, s_t, s_t, s_t, s_t, s_t, s_t, s_t, s_t, s_t))
        result = self.c.fetchall()
        for stud in result:
            self.list_box.itemconfig(list(stud)[0] - 1, {"bg":"PaleVioletRed2"})
        
 
 
#---------------------------------------------------------------------
#---------------------------------------------------------------------
#----------------------GUI--------------------------------------------
#---------------------------------------------------------------------
#---------------------------------------------------------------------

    def __init__(self, root):
        root.title("JWG Andmebaas")
        root.geometry("764x634")
        root.resizable(width = False, height = False)
        root["bg"] = "white" #taust - valge

        #Taustapilt
        bg_image = PhotoImage(file= "ut_taust4_fixed.png")
        x = Label(root, image = bg_image)
        x.photo=bg_image
        x.place(x = 0, y = 0)
        
        # ----- ESIMENE RIDA -----
        # ----- ESIMENE RIDA -----
        # ----- ESIMENE RIDA -----
        # ----- ESIMENE RIDA -----
        # ----- ESIMENE RIDA -----
 
        self.list_box = Listbox(root)
        #bind - kui vajutad listboxi ehk kasutad(<<ListboxSelect>>), siis kutsud funktsiooni load.student
        self.list_box.bind('<<ListboxSelect>>', self.load_student)
        self.list_box.insert(1, "Õpilased...")
        self.list_box.grid(row=0, column=0, columnspan = 2, rowspan = 5, padx=10, pady=10, sticky = W+E+N+S)
 
        #Andmebaas valmib funktsiooniga -> setup_db()
        #Uuenda õpilaste listboxi
        self.setup_db()
        self.update_listbox()
        
        #Eesnimi
        fn_label = Label(root, text="Eesnimi", bg = "white")
        fn_label.grid(row = 0, column = 2, padx = 10, pady = 10, sticky = W)

        self.fn_entry_value = StringVar(root, value = "")
        self.fn_entry = ttk.Entry(root,
                                  textvariable = self.fn_entry_value)
        self.fn_entry.grid(row = 0, column = 3, padx = 10, pady = 10, sticky = W)
        
        #Perekonnanimi
        ln_label = Label(root, text="Perekonnanimi", bg = "white")
        ln_label.grid(row=0, column=4, padx=10, pady=10, sticky=W)
        
        self.ln_entry_value = StringVar(root, value="")
        self.ln_entry = ttk.Entry(root,
                                  textvariable=self.ln_entry_value)
        self.ln_entry.grid(row=0, column=5, padx=10, pady=10, sticky=W)
        
        # ----- TEINE RIDA -----
        #------ TEINE RIDA -----
        #------ TEINE RIDA -----
        # ----- TEINE RIDA -----
        #------ TEINE RIDA -----
        
        #Sugu
        sugu_label = Label(root, text = "Sugu", bg = "white")
        sugu_label.grid(row = 1, column = 2, padx = 10, pady= 10, sticky = W)
        
        self.sugu_entry_value = StringVar(root, value = "")
        self.sugu_entry = ttk.Entry(root,
                                    textvariable = self.sugu_entry_value)
        self.sugu_entry.grid(row = 1, column = 3, padx = 10, pady= 10, sticky = E)

        #Mitmendas klassis käib
        klass_label = Label(root, text = "Klass", bg = "white")
        klass_label.grid(row = 1, column = 4, padx = 10, pady = 10, sticky = W)
        
        self.klass_entry_value = StringVar(root, value = "")
        self.klass_entry = ttk.Entry(root,
                                     textvariable = self.klass_entry_value)
        self.klass_entry.grid(row = 1, column = 5, padx = 10, pady = 10, sticky = E)
        
        # ------KOLMAS RIDA-----
        # ------KOLMAS RIDA-----
        # ------KOLMAS RIDA-----
        # ------KOLMAS RIDA-----
        # ------KOLMAS RIDA-----
        
        #Spordiala
        spordiala_label = Label(root, text = "Spordiala", bg = "white")
        spordiala_label.grid(row = 2, column = 2, padx= 10, pady= 10, sticky = W)
        
        self.spordiala_entry_value = StringVar(root, value = "")
        self.spordiala_entry = ttk.Entry(root,
                                         textvariable= self.spordiala_entry_value)
        self.spordiala_entry.grid(row = 2, column = 3, padx = 10, pady = 10, sticky = W)
        
        #Klubi olemasolu
        club_label = Label(root, text = "Klubi", bg = "white")
        club_label.grid(row = 2, column = 4, padx=10, pady=10, sticky=W)
        
        self.club_entry_value = StringVar(root, value = "")
        self.club_entry = ttk.Entry(root, textvariable = self.club_entry_value)
        self.club_entry.grid(row = 2, column = 5, padx=10, pady=10, sticky=W)
        
        # ----- NELJAS RIDA -----
        # ----- NELJAS RIDA -----
        # ----- NELJAS RIDA -----
        # ----- NELJAS RIDA -----
        # ----- NELJAS RIDA -----
        
        #Mitu korda nädalas trenn
        freq_label = Label(root, text="Treeningute arv", bg= "white")
        freq_label.grid(row = 3, column = 2, padx=10, pady=10, sticky=W)
        
        self.freq_entry_value = StringVar(root, value = "")
        self.freq_entry = ttk.Entry(root, textvariable = self.freq_entry_value)
        self.freq_entry.grid(row = 3, column = 3, padx=10, pady=10, sticky=W)
        
        #Treeneri olemasolu
        trainer_label = Label(root, text = "Treener", bg = "white")
        trainer_label.grid(row = 3, column = 4, padx=10, pady=10, sticky=W)
        
        self.trainer_entry_value = StringVar(root, value = "")
        self.trainer_entry = ttk.Entry(root,
                                       textvariable = self.trainer_entry_value)
        self.trainer_entry.grid(row = 3, column = 5, padx=10, pady=10, sticky=W)
        
        # ----- VIIES RIDA -----
        # ----- VIIES RIDA -----
        # ----- VIIES RIDA -----
        # ----- VIIES RIDA -----
        # ----- VIIES RIDA -----
        
        #Kaua tegelenud
        length_label = Label(root, text = "Kaua tegelenud", bg = "white")
        length_label.grid(row = 4, column = 2, padx=10, pady=10, sticky=W)
        
        self.length_entry_value = StringVar(root, value = "")
        self.length_entry = ttk.Entry(root,
                                      textvariable = self.length_entry_value)
        self.length_entry.grid(row = 4, column = 3, padx=10, pady=10, sticky=W)
        
        
        #Saavutused
        achiev_label = Label(root, text = "Tähts. saavutused", bg = "white")
        achiev_label.grid(row = 4, column = 4, padx=10, pady=10, sticky=W)
        self.achiev_entry_value = StringVar(root, value = "")
        self.achiev_entry = ttk.Entry(root, textvariable = self.achiev_entry_value)
        self.achiev_entry.grid(row = 4, column = 5, padx=10, pady=10, sticky=W)
        
        # ----- KUUES RIDA -----
        # ----- KUUES RIDA -----
        # ----- KUUES RIDA -----
        # ----- KUUES RIDA -----
        # ----- KUUES RIDA -----   
        
        #Exceli nupp
        self.excel_button = ttk.Button(root,
                                       text = "Lisa exceli fail",
                                       command = lambda: self.stud_from_excel())
        self.excel_button.grid(row = 5, column = 0, columnspan = 2,
                               padx= 10, pady = 10, sticky = W + E)


        #Õpilase lisamise nupp
        self.submit_button = ttk.Button(root,
                            text="Lisa õpilane",
                            command=lambda: self.stud_submit())
        self.submit_button 
        self.submit_button.grid(row=5, column=2, columnspan = 2,
                                padx=10, pady=10, sticky=W + E)
        
        #Õpilase info muutmise nupp
        self.update_button = ttk.Button(root,
                            text="Muuda andmeid",
                            command=lambda: self.update_student())
        self.update_button.grid(row=5, column=4, columnspan = 2,
                                padx=10, pady=10,sticky = W + E)
        
        # ----- SEITSMES RIDA -----
        # ----- SEITSMES RIDA -----
        # ----- SEITSMES RIDA -----
        # ----- SEITSMES RIDA -----
        # ----- SEITSMES RIDA -----
        
        #Õpilaste filtreerimine
        flt_label = ttk.Button(root,
                               text = "Otsi",
                               command = lambda: self.search_student_command())
        flt_label.grid(row = 6, column = 0,
                               padx = 10, pady = 10, sticky = W)
        
        self.flt_entry_value = StringVar(root, value = "")
        self.flt_entry = ttk.Entry(root,
                              textvariable = self.flt_entry_value)
        self.flt_entry.grid(row = 6, column = 1,
                            padx = 10, pady = 10, sticky = W)

# root - raam
root = Tk()
# Tee Andmebaas ekraanile
õpAB = ÕpilasteAB(root)
root.mainloop()

