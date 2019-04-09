import openpyxl
import sqlite3

from tkinter import *
from tkinter import ttk
from tkinter import filedialog

from PIL import Image, ImageTk


class LogIn:
    
    def __init__(self, root):
        root.title("JWG Andmebaas")
        root.geometry("240x150")
        root.resizable(width = False, height = False)
        root["bg"] = "white" #Background = white

        #Username
        label1 = Label(root, text = "Kasutajanimi", bg = "white")
        label1.grid(row = 0, column = 0, padx = 10, pady = 10, sticky = W)

        self.label1_entry_value = StringVar(root, value = "")
        self.label1_entry = ttk.Entry(root,
                                      textvariable = self.label1_entry_value)
        self.label1_entry.grid(row = 0, column = 1, padx = 10, pady = 10, sticky = W)

        #Password
        label2 = Label(root, text = "Parool", bg = "white")
        label2.grid(row = 1, column = 0, padx = 10, pady = 10, sticky = W)

        self.label2_entry_value = StringVar(root, value = "")
        self.label2_entry = ttk.Entry(root, show = "*",
                                      textvariable = self.label2_entry_value)
        self.label2_entry.grid(row = 1, column = 1, padx = 10, pady = 10, sticky = W)



        self.login_button = ttk.Button(root,
                            text="Logi sisse",
                            command=lambda: self.login())
        self.login_button.grid(row=2, column=0, columnspan = 2,
                                padx=10, pady=10,sticky = W + E)


    def login(self):
        if self.label1_entry_value.get() == "admin" and self.label2_entry_value.get() == "123":
            self.newWindow = Toplevel(root)
            self.õpAB = ÕpilasteAB(self.newWindow)
        

class ÕpilasteAB :
    
    #db_conn to represent the connection with the database
    db_conn = 0
    
    #c - cursor to perform SQL functions
    c = 0
    
    #It's neccessary to index students
    curr_student = 0
    
    def setup_db(self):
        
        #Connect to the given database, otherwise create a new database
        self.db_conn = sqlite3.connect('UT_Database.db')
        self.c = self.db_conn.cursor()
 
        # Try for errors
        try:
            self.db_conn.execute("CREATE TABLE if not exists Õpilased(ID INTEGER PRIMARY KEY AUTOINCREMENT, FName TEXT, "
                                 "LName TEXT, Class INTEGER, Sports TEXT, Sex TEXT, Length INTEGER, "
                                 "Club TEXT, Freq INTEGER, Trainer TEXT, Achiev TEXT);")
 
            self.db_conn.commit()
 
        except sqlite3.OperationalError:
            print("ERROR : 1 : Tabelit ei olnud võimalik teha")
        
 
    def stud_submit(self):
        #Add a student to the database
        self.db_conn.execute("INSERT INTO Õpilased (FName, LName, Class, Sports, Sex, Length, Club, Freq, Trainer, Achiev) " +
                             "VALUES ('" +
                             self.fn_entry_value.get() + "', '" +
                             self.ln_entry_value.get() + "', '" +
                             self.class_entry_value.get() + "', '" +
                             self.sports_entry_value.get() + "', '" +
                             self.sex_entry_value.get() + "', '" +
                             self.length_entry_value.get() + "', '" +
                             self.club_entry_value.get() + "', '" +
                             self.freq_entry_value.get() + "', '" +
                             self.trainer_entry_value.get() + "', '" +
                             self.achiev_entry_value.get() + "')")

        
        #Clear the entry boxes
        self.fn_entry.delete(0, "end")
        self.ln_entry.delete(0, "end")
        self.class_entry.delete(0, "end")
        self.sports_entry.delete(0, "end")
        self.sex_entry.delete(0, "end")
        self.length_entry.delete(0, "end")
        self.club_entry.delete(0, "end")
        self.freq_entry.delete(0, "end")
        self.trainer_entry.delete(0, "end")
        self.achiev_entry.delete(0, "end")
        
        #It's neccessary to commit the changes and update the listbox to show the changes
        self.db_conn.commit()
        self.update_listbox()

         
    def update_listbox(self):
        #Clear listbox to show changes made after
        self.list_box.delete(0, END)
 
        #For loop every student in the database
        try:
            result = self.c.execute("SELECT ID, FName, LName FROM Õpilased")
 
            for row in result:
                stud_id = row[0]
                stud_fname = row[1]
                stud_lname = row[2]
                
                #Add all the students into the database
                #We had to clear the listbox at the start of the fuction to avoid creating duplicated students into listbox
                self.list_box.insert(stud_id,
                                     stud_fname + " " +
                                     stud_lname)
 
        except sqlite3.OperationalError:
            print("ERROR : 2.1 : Given database doesn't exist")
 
        except:
            print("ERROR : 2.2 : Couldn't recieve data from DB ")
    
    def load_student(self, event = None):
        
        #Get the id
        lb_widget = event.widget
        index = str(lb_widget.curselection()[0] + 1)

        self.curr_student = index

        result = self.c.execute("SELECT ID, FName, LName, Class, Sports, Sex, Length, Club, Freq, Trainer, Achiev FROM Õpilased WHERE ID=" + index)
        for row in result:
            stud_id = row[0]
            stud_fname = row[1]
            stud_lname = row[2]
            stud_class = row[3]
            stud_sports = row[4]
            stud_sex = row[5]
            stud_length = row[6]
            stud_club = row[7]
            stud_freq = row[8]
            stud_trainer = row[9]
            stud_achiev = row[10]
            
            #Change all the entry values with data from the chosen student
            self.fn_entry_value.set(stud_fname)
            self.ln_entry_value.set(stud_lname)
            self.class_entry_value.set(stud_class)
            self.sports_entry_value.set(stud_sports)
            self.sex_entry_value.set(stud_sex)
            self.length_entry_value.set(stud_length)
            self.club_entry_value.set(stud_club)
            self.freq_entry_value.set(stud_freq)
            self.trainer_entry_value.set(stud_trainer)
            self.achiev_entry_value.set(stud_achiev)
 
    
    def update_student(self, event=None):
        #Update student's data
        try:
            self.db_conn.execute("UPDATE Õpilased SET FName='" +
                                self.fn_entry_value.get() +
                                "', LName='" +
                                self.ln_entry_value.get() +
                                "', Class='" +
                                self.class_entry_value.get() +
                                "', Sports='" +
                                self.sports_entry_value.get() +
                                "', Sex='" +
                                self.sex_entry_value.get() +
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
            print("ERROR : 3 : Couldn't update the database")
 
        self.fn_entry.delete(0, "end")
        self.ln_entry.delete(0, "end")
        self.class_entry.delete(0, "end")
        self.sports_entry.delete(0, "end")
        self.sex_entry.delete(0, "end")
        self.length_entry.delete(0, "end")
        self.club_entry.delete(0, "end")
        self.freq_entry.delete(0, "end")
        self.trainer.delete(0, "end")
        self.achiev.delete(0, "end")
        
        self.update_listbox()
        
    def stud_from_excel(self):
        #Creating the file dialog
        fileName = filedialog.askopenfilename(filetypes = (("Excel file", "*.xlsx"), ("All files", "*.*")))

        #Opening the workbook
        wb = openpyxl.load_workbook(fileName, data_only = True)
        wb_sheet = wb.active

        data_list = []
        empty_string = ""
            
        #Because the Excel file cells are from J10 to J19,
        #I made a for loop to check for empty cells
        for m in range(0, 10):
            if wb_sheet["J1"+str(m)].value:
                data_list.append(wb_sheet["J1"+str(m)].value)
            else:
                data_list.append(empty_string)
                
        #Like before, add the data from the list to the DB
        self.db_conn.execute("INSERT INTO Õpilased (FName, LName, Class, Sports, Sex, Length, Club, Freq, Trainer, Achiev) "
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


        self.db_conn.commit()
        self.update_listbox()        
        
    def search_student_command(self):
        #I am refreshing the listbox to clear all the filters (clear the colour from filtered students)
        ÕpilasteAB.update_listbox(self)

        #All the students that contain searched data are coloured
        #s_t - searched student
        s_t = self.flt_entry_value.get()
        self.c.execute("SELECT ID FROM Õpilased WHERE FName = ? OR LName = ? OR Sports = ? OR Class = ? OR Sex = ? OR Length = ? OR Club = ? OR Freq = ? OR Trainer = ? OR Achiev = ?",
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
        root["bg"] = "white" #Background = white

        #Taustapilt
        bg_image = PhotoImage(file= "ut_taust4_fixed.png")
        x = Label(root, image = bg_image)
        x.photo=bg_image
        x.place(x = 0, y = 0)
        
        # ----- FIRST ROW -----
        # ----- FIRST ROW -----
        # ----- FIRST ROW -----
        # ----- FIRST ROW -----
        # ----- FIRST ROW -----
 
        self.list_box = Listbox(root)
        #.bind allows to call a function when Listbox is pressed with mouse click
        self.list_box.bind('<<ListboxSelect>>', self.load_student)
        self.list_box.insert(1, "Õpilased...")
        self.list_box.grid(row=0, column=0, columnspan = 2, rowspan = 5, padx=10, pady=10, sticky = W+E+N+S)
 
        #Loading/Creating the database
        #Updating the listbox
        self.setup_db()
        self.update_listbox()
        
        #First name
        fn_label = Label(root, text="Eesnimi", bg = "white")
        fn_label.grid(row = 0, column = 2, padx = 10, pady = 10, sticky = W)

        self.fn_entry_value = StringVar(root, value = "")
        self.fn_entry = ttk.Entry(root,
                                  textvariable = self.fn_entry_value)
        self.fn_entry.grid(row = 0, column = 3, padx = 10, pady = 10, sticky = W)
        
        #Last name
        ln_label = Label(root, text="Perekonnanimi", bg = "white")
        ln_label.grid(row=0, column=4, padx=10, pady=10, sticky=W)
        
        self.ln_entry_value = StringVar(root, value="")
        self.ln_entry = ttk.Entry(root,
                                  textvariable=self.ln_entry_value)
        self.ln_entry.grid(row=0, column=5, padx=10, pady=10, sticky=W)
        
        # ----- SECOND ROW -----
        #------ SECOND ROW -----
        #------ SECOND ROW -----
        # ----- SECOND ROW -----
        #------ SECOND ROW -----
        
        #Sex
        sex_label = Label(root, text = "Sugu", bg = "white")
        sex_label.grid(row = 1, column = 2, padx = 10, pady= 10, sticky = W)
        
        self.sex_entry_value = StringVar(root, value = "")
        self.sex_entry = ttk.Entry(root,
                                    textvariable = self.sex_entry_value)
        self.sex_entry.grid(row = 1, column = 3, padx = 10, pady= 10, sticky = E)

        #Student's class
        class_label = Label(root, text = "Klass", bg = "white")
        class_label.grid(row = 1, column = 4, padx = 10, pady = 10, sticky = W)
        
        self.class_entry_value = StringVar(root, value = "")
        self.class_entry = ttk.Entry(root,
                                     textvariable = self.class_entry_value)
        self.class_entry.grid(row = 1, column = 5, padx = 10, pady = 10, sticky = E)
        
        # ------THIRD ROW-----
        # ------THIRD ROW-----
        # ------THIRD ROW-----
        # ------THIRD ROW-----
        # ------THIRD ROW-----
        
        #Sports
        sports_label = Label(root, text = "Spordiala", bg = "white")
        sports_label.grid(row = 2, column = 2, padx= 10, pady= 10, sticky = W)
        
        self.sports_entry_value = StringVar(root, value = "")
        self.sports_entry = ttk.Entry(root,
                                         textvariable= self.sports_entry_value)
        self.sports_entry.grid(row = 2, column = 3, padx = 10, pady = 10, sticky = W)
        
        #Whether he has a club or not
        club_label = Label(root, text = "Klubi", bg = "white")
        club_label.grid(row = 2, column = 4, padx=10, pady=10, sticky=W)
        
        self.club_entry_value = StringVar(root, value = "")
        self.club_entry = ttk.Entry(root, textvariable = self.club_entry_value)
        self.club_entry.grid(row = 2, column = 5, padx=10, pady=10, sticky=W)
        
        # ----- FOURTH ROW -----
        # ----- FOURTH ROW -----
        # ----- FOURTH ROW -----
        # ----- FOURTH ROW -----
        # ----- FOURTH ROW -----
        
        #Frequency of trainings
        freq_label = Label(root, text="Treeningute arv", bg= "white")
        freq_label.grid(row = 3, column = 2, padx=10, pady=10, sticky=W)
        
        self.freq_entry_value = StringVar(root, value = "")
        self.freq_entry = ttk.Entry(root, textvariable = self.freq_entry_value)
        self.freq_entry.grid(row = 3, column = 3, padx=10, pady=10, sticky=W)
        
        #Trainer's existence
        trainer_label = Label(root, text = "Treener", bg = "white")
        trainer_label.grid(row = 3, column = 4, padx=10, pady=10, sticky=W)
        
        self.trainer_entry_value = StringVar(root, value = "")
        self.trainer_entry = ttk.Entry(root,
                                       textvariable = self.trainer_entry_value)
        self.trainer_entry.grid(row = 3, column = 5, padx=10, pady=10, sticky=W)
        
        # ----- FIFTH ROW -----
        # ----- FIFTH ROW -----
        # ----- FIFTH ROW -----
        # ----- FIFTH ROW -----
        # ----- FIFTH ROW -----
        
        #How many years has the student trained for
        length_label = Label(root, text = "Kaua tegelenud", bg = "white")
        length_label.grid(row = 4, column = 2, padx=10, pady=10, sticky=W)
        
        self.length_entry_value = StringVar(root, value = "")
        self.length_entry = ttk.Entry(root,
                                      textvariable = self.length_entry_value)
        self.length_entry.grid(row = 4, column = 3, padx=10, pady=10, sticky=W)
        
        
        #Achievements
        achiev_label = Label(root, text = "Tähts. saavutused", bg = "white")
        achiev_label.grid(row = 4, column = 4, padx=10, pady=10, sticky=W)
        self.achiev_entry_value = StringVar(root, value = "")
        self.achiev_entry = ttk.Entry(root, textvariable = self.achiev_entry_value)
        self.achiev_entry.grid(row = 4, column = 5, padx=10, pady=10, sticky=W)
        
        # ----- SIXTH ROW -----
        # ----- SIXTH ROW -----
        # ----- SIXTH ROW -----
        # ----- SIXTH ROW -----
        # ----- SIXTH ROW -----   
        
        #The button for the Excel file
        self.excel_button = ttk.Button(root,
                                       text = "Lisa exceli fail",
                                       command = lambda: self.stud_from_excel())
        self.excel_button.grid(row = 5, column = 0, columnspan = 2,
                               padx= 10, pady = 10, sticky = W + E)


        #Button for adding students into DB
        self.submit_button = ttk.Button(root,
                            text="Lisa õpilane",
                            command=lambda: self.stud_submit())
        self.submit_button 
        self.submit_button.grid(row=5, column=2, columnspan = 2,
                                padx=10, pady=10, sticky=W + E)
        
        #Button for changing student's data
        self.update_button = ttk.Button(root,
                            text="Muuda andmeid",
                            command=lambda: self.update_student())
        self.update_button.grid(row=5, column=4, columnspan = 2,
                                padx=10, pady=10,sticky = W + E)
        
        # ----- SEVENTH ROW -----
        # ----- SEVENTH ROW -----
        # ----- SEVENTH ROW -----
        # ----- SEVENTH ROW -----
        # ----- SEVENTH ROW -----
        
        #Button for filtering students
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


root = Tk()
õpAB = LogIn(root)
root.mainloop()
