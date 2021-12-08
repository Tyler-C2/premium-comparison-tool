import tkinter as tk
from tkinter import filedialog
from tkcalendar import Calendar, DateEntry
import xl_format as xl_f
import os

class Window(tk.Frame):
    def __init__(self, window):
        self.this_files_path = os.path.dirname(os.path.abspath(__file__))
        self.new_wb = xl_f.WorkbookCreator()
        self.window=window
        self.window.title("Excel program")
        self.window.geometry("550x450")
        self.create_frames()
        self.create_string_vars()
        self.create_widgets()
        
    def create_menu(self):
        self.menu = tk.Menu(self.window)
        self.window.config(menu=self.menu)

        fileMenu = tk.Menu(self.menu, tearoff=0)
        fileMenu.add_command(label="Select File", command=self.select_path)
        self.menu.add_cascade(label="File", menu=fileMenu)

    def select_path(self):
        self.window.filePath = filedialog.askopenfilename(initialdir=self.this_files_path, title="File Select", filetypes=[("Excel Files", "*.xlsx")]) 
        self.new_wb.add_path(self.window.filePath)
        if self.new_wb.path != None:
            self.feedback["text"] = f"You are now working with {self.new_wb.file_name}"
        else:
            self.feedback["text"] = "No file selected."

    def create_string_vars(self):
        self.c1_items = {}
        self.c2_items = {}

        self.car1_var = tk.StringVar()
        self.car2_var = tk.StringVar()

        for i in range(1,9):
            self.c1_items[f'c1_item{i}'] = tk.StringVar()
            self.c2_items[f'c2_item{i}'] = tk.StringVar()  
    
    def create_frames(self):
        self.top_frame = tk.Frame(self.window, pady=3)
        self.top_frame.grid(row=0)

        self.mid_frame = tk.Frame(self.window, width=555, height=310, pady=3)
        self.mid_frame.grid(row=1)

        self.bot_frame = tk.Frame(self.window, width=550, height=60, pady=3)
        self.bot_frame.grid(row=2)

    def create_widgets(self):
        
        # create top menu
        self.create_menu()

        # information and feedback labels
        self.top_inform1 = tk.Label(self.top_frame, text="Add information to the form below", font=25).grid(row=1)
        self.top_inform2 = tk.Label(self.top_frame, text="The Submit button holds entered data", font=25) .grid(row=2)
        self.top_inform3 = tk.Label(self.top_frame, text="The Finish button adds the held data to the excel sheet", font=25).grid(row=3)

        self.feedback = tk.Label(self.bot_frame, text="", font=25)
        self.feedback.grid()

        #start/end date 
        self.start = DateEntry(self.mid_frame, width=20)
        self.end = DateEntry(self.mid_frame, width=20)
        self.start.grid(row = 1, column = 2)
        self.end.grid(row = 1, column = 4)

        # top labels 
        self.top_label = tk.Label(self.mid_frame, text="Car 1").grid(row = 2, column = 2)
        self.top_label = tk.Label(self.mid_frame, text="Car 2").grid(row = 2, column = 4)
        
        # pads
        self.pad_label1 = tk.Label(self.mid_frame, text="  ").grid(row=0,column=1)
        self.pad_label2 = tk.Label(self.mid_frame, text="  ").grid(row=0,column=3)
        self.pad_label3 = tk.Label(self.mid_frame, text="  ").grid(row=12, column=3)
        self.pad_label4 = tk.Label(self.mid_frame, text="  ").grid(row=0, column=3)
        self.pad_label5 = tk.Label(self.mid_frame, text="  ").grid(row=14, column=3)

        # item labels
        self.data_label = tk.Label(self.mid_frame, text="Start/End Date : ").grid(row = 1)
        self.car_label = tk.Label(self.mid_frame, text="Car Name : ").grid(row = 3)
        self.item1_label = tk.Label(self.mid_frame, text=f"{self.new_wb.items[0]} : ").grid(row = 4)     
        self.item2_label = tk.Label(self.mid_frame, text=f"{self.new_wb.items[1]} : ").grid(row = 5)     
        self.item3_label = tk.Label(self.mid_frame, text=f"{self.new_wb.items[2]} : ").grid(row = 6)
        self.item4_label = tk.Label(self.mid_frame, text=f"{self.new_wb.items[3]} : ").grid(row = 7)
        self.item5_label = tk.Label(self.mid_frame, text=f"{self.new_wb.items[4]} : ").grid(row = 8)
        self.item6_label = tk.Label(self.mid_frame, text=f"{self.new_wb.items[5]} : ").grid(row = 9)
        self.item7_label = tk.Label(self.mid_frame, text=f"{self.new_wb.items[6]} : ").grid(row = 10)  
        self.item8_label = tk.Label(self.mid_frame, text=f"{self.new_wb.items[7]} : ").grid(row = 11)
        
        # first column of entries
        self.car1_name = tk.Entry(self.mid_frame, textvariable=self.car1_var).grid(row=3,column=2)
        self.car1_item1 = tk.Entry(self.mid_frame, textvariable=self.c1_items['c1_item1']).grid(row=4,column=2)
        self.car1_item2 = tk.Entry(self.mid_frame, textvariable=self.c1_items['c1_item2']).grid(row=5,column=2) 
        self.car1_item3 = tk.Entry(self.mid_frame, textvariable=self.c1_items['c1_item3']).grid(row=6,column=2)
        self.car1_item4 = tk.Entry(self.mid_frame, textvariable=self.c1_items['c1_item4']).grid(row=7,column=2)
        self.car1_item5 = tk.Entry(self.mid_frame, textvariable=self.c1_items['c1_item5']).grid(row=8,column=2)
        self.car1_item6 = tk.Entry(self.mid_frame, textvariable=self.c1_items['c1_item6']).grid(row=9,column=2) 
        self.car1_item7 = tk.Entry(self.mid_frame, textvariable=self.c1_items['c1_item7']).grid(row=10,column=2)
        self.car1_item8 = tk.Entry(self.mid_frame, textvariable=self.c1_items['c1_item8']).grid(row=11,column=2)
        
        #second column of entries
        self.car2_name = tk.Entry(self.mid_frame, textvariable=self.car2_var).grid(row=3,column=4)
        self.car2_item1 = tk.Entry(self.mid_frame, textvariable=self.c2_items['c2_item1']).grid(row=4,column=4)
        self.car2_item2 = tk.Entry(self.mid_frame, textvariable=self.c2_items['c2_item2']).grid(row=5,column=4)
        self.car2_item3 = tk.Entry(self.mid_frame, textvariable=self.c2_items['c2_item3']).grid(row=6,column=4)
        self.car2_item4 = tk.Entry(self.mid_frame, textvariable=self.c2_items['c2_item4']).grid(row=7,column=4)
        self.car2_item5 = tk.Entry(self.mid_frame, textvariable=self.c2_items['c2_item5']).grid(row=8,column=4)
        self.car2_item6 = tk.Entry(self.mid_frame, textvariable=self.c2_items['c2_item6']).grid(row=9,column=4)
        self.car2_item7 = tk.Entry(self.mid_frame, textvariable=self.c2_items['c2_item7']).grid(row=10,column=4)
        self.car2_item8 = tk.Entry(self.mid_frame, textvariable=self.c2_items['c2_item8']).grid(row=11,column=4)

        #submit button
        self.button = tk.Button(self.mid_frame, text= "submit", command=self.process_entry)
        self.button.grid(ipadx=25,row=13, column=2)

        #finish button
        self.fin_button = tk.Button(self.mid_frame, text="Finish", command=self.finshed)
        self.fin_button.grid(ipadx=25,row=13, column=4)

    def process_entry(self):
        #date processing
        start_date = self.start.get_date()
        end_date = self.end.get_date()
        dates = xl_f.Parser.format_date(start_date, end_date)
        
        # car processing
        car1 = self.car1_var.get()
        car2 = self.car2_var.get()
        cars =[car1,car2] 
        self.car1_var.set(car1)
        self.car2_var.set(car2)

        # line item processing
        formatted_items = xl_f.Parser.parse_line_items(self.c1_items,self.c2_items)
        self.new_wb.create_period(dates,cars,formatted_items[0],formatted_items[1])

        self.feedback["text"] = "Entry accepted. Press Finish to add to Excel file "    

    def finshed(self):
        while self.new_wb.path == None:
            self.select_path()
            
        self.new_wb.finished_input(self.new_wb.path)

        self.feedback["text"] = "Data added to selected Excel file."

if __name__ == "__main__":
    root = tk.Tk()
    Window(root)
    root.mainloop()