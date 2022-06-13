import tkinter as tk
from tkinter import filedialog
from tkcalendar import Calendar, DateEntry
import xl_format as xl_f
import image_input
import os

class Window(tk.Frame):
    def __init__(self, window):
        self.this_files_path = os.path.dirname(os.path.abspath(__file__))
        self.new_wb = xl_f.WorkbookCreator()
        self.new_img_read_car_name = image_input.Top_ROI()
        self.new_img_read_vals = image_input.Right_ROI()
        self.selected_car = None
        self.window = window
        self.window.title("Excel program")
        self.window.geometry("550x450")
        self.create_frames()
        self.create_string_vars()
        self.create_widgets()
        
    def create_menu(self):
        self.menu = tk.Menu(self.window)
        self.window.config(menu=self.menu)

        fileMenu = tk.Menu(self.menu, tearoff=0)
        fileMenu.add_command(label="Select File", command=self.select_path_to_xl)
        autoFillMenu = tk.Menu(self.menu, tearoff=0, )
        autoFillMenu.add_command(label="Select Image", command=self.select_path_to_image)
        self.menu.add_cascade(label="File", menu=fileMenu)
        self.menu.add_cascade(label="Auto Fill", menu=autoFillMenu)

    def select_path_to_xl(self):
        file_type = [("Excel Files", "*.xlsx")]
        self.window.filePath = filedialog.askopenfilename(initialdir=self.this_files_path, title="File Select", filetypes=file_type) 
        self.new_wb.add_path(self.window.filePath)
        if self.new_wb.path != None:
            self.feedback["text"] = f"You are now working with {self.new_wb.file_name}"
        else:
            self.feedback["text"] = "No file selected."

    def select_path_to_image(self):
        file_types = [("PNG Files", "*.png"),("Jpeg Files", "*.jpeg"),("Jpg Files", "*.jpg")]
        self.window.filePath = filedialog.askopenfilename(initialdir=self.this_files_path, title="File Select", filetypes=file_types)
        self.new_img_read_car_name.add_path(self.window.filePath)
        self.new_img_read_vals.add_path(self.window.filePath)
        if self.new_img_read_vals.path != None:
            try:
                self.new_img_read_car_name.creator()
                self.new_img_read_vals.creator()
                self.car_select_popup()
                self.window.wait_window(pop)
                
                if self.selected_car == 1:
                    self.car1_name.delete(0,tk.END)
                    self.car1_name.insert(0,self.new_img_read_car_name.vehicle_text)
                    
                    for i in range(len(self.car1_entry_fields)):
                        self.car1_entry_fields[i].delete(0,tk.END)
                        self.car1_entry_fields[i].insert(0,self.new_img_read_vals.premium_values[i])

                    self.feedback["text"] = f"Car 1 has been auto filled from provided document."
                
                elif self.selected_car == 2:
                    self.car2_name.delete(0,tk.END)
                    self.car2_name.insert(0,self.new_img_read_car_name.vehicle_text)
                    
                    for i in range(len(self.car2_entry_fields)):
                        self.car2_entry_fields[i].delete(0,tk.END)
                        self.car2_entry_fields[i].insert(0,self.new_img_read_vals.premium_values[i])

                    self.feedback["text"] = f"Car 2 has been auto filled from provided document."

                elif self.selected_car == None :

                    self.feedback["text"] = f"auto fill was not performed."
            except:
                self.feedback["text"] = f"file not recognized. Please check the document."
        else:
            self.feedback["text"] = f"auto fill was not performed."
            
        self.selected_car = None
        self.new_img_read_car_name.clear()
        self.new_img_read_vals.clear()

    def car_select_popup(self):
        global pop 
        pop = tk.Toplevel(self.window)
        pop.geometry("300x100")
        pop.title("Car Select")
        pop_label = tk.Label(pop, text="Select which car to auto fill.")
        pop_label.pack(pady=10)
        pop_frame = tk.Frame(pop)
        pop_frame.pack(pady=5)
        button1 = tk.Button(pop_frame, text= "Car 1", command=lambda: self.car_choice(1))
        button1.grid(row=2,column=1,padx=10)
        button2 = tk.Button(pop_frame, text= "Car 2", command=lambda: self.car_choice(2))
        button2.grid(row=2,column=2,padx=10)

    def car_choice(self,choice):
        pop.destroy()
        if choice == 1:
            self.selected_car = 1
        elif choice == 2:
            self.selected_car = 2
        else:
            self.selected_car = None

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
        self.create_item_labels()

        # first column of entries
        self.car1_entry_fields = []

        self.car1_name = tk.Entry(self.mid_frame, textvariable=self.car1_var)
        self.car1_name.grid(row=3,column=2)

        self.create_entry_field(self.car1_entry_fields,self.c1_items,"c1_item",2)
        
        #second column of entries
        self.car2_entry_fields = []

        self.car2_name = tk.Entry(self.mid_frame, textvariable=self.car2_var)
        self.car2_name.grid(row=3,column=4)

        self.create_entry_field(self.car2_entry_fields,self.c2_items,"c2_item",4)

        #submit button
        self.button = tk.Button(self.mid_frame, text= "submit", command=self.process_entry)
        self.button.grid(ipadx=25,row=13, column=2)

        #finish button
        self.fin_button = tk.Button(self.mid_frame, text="Finish", command=self.finshed)
        self.fin_button.grid(ipadx=25,row=13, column=4)

    def create_item_labels(self):
        self.item_labels = []
        row = 4
        for i in range(8):
            self.item_labels.append(tk.Label(self.mid_frame, text=f"{self.new_wb.items[i]} : ").grid(row = row))
            row += 1

    def create_entry_field(self,field_lst,text_var,item_str,column):
        row = 4
        for i in range(8):
            field_lst.append(tk.Entry(self.mid_frame, textvariable=text_var[f"{item_str}{i+1}"]))
            field_lst[i].grid(row=row,column=column)
            row += 1

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
            self.select_path_to_xl()
            
        self.new_wb.finished_input(self.new_wb.path)

        self.feedback["text"] = "Data added to selected Excel file."

if __name__ == "__main__":
    root = tk.Tk()
    Window(root)
    root.mainloop()
