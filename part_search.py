import tkinter as tk
import tkinter.ttk as ttk
import requests
import pandas as pd
from bs4 import BeautifulSoup as bs
from tkinter import simpledialog
import tkinter.font as tkFont

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()
        self.text_value = ""

    def console_print(self,value):
        self.output_field.insert(tk.INSERT,value)
        self.output_field.see("end")

    def console_clear(self):
        self.output_field.delete('1.0', tk.END)
    def text_clear(self):
        self.text_field.delete('1.0', tk.END)


    def create_widgets(self):
        self.label_font = tkFont.Font(family="Lucida Grande", size=14)

        self.header = tk.Label(self,text="Paste Part Numbers",font=self.label_font)
        self.header.pack()

        self.main_frame = tk.Frame(self,borderwidth=10)
        self.main_frame.pack()

        self.bottom_frame = tk.Frame(self.main_frame)
        self.bottom_frame.pack(side="bottom")
        self.button_frame = tk.Frame(self.main_frame)
        self.button_frame.pack(side="bottom")

        self.left_frame = tk.Frame(self.main_frame)
        self.left_frame.pack(side = "left")
        self.right_frame = tk.Frame(self.main_frame)
        self.right_frame.pack(side = "right")

        self.text_frame = tk.Frame(self.left_frame)
        self.text_frame.pack()
        self.text_field = tk.Text(self.text_frame,borderwidth=3,width=35, height=10,background="white",relief="solid",fg="black")
        self.text_field.pack()

        self.output_field = tk.Text(self.right_frame,borderwidth=3,width=35, height=10,background="white",relief="solid",fg="black")
        self.output_field.insert(tk.END, "Information will show here!\n\nPart numbers go in the left box!")
        self.output_field.pack()


        self.quit = tk.Button(self.button_frame, text="QUIT",command=self.master.destroy)
        self.quit.pack(side="right")
        self.submit = tk.Button(self.button_frame, text="Submit!",command=self.submit)
        self.submit.pack(side="left")

        self.progress_header = tk.Label(self.bottom_frame,text="Progress:",font=self.label_font)
        self.progress_header.pack()
        self.progress = ttk.Progressbar(self.bottom_frame, orient = tk.HORIZONTAL, length = 300, mode = 'determinate')
        self.progress.pack(side="right")


    def submit(self):
        self.text_value=self.text_field.get("1.0","end-1c")
        self.text_value = self.text_value.split("\n")
        #remove blank lines
        while("" in self.text_value) :
            self.text_value.remove("")
        print(self.text_value)
        self.search_for_models(self.text_value)

    @classmethod
    def format_model_number(self,model_numbers_to_format):
        formatted_model_numbers = []

        #cut the garbage away from our model numbers
        for model_number in model_numbers_to_format:
            model_number = model_number.upper()
            model_number = model_number.replace("-","").replace(" ","")
            model_number = model_number.replace("TONERCARTRIDGES,SUPPLIESANDPARTS","")
            model_number = model_number.replace("RICOH","")
            model_number = model_number.replace("ALFICIO","")
            model_number = model_number.replace("AFICIO","")
            model_number = model_number.replace("SP","")
            model_number = model_number.replace("DN","")
            model_number = model_number.replace("EX","")
            model_number = model_number.replace("G","")
            model_number = model_number.replace("F","")
            model_number = model_number.replace("S","")
            model_number = model_number.replace("E1","")
            model_number = model_number.replace("A","")
            model_number = model_number.replace("B","")

            #cleaning data
            if not model_number.count("MP3003") and not  model_number.count("MP4503") and not  model_number.count("MPC80022"):
                formatted_model_numbers.append(model_number)

        #remove duplicates and return
        formatted_model_numbers = sorted(set(formatted_model_numbers))
        return(formatted_model_numbers)

    def search_for_models(self,part_num):
        part_numbers = []
        self.console_clear()
        self.progress["value"] = 0
        for i in range(len(part_num)):
            self.console_print("Searching part:{}/{}...{}".format(i+1,len(part_num),str(part_num[i]).rjust(0,'.'))+"\n")
            self.progress["value"] += 100/len(part_num)
            self.update_idletasks()
            #Create a temporary list to store model numbers
            unformatted_model_numbers = []
            #Process part number into URL then search html with Beautiful Soup
            url = "https://www.precisionroller.com/search.php?q=" + part_num[i]
            request = requests.get(url)
            soup = bs(request.text,"html.parser")

            #Create a list to contain links that we find.
            links_from_url = soup.find_all("a")
            #Search the html for model numbers
            for link in links_from_url:
                if(str(link.get('title')).count('Toner Cartridges,')):
                    if(str(link.get('title')).upper().count('RICOH') or str(link.get('title')).upper().count('ALFICIO')):
                        unformatted_model_numbers.append(link.get('title'))

            #Format our model numbers to be readable
            part_numbers.append(self.format_model_number(unformatted_model_numbers))
        self.export_to_xlsx(part_numbers,self.text_value)


    def export_to_xlsx(self,model_numbers_formatted,part_numbers):
        final_data_frame = pd.DataFrame({'Part Numbers':part_numbers},columns=['Part Numbers','Model Numbers'.ljust(150)])
        for i in range(len(model_numbers_formatted)):
            model_num_str = ""
            for model in model_numbers_formatted[i]:
                model_num_str=model_num_str+model+"/"
            final_data_frame.iat[i,1] = model_num_str

        path_to_exported_xlsx = "Untitled"
        userinput = simpledialog.askstring(title="Complete!",prompt="Please name your file:")
        if userinput:
            path_to_exported_xlsx = userinput

        writer = pd.ExcelWriter(path_to_exported_xlsx+".xlsx", engine='xlsxwriter')
        final_data_frame.to_excel(writer, sheet_name='Model Numbers')
        workbook = writer.book
        worksheet = writer.sheets['Model Numbers']

        #format column width and textwrap
        text_wrap = workbook.add_format({'text_wrap': True})
        worksheet.set_column(2,2, 150,text_wrap)
        worksheet.set_column(1,1,15)

        try:
            writer.save()
        except Exception as err:
            self.console_print("Error saving file! Maybe your filename is invalid?")
        self.console_print("Done! Please close or input more part numbers!")
        self.text_clear()


root = tk.Tk()
root.title("Part Searcher")
root.geometry("600x280+400+200")
app = Application(master=root)
app.mainloop()
