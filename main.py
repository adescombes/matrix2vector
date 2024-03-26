import os
import sys
from functions import config_1, config_2
import pandas as pd
import tkinter as tk
from tkinter.filedialog import *
import tkinter.scrolledtext as st 
from PIL import Image, ImageTk
import numpy as np

main = tk.Tk()
main.title("Matrix -> vector")
main.geometry("950x500")
dict_params = {}

# zone for scrollable text : displays user inputs
text_area = st.ScrolledText(main, 
                                width = 120,  
                                height = 4,  
                                font = ("Helvetica", 
                                        10))     
text_area.place(x = 10, y = 10) 
text_area.insert(tk.INSERT, "") 


def make_thumbnail(filename, max_width, max_height):
    width, height = Image.open(filename).size
    ratio = min( max_width / width, max_height / height )
    new_width = int( np.round( ratio * width ) )
    new_height = int( np.round( ratio * height ) )
    thumb = Image.open(filename).resize((new_width, new_height), Image.LANCZOS)
    return thumb


def input_button():
    folder = tk.filedialog.askopenfilename(filetypes=[('xlsx','.xlsx')])
    input_path.set(folder)
    dict_params['input'] = folder
    text_area.insert(tk.INSERT, "INPUT : %s" % folder)

input_path = tk.StringVar()
input_button = tk.Button(main, text = "Input", command = input_button)
input_button.place(x = 26, y = 100)

def output_button():
    folder = tk.filedialog.asksaveasfilename(filetypes=[('xlsx','.xlsx')])
    output_path.set(folder)
    if '.xlsx' not in folder:
        folder += '.xlsx'
    dict_params['output'] = folder
    text_area.insert(tk.INSERT, "\nSAVE TO : %s" % folder)

output_path = tk.StringVar()
output_button = tk.Button(main, text = "Save file to...", command = output_button)
output_button.place(x = 26, y = 140)

filename_img_c1 = './config_1.jpg'
box_c1 = tk.Label(main)
box_c1.place(x = 26, y = 180)
thumb_c1 = make_thumbnail(filename_img_c1, 400, 400)
img_1 = ImageTk.PhotoImage(thumb_c1)
box_c1.config(image=img_1)
box_c1.image = img_1

filename_img_c2 = './config_2.jpg'
box_c2 = tk.Label(main)
box_c2.place(x = 500, y = 180)
thumb_c2 = make_thumbnail(filename_img_c2, 400, 400)
img_2 = ImageTk.PhotoImage(thumb_c2)
box_c2.config(image=img_2)
box_c2.image = img_2



def run_config_1():

    writer = pd.ExcelWriter(dict_params.get('output'), engine = 'xlsxwriter')
    excel_file = pd.ExcelFile(dict_params.get('input'))
    sheet_names_list = excel_file.sheet_names

    for s_n in sheet_names_list:
        print("Reading file, sheet %s" % s_n)

        try:   
            df_vector = config_1(excel_file, s_n, writer)
            print('try')
            df_vector.to_excel(writer, index = False, sheet_name = s_n)

        except:
            print('except')

    writer.close()
    print('Vector saved : %s' % dict_params.get('output'))
    main.destroy()


def run_config_2():

    writer = pd.ExcelWriter(dict_params.get('output'), engine = 'xlsxwriter')
    excel_file = pd.ExcelFile(dict_params.get('input'))
    sheet_names_list = excel_file.sheet_names

    for s_n in sheet_names_list:
        print("Reading file, sheet %s" % s_n)

        try:   
            df_vector = config_2(excel_file, s_n, writer)
            print('try')
            

        except Exception as e:
            print('except')
            print(type(e))
            print(e.args)
            print(e)

    writer.close()
    print('Vector saved : %s' % dict_params.get('output'))
    main.destroy()

config_1_button = tk.Button(main, text="disposition 1", command=run_config_1)
config_1_button.place( x = 210, y = 380)

config_2_button = tk.Button(main, text="disposition 2", command=run_config_2)
config_2_button.place( x = 700, y = 380)

main.mainloop()
