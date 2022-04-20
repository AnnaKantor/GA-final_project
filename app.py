import tkinter as tk
import openpyxl
from PIL import Image, ImageTk
from tkinter.filedialog import askopenfile

root=tk.Tk()
root.title('Translator Dr.Oligo 1')
root.iconbitmap(r'icon.ico')

canvas = tk.Canvas(root, width=600, height=300)
canvas.grid(columnspan=3, rowspan=3)

#logo
logo=Image.open('image.png')
logo=ImageTk.PhotoImage(logo)
logo_label=tk.Label(image=logo)
logo_label.image=logo
logo_label.grid(column=1, row=0)

#instractions
instractions=tk.Label(root, text="Select an Excel file on your computer for translation!", font='Roboto')
instractions.grid(columnspan=3, column=0, row=1)

def open_file():
    browse_text.set('loading...')
    file= askopenfile(parent=root, mode='rb', title='Choose a file', filetype=[('Excel file', '*.xlsx')])
    if file:
        import pandas as pd
        import numpy as np
        import re

        def get_name_excel():
             name=f'{file}'
             name_sliced=name.split("'",1)[1].split("'",1)[0] 
             return(name_sliced)

        df = pd.read_excel(f'{get_name_excel()}', sheet_name="Sense") 
        sense = list(df['Full oligo description: sequence with chemical pattern modifications (Auto generated)'])

        s2=list()   #list of strings - sense
        for i in sense:
            oligo_updated=i.replace("-TegChol", "").replace(")#", ")# ").replace(")(", ") (").replace("P","V ")
            s2.append(oligo_updated)
    
        s3=list()    #list of lists
        for i in s2:
            oligo_updated=re.split(" ", i)
            s3.append(oligo_updated)
 

        df1 = pd.read_excel(f'{get_name_excel()}', sheet_name="Antisense") 
        antisense = list(df1['Full oligo description: sequence with chemical pattern modifications (Auto generated)'])

        as2=list()   #list of strings - antisense
        for i in antisense:
            oligo_updated=i.replace(")#", ")# ").replace(")(", ") (").replace("P","V ").replace("-TegChol", "")
            as2.append(oligo_updated)
    
        as3=list()    #list of lists
        for i in as2:
            oligo_updated=re.split(" ", i)
            as3.append(oligo_updated)

        def translator(a):
            for oligo in a:
                for i in range(len(oligo)):
                    if oligo[i]=="(mA)#":
                         oligo[i]="f"

                    if oligo[i]=="(mC)#":
                         oligo[i]="h"  

                    if oligo[i]=="(mG)#":
                         oligo[i]="i"

                    if oligo[i]=="(mU)#":
                         oligo[i]="j"
                  
                    if oligo[i]=="(fA)#":
                         oligo[i]='k'

                    if oligo[i]=="(fC)#":
                         oligo[i]='l'

                    if oligo[i]=="(fG)#":
                         oligo[i]='m'

                    if oligo[i]=='(fU)#':
                         oligo[i]='n'
    
                    if oligo[i]=="(mA)":
                         oligo[i]='F'
    
                    if oligo[i]=="(mC)":
                         oligo[i]='H'
    
                    if oligo[i]=="(mG)":
                         oligo[i]='I'
    
                    if oligo[i]=="(mU)":
                         oligo[i]='J'
    
                    if oligo[i]=="(fA)":
                         oligo[i]='K'
    
                    if oligo[i]=="(fC)":
                         oligo[i]='L'
    
                    if oligo[i]=="(fG)":
                         oligo[i]='M'
    
                    if oligo[i]=="(fU)":
                         oligo[i]='N'
    
                    if oligo[i]=="P":
                         oligo[i]="V"

            b=list()    #list of translated strings
            for i in a:
                oligo_updated=''.join(i)
                b.append(oligo_updated)     
            return(b)   

        s4=translator(s3)
        sense_first_part=s4[:48]
        sense_second_part=s4[48:96]
        sense_third_part=s4[96:144]
        sense_fourth_part=s4[144:]

        as4=translator(as3)
        antisense_first_part=as4[:48]
        antisense_second_part=as4[48:96]
        antisense_third_part=as4[96:144]
        antisense_fourth_part=as4[144:]

        df_new=df[['Existing Atalanta Oligo ID #']]
        df_new.rename(columns={'Existing Atalanta Oligo ID #':'ID #'}, inplace=True)

        df_s1=df_new[:48]
        df_s1["Sequence"] = sense_first_part
        df_s1.insert(2, "DMT (ON/OFF)", "DMT(OFF)")
        df_s1.index = np.arange(1, len(df_s1)+1)

        df_s2=df_new[48:96]
        df_s2["Sequence"] = sense_second_part
        df_s2.insert(2, "DMT (ON/OFF)", "DMT(OFF)")
        df_s2.index = np.arange(1, len(df_s2)+1)

        df_s3=df_new[96:144]
        df_s3["Sequence"] = sense_third_part
        df_s3.insert(2, "DMT (ON/OFF)", "DMT(OFF)")
        df_s3.index = np.arange(1, len(df_s3)+1)

        df_s4=df_new[144:]
        df_s4["Sequence"] = sense_fourth_part
        df_s4.insert(2, "DMT (ON/OFF)", "DMT(OFF)")
        df_s4.index = np.arange(1, len(df_s4)+1)

        df1_new=df1[['Existing Atalanta Oligo ID #']]
        df1_new.rename(columns={'Existing Atalanta Oligo ID #':'ID #'}, inplace=True)

        df_as1=df1_new[:48]
        df_as1["Sequence"] = antisense_first_part
        df_as1.insert(2, "DMT (ON/OFF)", "DMT(OFF)")
        df_as1.index = np.arange(1, len(df_as1)+1)
        
        df_as2=df1_new[48:96]
        df_as2["Sequence"] = antisense_second_part
        df_as2.insert(2, "DMT (ON/OFF)", "DMT(OFF)")
        df_as2.index = np.arange(1, len(df_as2)+1)

        df_as3=df1_new[96:144]
        df_as3["Sequence"] = antisense_third_part
        df_as3.insert(2, "DMT (ON/OFF)", "DMT(OFF)")
        df_as3.index = np.arange(1, len(df_as3)+1)
        
        df_as4=df1_new[144:]
        df_as4["Sequence"] = antisense_fourth_part
        df_as4.insert(2, "DMT (ON/OFF)", "DMT(OFF)")
        df_as4.index = np.arange(1, len(df_as4)+1)

        df_s1.to_csv('sense-plate-1-oligos-1-48.csv')
        df_s2.to_csv('sense-plate-1-oligos-49-96.csv')
        df_s3.to_csv('sense-plate-2-oligos-1-48.csv')
        df_s4.to_csv('sense-plate-2-oligos-49-96.csv')

        df_as1.to_csv('antisense-plate-1-oligos-1-48.csv')
        df_as2.to_csv('antisense-plate-1-oligos-49-96.csv')
        df_as3.to_csv('antisense-plate-2-oligos-1-48.csv')
        df_as4.to_csv('antisense-plate-2-oligos-49-96.csv')

        browse_text.set("Browse")

#browse button
browse_text=tk.StringVar()
browse_btn=tk.Button(root, textvariable=browse_text, command=lambda:open_file(), font="Roboto", bg='#4982eb', fg='#fff', height=2, width=15)
browse_text.set('Browse')
browse_btn.grid(column=1, row=2)

canvas = tk.Canvas(root, width=600, height=50)
canvas.grid(columnspan=3)

root.mainloop()

