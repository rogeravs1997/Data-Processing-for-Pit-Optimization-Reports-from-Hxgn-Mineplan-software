from io import open
import re
import pandas as pd
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os 
import tkinter.font as font
def get_folder():
	
	folder_path.set('{}'.format(filedialog.askdirectory(title="Select the folder where your file are.")))


def get_final_route():
	
	destiny_path.set('{}'.format(filedialog.askdirectory(title='Select the folder where you want to save your .xlsx file.')))

def find_TotalValueofPit(path):
	fichero_origen=open(r'{}'.format(path),'r').readlines()
	for i in fichero_origen:
		if (re.search("Total value  of Pit",i))!=None:
			return float(re.search("Total value  of Pit",i).string.split()[-1].replace('.',''))

def find_OreTons(path):
	lista=[]
	fichero_origen=open(r'{}'.format(path),'r').readlines()
	for i in fichero_origen:
		if (re.search("                          Tons",i))!=None:
			lista.append((re.search("                          Tons",i)).string)
	
	return float(lista[-1].split()[1].replace('.',''))

def find_WasteTons(path):
	lista=[]
	fichero_origen=open(r'{}'.format(path),'r').readlines()
	for i in fichero_origen:
		if (re.search("                          Tons",i))!=None:
			lista.append((re.search("                          Tons",i)).string)
	
	return float(lista[-1].split()[2].replace('.',''))

def find_Grade(path):
	lista=[]
	fichero_origen=open(r'{}'.format(path),'r').readlines()
	for i in fichero_origen:
		if (re.search("                          CU",i))!=None:
			lista.append((re.search("                          CU",i)).string)
	
	return (lista[-1].split()[1].replace('.','').replace('0','0.',1))



def generate_DataFrame():
	if folder_path.get()=='' or destiny_path.get()=='' or file_name.get()=='':
		messagebox.showerror('Error','Please fill all the fields above.')   
		
	else: 
		files=[]
		path = folder_path.get()
	
		# r=root, d=directories, f = files
		for r, d, f in os.walk(path):
			for file in f:
				files.append(os.path.join(r, file))
			
		Files=[]
		Ore_Tons=[]
		Waste_Tons=[]
		Total_Tons=[]
		Grades=[]
		Total_Value_Pits=[]
		
		for f in files:
			
			Ore_Tons.append(find_OreTons(f))  
			Waste_Tons.append(find_WasteTons(f))
			Grades.append(find_Grade(f))
			Total_Value_Pits.append(find_TotalValueofPit(f))
			Files.append(f.split('\\')[-1])
		Total_Tons=[sum(x) for x in zip(Ore_Tons, Waste_Tons)]
		
		df=pd.DataFrame(
		{'File': Files,
		 'Ore Tons': Ore_Tons,
		 'Waste Tons': Waste_Tons,
		 'Total Tons':Total_Tons,
		 'Grade': Grades,
		 'Total Value of Pit':Total_Value_Pits
		},index=[x for x in range(1,len(files)+1)])
	
	
	  
		
		df.to_excel ('{}\{}.xlsx'.format(destiny_path.get(),file_name.get()), index = True, header=True)
		
		messagebox.showinfo(title='File created succesfully.',message='Your file was succesfully created at {}'.format(destiny_path.get()))
	

	
def show_info():
	msg=messagebox.showinfo(title="Read me",message='''In order to use correctly this script, please follow the following instructions:
						
-Place all your reports in one single folder w/o any 
 other file in there.

-Be aware that the final name of the file its not 
 repeated in the destination folder.

-Do not add .xlsx extension manually. Just type the name.''')


#Interfaz

root=Tk()


folder_path=StringVar()   
destiny_path=StringVar()
file_name=StringVar()

root.title('Excel Generator from Pit Reports')
root.resizable(0,0)

frm1=Frame(root,width=800,height=600)
frm1.config(bd=3,relief='sunken')
frm1.pack()

espacio0=Label(frm1,text='')
espacio0.grid(row=0, column=0)

title=Label(frm1,text='               Excel Generator from Pit Reports     ',font=('Consolas',20,'bold'))
title.grid(row=1, column=1, columnspan=2)


espacio1=Label(frm1,text='     ')
espacio1.grid(row=2, column=0, columnspan=2)

espacio2=Label(frm1,text='     ')
espacio2.grid(row=3, column=0, columnspan=2)

select_folder=Label(frm1,text="          Select the folder's path where your files are: ",font=('',14))
select_folder.grid(row=3, column=1)

folder_route=Entry(frm1,width=50,textvariable=folder_path)
folder_route.grid(row=3, column=2)

examinate_route=Button(frm1,text='Examinate',command=get_folder)
examinate_route.grid(row=3, column=3)

espacio3=Label(frm1,text='     ')
espacio3.grid(row=3, column=4)

espacio4=Label(frm1,text='     ')
espacio4.grid(row=4, column=0, columnspan=2)

select_destiny=Label(frm1,text="          Select the destination folder: ",font=('',14))
select_destiny.grid(row=5, column=1)

destiny_route=Entry(frm1,width=50,textvariable=destiny_path)
destiny_route.grid(row=5, column=2)

examinate_destiny=Button(frm1,text='Examinate',command=get_final_route)
examinate_destiny.grid(row=5, column=3)

espacio5=Label(frm1,text='')
espacio5.grid(row=6, column=0, columnspan=2)

select_destiny=Label(frm1,text="          Name your file(.xlsx): ",font=('',14))
select_destiny.grid(row=7, column=1)

destiny_route=Entry(frm1,width=50,textvariable=file_name)
destiny_route.grid(row=7, column=2)

espacio7=Label(frm1,text='                        ')
espacio7.grid(row=8, column=4)

espacio6=Label(frm1,text='')
espacio6.grid(row=9, column=0)

Guardar=Button(frm1,text='SAVE',height=2,width=20,command=generate_DataFrame)
Guardar.grid(row=9, column=1,columnspan=4)

Terminar=Button(frm1,text='CLOSE',command=root.destroy,height=2,width=20)
Terminar.grid(row=9, column=2,columnspan=3)

myFont = font.Font(family='Helvetica')

Info=Button(frm1,text='CLICK BEFORE USING',height=2,width=20,bg='red', fg='#ffffff',command=show_info)
Info['font'] = myFont
Info.grid(row=9, column=1)

espacio7=Label(frm1,text='                        ')
espacio7.grid(row=10, column=4)

root.mainloop()
	

	
	
	
	