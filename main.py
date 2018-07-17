from tkinter import filedialog
from tkinter import *
import tkinter
import pandas as pd
from pandas import *

dfResult = pd.DataFrame(columns=['Index1', 'Name1', 'Index2', 'Name2', 'Common', 'Similarity'])
writer = pd.ExcelWriter('result.xlsx',engine= 'xlsxwriter')

def findSim(stringA,stringB,intA,intB):
    global dfResult

    try:
        lowerA = stringA.lower()
        lowerB = stringB.lower()
        listA = set(lowerA.split())
        listB = set(lowerB.split())
    except:
        #print('Float detected: ', stringA, '&&', stringB, ' \nAt', intA, ' and ', intB)
        return
    notinclude = {'and','is','of','for','set'}
    common = set.intersection(listA,listB)
    common.difference(notinclude)
    total = set.union(listA, listB)
    total.difference(notinclude)
    if len(common) > 3:
        data = {'Index1': intA, 'Name1': stringA, 'Index2': intB, 'Name2': stringB, 'Common': len(common),
                'Similarity': len(common) * 100 / len(total)}
        dfResult = dfResult.append(data,ignore_index = True)
        #print('Line number : ',intA,' and ',intB,'\nThat is :-',stringA , ' &&& ', stringB)

'''
def proces():
    number1=Entry.get(E1)
    number2=Entry.get(E2)
    operator=Entry.get(E3)  
    number1=int(number1)
    number2=int(number2)
    if operator =="+":
        answer=number1+number2
    if operator =="-":
        answer=number1-number2
    if operator=="*":
        answer=number1*number2
    if operator=="/":
        answer=number1/number2
    Entry.insert(E4,0,answer)
    print(answer)

top = Tk()
L1 = Label(top, text="My calculator",).grid(row=0,column=1)
L2 = Label(top, text="Number 1",).grid(row=1,column=0)
L3 = Label(top, text="Number 2",).grid(row=2,column=0)
L4 = Label(top, text="Operator",).grid(row=3,column=0)
L4 = Label(top, text="Answer",).grid(row=4,column=0)
E1 = Entry(top, bd =5)
E1.grid(row=1,column=1)
E2 = Entry(top, bd =5)
E2.grid(row=2,column=1)
E3 = Entry(top, bd =5)
E3.grid(row=3,column=1)
E4 = Entry(top, bd =5)
E4.grid(row=4,column=1)
B=Button(top, text ="Submit",command = proces).grid(row=5,column=1,)

top.mainloop()
'''
def browsefunc():
    filename = filedialog.askopenfile( title = "Select file 1")
    pathlabel.config(text=filename.name)

def browsefunc2():
    filename = filedialog.askopenfile( title = "Select file 1")
    pathlabel2.config(text=filename.name)

def process():

    print("Starting:   ",pathlabel.cget("text"),"  ",pathlabel2.cget("text"))

    dfA = pd.read_excel(pathlabel.cget("text"))
    dfB = pd.read_excel(pathlabel2.cget("text"))

    NamesA = dfA['Product Name']
    NamesB = dfB['Product Name']

    print(NamesA)
    print(NamesB)

    print('Starting search :- ')

    for i in NamesA.index:
        print(i, '\n')
        root.update()
        label = tkinter.Label(root, text= "Currently done "+str(i)+" rows.\n")
        label.grid(row=7, column=1, padx=4, pady=4, sticky='ew')
        for j in NamesB.index:
            findSim(NamesA[i], NamesB[j], i, j)

    dfResult.sort_values(by='Similarity', ascending=False)
    dfResult.to_excel(writer, sheet_name='Sheet 1')
    writer.save()

root = Tk()



browsebutton = Button(root, text="Browse", command=browsefunc)
browsebutton.grid(row=0, column=0, padx=4, pady=4, sticky='ew')

pathlabel = Label(root)
pathlabel.grid(row=0, column=1, padx=4, pady=4, sticky='ew')

browsebutton2 = Button(root, text="Browse", command=browsefunc2)
browsebutton2.grid(row=3, column=0, padx=4, pady=4, sticky='ew')

pathlabel2 = Label(root)
pathlabel2.grid(row=3, column=1, padx=4, pady=4, sticky='ew')

submit_button = Button(root, text="Start", command=process)
submit_button.grid(row=5, column=1, padx=4, pady=4, sticky='ew')


root.mainloop()







