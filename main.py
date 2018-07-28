from tkinter import filedialog
from tkinter import *
import time
import tkinter
import pandas as pd
from pandas import *

#dataframe and writer

dfResult = pd.DataFrame(columns=['Index1', 'Name1', 'Index2', 'Name2', 'Common', 'Similarity'])
writer = pd.ExcelWriter('result.xlsx',engine= 'xlsxwriter')

#Function that compares elements
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
    if len(common)*100/len(total) > 10:
        data = {'Index1': intA, 'Name1': stringA, 'Index2': intB, 'Name2': stringB, 'Common': len(common),
                'Similarity': len(common) * 100 / len(total)}
        dfResult = dfResult.append(data,ignore_index = True)
        #print('Line number : ',intA,' and ',intB,'\nThat is :-',stringA , ' &&& ', stringB)

#Fucntion to open first file
def browsefunc():
    filename = filedialog.askopenfile( title = "Select file 1")
    pathlabel1.config(text=filename.name)

#Function to open second file
def browsefunc2():
    filename = filedialog.askopenfile( title = "Select file 1")
    pathlabel2.config(text=filename.name)

#Function that opens files and calls the compare function
def process():

    print("Starting:   ",pathlabel1.cget("text"),"  ",pathlabel2.cget("text"))

    start_time1 = time.time()

    dfA = pd.read_excel(pathlabel1.cget("text"))
    dfB = pd.read_excel(pathlabel2.cget("text"))

    NamesA = dfA['Product Name']
    NamesB = dfB['Product Name']

    print(NamesA)
    print(NamesB)

    print('Starting search :- ')

    global label

    start_time = time.time()

    label = tkinter.Label(root)
    label.grid(row=7, column=1, padx=4, pady=4, sticky='ew')
    label.config(width = 60, height = 3)

    for i in NamesA.index:
        print(i, '\n')
        for j in NamesB.index:
            if int(time.time()-start_time)>0:
                start_time = time.time()
                root.update()
                label = tkinter.Label(root, text="Currently done %d percentage.\nTime Taken till now = %d secs."
                                                 "\n Further time required = %d secs" % (
                    int(i * 100 / len(NamesA.index)), int(time.time() - start_time1), int((len(NamesA.index)-i)*(time.time() - start_time1)/i)))
                label.grid(row=7, column=1, padx=4, pady=4, sticky='ew')
                root.update()
            findSim(NamesA[i], NamesB[j], i, j)

    print("Writing result :- \n")

    dfResult.sort_values(by='Similarity', ascending=False, inplace= True)
    dfResult.to_excel(writer, sheet_name='Sheet 1')
    writer.save()

#Main

root = Tk()

root.title("Sentence Matching Problem")

browsebutton = Button(root, text="Browse", command=browsefunc)
browsebutton.grid(row=0, column=0, padx=4, pady=4, sticky='ew')

pathlabel1 = Label(root)
pathlabel1.grid(row=0, column=1, padx=4, pady=4, sticky='ew')
pathlabel1.config(width = 60, height = 2)

browsebutton2 = Button(root, text="Browse", command=browsefunc2)
browsebutton2.grid(row=3, column=0, padx=4, pady=4, sticky='ew')

pathlabel2 = Label(root)
pathlabel2.grid(row=3, column=1, padx=4, pady=4, sticky='ew')
pathlabel2.config(width = 60, height = 2)

submit_button = Button(root, text="Start", command=process)
submit_button.grid(row=5, column=1, padx=4, pady=4, sticky='ew')


root.mainloop()







