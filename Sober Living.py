# -*- coding: utf-8 -*-
"""
Created on Tue Nov 21 13:58:56 2017

@author: backesj
"""

#!/usr/bin/env python

# -*- coding: utf-8 -*-
"""
Created on Tue Nov 21 14:41:33 2017

@author: backesj
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Oct 18 08:27:40 2017

@author: backesj
"""

from bs4 import BeautifulSoup as bs4
import requests
import time
import re

from tkinter import *
import os

def show_entry_fields():
    if os.path.isdir(e2.get()) == True:
        outDIR = e2.get()
    else:
        outDIR = 'H:/'
    state =  e1.get()
    outFile = e3.get()

    return state, outDIR, outFile

master = Tk()

Label(master, text="State").grid(row=0)
Label(master, text="Output Directory").grid(row=1)
Label(master, text="Output File Name").grid(row=2)


e1 = Entry(master)
e2 = Entry(master)
e3 = Entry(master)

e1.grid(row=0, column=1)
e2.grid(row=1, column=1)
e3.grid(row=2, column=1)

Button(master, text='Run', command=master.quit).grid(row=4, column=1, sticky=W, pady=1)
#Button(master, text='Run', command= show_entry_fields).grid(row=4, column=1, sticky=W, pady=4)

master.mainloop()

state, outDIR, outFile = show_entry_fields()




url = 'http://soberliving.interventionamerica.org/Searchdirectory.cfm?State=%s' % state



content = requests.get(url).content


soup = bs4(content, 'html.parser')

main = soup.find('div', attrs = {'id':"main-inner-left"})

links = []
for link in soup.find_all('a'):
    if 'ID' in link['href'] and 'soberliving' in link['href']:
        links.append(link['href'])
        
#header_list = []

fullDict = {}
iter = 0
for i, value in enumerate(links):
    if i == 1:
        print('Program Started')
        print('0% Complete')
    percentComplete = int(round(i/len(links),2)*100)
    if i > 1 and percentComplete%10 == 0:
        print(str(percentComplete) + '% Complete')
    if i == len(links)-1:
        print('Program Complete')
        time.sleep(2)
        print('Closing in...')
        time.sleep(3)
        print('3')
        time.sleep(3)
        print('2')
        time.sleep(3)            
        print('1')
    infoDict = {}
    #print(value)
    temp_content = requests.get(value).content
    time.sleep(5)
    temp_soup = bs4(temp_content, 'html.parser')
    main = temp_soup.find('div', attrs = {'id':"main-inner-left"})
    infoDict['Name'] = main.find('h1').text
    #header_list.append(header.text)
    table = main.find('table')
    addy = []
    for span in table.find_all('span'):
        addy.append(span.text.strip())
    addy2 = [re.split(r'\s{2,}', j) for j in addy if j != '' ][1]
    try:
        infoDict['Address1'] = addy2[0]
    except IndexError:
        infoDict['Address1'] = 'Entry Error'
        infoDict['error'] = value
    else:
        infoDict['Address1'] = addy2[0]
    
    try:
        infoDict['City'] = addy2[1][:addy2[1].find(',')]   
    except IndexError:
        infoDict['City'] = 'Entry Error' 
        infoDict['error'] = value
    else:
        infoDict['City'] = addy2[1][:addy2[1].find(',')]   
    
    try:
        infoDict['State'] = addy2[1][addy2[1].find(',')+2:-5]
    except IndexError:
        infoDict['State'] = 'Entry Error'
        infoDict['error'] = value
    else:
        infoDict['State'] = addy2[1][addy2[1].find(',')+2:-5]
    
    try:
        infoDict['Zip'] = addy2[1][-5:]
    except IndexError:
        infoDict['Zip'] = 'Entry Error'
        infoDict['error'] = value
    else:
        infoDict['Zip'] = addy2[1][-5:]
    fullDict['%s' % i] = infoDict

    
import pandas as pd
df = pd.DataFrame.from_dict(fullDict, orient='Index')                

# create writer
writer = pd.ExcelWriter('%s/%s.xlsx' % (outDIR, outFile))
# write file to outDIR renamed
df.to_excel(writer,"Sheet1", index=False)
    
writer.close()
