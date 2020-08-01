import pandas as pd
from bs4 import BeautifulSoup
import csv
import requests
import time
import re
from csv import DictWriter

#configurationa
file_name = 'hall_tickets.xlsx'
sheet_name = 'Urdu and English'
index_value = [1,2,3]
request_url='http://mahresult.nic.in/sscmarch2020/sscresultviewmarch.asp'
file_header=['seat_no','name','mother_name','english','marathi','hindi','mathematics','social','science','sanskrit','urdu','marathihindi','tamil','hindi-sanskrit','retail','result','total_marks','percentage','out_of','_']

counter=1
#Importing Data
dataset=pd.read_excel(file_name,sheet_name,header=None)
all_data=dataset.iloc[0:,index_value].values

def request_data(url,data):
    time.sleep(3)
    headers={'Content-Type': 'application/x-www-form-urlencoded',
    'Referer': 'http://mahresult.nic.in/sscmarch2020/sscmarch2020.htm'}

    r = requests.post(url, data = f'regno={data["seat_no"]}&mname={data["mother_name"]}',headers=headers)
    return {'content':r.content,**data}

def process_data(data):
    collected_data={}
    soup = BeautifulSoup(data['content'], 'lxml')
    validating_data=soup.find(class_="bg-danger") or soup.find(class_="resultviewbtn")
    if validating_data==None:
        table=soup.find("table")
        table_rows = table.find_all('tr')
        selected_rows=[]
        final_data={}
        for tr in table_rows:
            td = tr.find_all('td')
            row = [re.sub(r'[^A-Za-z0-9_ ]+', "", i.text) for i in td]
            if(row!=[] and len(row)>=2):
                if(len(row)==3):
                    if row[2].isdigit():
                        final_data[row[1].split(" ", 1)[0].lower()]=float(row[2])
                    else:
                        final_data[row[1].split(" ", 1)[0].lower()]=row[2]
                elif len(row)==4:
                    final_data[row[0].replace(" ", "_").lower()]=row[1]
                    final_data[row[2].replace(" ", "_").lower()]=row[3]
        collected_data.update(final_data)
    collected_data['seat_no']=data['seat_no']
    collected_data['name']=data['name']
    collected_data['mother_name']=data['mother_name']
    return collected_data

def export_data(data):
    print(data)
    with open('results.csv', 'a+', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=file_header)
        # writer.writeheader()
        writer.writerow(data)

#Looping over the data row wise
for row_data in all_data:
    if  pd.isna(row_data[0]):
        row_data[0]='NA'
    if pd.isna(row_data[1]):
        row_data[1]='NA'
    if pd.isna(row_data[2]):
        row_data[2]='NA'

    data={
        'seat_no':row_data[0].strip(),
        'name':row_data[1].strip(),
        'mother_name':row_data[2].strip()
    }
    response_data=request_data(request_url,data)
    processed_data=process_data(response_data)
    export_data(processed_data)
    print(counter)
    counter=counter+1
