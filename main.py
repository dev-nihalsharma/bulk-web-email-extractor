import re
from requests_html import HTMLSession
import pandas as pd
from openpyxl import load_workbook


EMAIL_REGEX = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.com'

path =  input('Enter Excel File Path: ')
col_name =  input('Enter Websites Column Name: ')
sheet_name = input('Enter Sheet Name: ')

df = pd.read_excel(path, sheet_name=sheet_name, index_col=False)
urls = df[col_name].to_list()
emails = []


for url in urls:
    try:

        if url == '':
            emails.append('')
            continue

        session = HTMLSession()
        r = session.get(url)
    
        # for JAVA-Script driven websites  
        # r.html.render(timeout=30)
     
        match =  re.search(EMAIL_REGEX, r.html.raw_html.decode()) or ''
        email = match.group()
        print(email)
        emails.append(email)

        r.close()
        session.close()
    except Exception as e:
        emails.append('')


xlworkbook = load_workbook(path)
writer = pd.ExcelWriter(path, engine='openpyxl', mode='a') 
writer.workbook = xlworkbook

df.insert(len(df.columns), "Email", emails,True)
try:
    df.to_excel(writer,sheet_name=sheet_name + ' - email', index=False)
except Exception as e:
    print(e)

writer.close()

print('All Emails Extracted and Saved!')
