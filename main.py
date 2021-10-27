from imap_tools import MailBox, AND
from redminelib import Redmine
import pandas
import os
from bs4 import BeautifulSoup

rm_url = "http://10.77.1.70/" # Боевой
#rm_url = "http://10.77.1.60/" # Тестовый
api_key = "762a0212b89630a2e83c77fdd5b71c70739daf45" #70
#api_key = "4de8d2365b04c77f605ed100f3636be0166948f0" #60
rm_project = "it"
redmine = Redmine(url=rm_url, key=api_key)
file='/scripts/auto_create_issues_from_email/1.xlsx'
mail1='rep@vox-com.ru'
mail2='solominns@zdrav.mos.ru'
#mail2='abzaliyev.valentin@vox-com.ru'

with MailBox('imap.yandex.com').login('naumen_stp_report@vox-com.ru', 'qwe123QWE') as mailbox: #коннектимся к почте и смотрим последнее письмо, скачиваем файл из письма
    uid=[]
    uid=[msg.uid for msg in mailbox.fetch()]
    from1=[]
    from1=[msg.from_ for msg in mailbox.fetch()]
    #print(uid)
    x=0
    while x!=len(uid):
        if from1[x]!=mail1 and from1[x]!=mail2:
            mailbox.delete(uid[x])
            #print("1")
        x+=1
    z=0
    uid1=[]
    uid1=[msg.uid for msg in mailbox.fetch()]
    attachments=[]
    attachments = [msg.attachments for msg in mailbox.fetch()]
    subjects=[]
    subjects = [msg.subject for msg in mailbox.fetch()]
    from2=[]
    from2=[msg.from_ for msg in mailbox.fetch()]
    text=[]
    text=[msg.html for msg in mailbox.fetch()]
    while z!=len(uid1):
        if from2[z]==mail1:
            for att in attachments[z]: 
                #print(att.filename, att.content_type)
                content_type=att.content_type
                filename=att.filename
                with open(file, 'wb') as f:
                    f.write(att.payload)
            excel_data_df = pandas.read_excel(file, sheet_name='Отчет', engine='openpyxl', header=1) #Читаем файл
            if any(excel_data_df)==True: #Проверка файла на наличие записи и если такова есть, то создаем заявку в редмайн
                #print(excel_data_df)
                redmine.issue.create(project_id=rm_project, subject=subjects[z], status_id=1, priority_id=3, assigned_to_id=46, custom_fields=[{'id':3, 'value':'Все'},{'id':13, 'value':'Без проекта'}], uploads=[{'path':file,'content_type':content_type,'filename':filename}], description='Ошибки в оргструктуре Naumen, все подробности в приложеном файле')
            os.remove(file) #удаляем файл
        else:
            soup = BeautifulSoup(text[z],'lxml')   
            parse=soup.get_text(' ')
            redmine.issue.create(project_id=rm_project, subject=subjects[z], status_id=1, priority_id=3, assigned_to_id=46, custom_fields=[{'id':3, 'value':'Все'},{'id':13, 'value':'Без проекта'}], description=parse)
        mailbox.delete(uid1[z]) #удаляем письмо
        z+=1
