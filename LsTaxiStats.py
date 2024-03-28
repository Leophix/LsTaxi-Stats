# -*- coding: utf-8 -*-
import tkinter as tk
import time,datetime
import pandas as pd

from googleapiclient.discovery import build

import pickle
import os.path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import gspread

import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning) 

def gsheet_api_check(SCOPES):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds



class Travail():
  def __init__(self,Nom="Nom",Prenom="Prenom",Grade="Grade",Salaire=0.):
    self.Nom=Nom
    self.Prenom=Prenom
    self.Grade=Grade
    self.Salaire=Salaire
    

def pull_sheet_data(SCOPES,SPREADSHEET_ID,DATA_TO_PULL):
    creds = gsheet_api_check(SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    rows = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                              range=DATA_TO_PULL).execute()
    data = rows.get('values')
    return data



def PRINT(DF):
    with pd.option_context('display.max_rows', 1000, 'display.max_columns', 1000):
        print("\n\n",DF)

def update_time():
    global current_time
    current_time = datetime.datetime(*time.localtime()[:6])
    clock_label.config(text=current_time)
    clock_label.after(1000, update_time)

def EventTime():
    global LastEvent
    LastEvent=time.time()
    
def AddEvent(Event,Compta,Reason):
    if not current_time in Data.index:
        sec=int(round(time.time()-LastEvent,0))
        Data.loc[current_time]=[Uservar.get(),str(int(sec//60.))+" min "+str(sec%60)+" s",Event,Compta,Reason]
        if Event=="Fin de Service":
            Compta=-CalcSalaire()
            Data.at[current_time,"Comptabilite"]=Compta
        EventTime()
        #PRINT(Data)
    
def DelCourse():
    Data.drop(index=Data.loc[Data["Evenement"]=="Fin de course"].index[-1],inplace=True)
    
def CalcSalaire():
    Debut=Data.loc[Data["Evenement"]=="Prise de Service"].index[-1]
    Fin=Data.loc[Data["Evenement"]=="Fin de Service"].index[-1]
    TempsService=(Fin-Debut).total_seconds()/60
    return TempsService//15*Employe[Uservar.get()].Salaire
    
def UpdateLab(lab,value):
  lab.config(text=value)

def Enable(Widgets,TurnOn=True):
  if Widgets:
    if not type(Widgets)==list:
      Widgets=[Widgets]
    for i in Widgets:
      if TurnOn:
        i.config(state="normal")
      else:
        i.config(state="disabled")
        
        
def Save():
    global Data
    gc = gspread.authorize(gsheet_api_check(SCOPES))
    SPREADSHEET_ID = '1voUvSq6nIkc8-iDPBXCs2VtT9ItV1hA0EA1-ogDQ4NU'
    sht1 = gc.open_by_key(SPREADSHEET_ID)
    
    try:
        worksheet = sht1.worksheet(Uservar.get())
    except gspread.WorksheetNotFound:
        worksheet = sht1.add_worksheet(title=Uservar.get(),rows=0,cols=0)
        worksheet = sht1.worksheet(Uservar.get())
    if not worksheet.find(Data.iloc[-1].name.strftime("%d/%m/%Y, %H:%M:%S")):
        LastRow=worksheet.findall(Uservar.get())[-1].row+1 if worksheet.findall(Uservar.get()) else 1
        if not worksheet.find(current_time.date().strftime("%d/%m/%Y")):
            worksheet.update("A"+str(LastRow+2),[[Data.loc[Data["Evenement"]=="Prise de Service"].index[-1].strftime("%d/%m/%Y")]])
            worksheet.update("A"+str(LastRow+3)+":F"+str(LastRow+3),[["Date et Heure"]+[d for d in Data.columns]])
            LastRow+=4
            
        worksheet.update("A"+str(LastRow)+":A"+str(LastRow+len(Data)),[[d] for d in Data.index.strftime("%d/%m/%Y, %H:%M:%S")])
        worksheet.update("B"+str(LastRow)+":F"+str(LastRow+len(Data)),Data.to_numpy().tolist())
        Data=pd.DataFrame(columns=["Qui ?","Temps entre evenements","Evenement","Comptabilite","Raison"])
    
    with open("User.txt","w")as f:
        f.write(Uservar.get()+"\n"+VarCP)
        
def ChangeUser(*args):
    with open("User.txt","w")as f:
        f.write(Uservar.get()+"\n"+VarCP)
        
        
if not os.path.exists("USER.txt"):
    with open("User.txt","w")as f:
        f.write("User\n0")
    
with open("User.txt","r")as f:
    USER,VarCP=list(f)
    USER=USER[:-1]
     


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1ebpSxOrjkF2P3tSAXusP6FKscL_Pysnk3TEcPnkSz54'
DATA_TO_PULL = "'🧍Ressources Humaines'!C3:F11"
data = pull_sheet_data(SCOPES,SPREADSHEET_ID,DATA_TO_PULL)
df = pd.DataFrame(data, columns=["Employés","Numéro","Grade","Salaire"])



Employe={}
for i,j,k in zip(df["Employés"],df["Grade"],df["Salaire"]):
    if i:
        Employe[i]=Travail(Nom=i.split(" ")[0],Prenom=i.split(" ")[1],Grade=j,Salaire=int(k))


Data=pd.DataFrame(columns=["Qui ?","Temps entre evenements","Evenement","Comptabilite","Raison"])    
    
  
window = tk.Tk()
window.title("Statistiques LS Taxi")
window.geometry("350x540")
window.columnconfigure(0, minsize=175, weight=0)
window.columnconfigure(1, minsize=165, weight=0)


clock_label = tk.Label(window, font=("Arial", 12), fg="black")
clock_label.grid(row=6, column=0,columnspan=2,padx=10,pady=7)
Uservar=tk.StringVar(value=USER)
user=tk.OptionMenu(window, Uservar,*list(Employe.keys()),command=lambda*args:(ChangeUser(),Enable(PService)))
user.grid(row=0, column=0,padx=5,pady=7,sticky="w")


CourseFrame=tk.LabelFrame(window, text="Nombre de course", font=("Arial", 12),fg="goldenrod3")
CourseFrame.grid(row=2, column=0,columnspan=2)
Courses=tk.Label(CourseFrame, text="0", font=("Arial", 24),width=3)
Courses.grid(row=0, column=1,padx=10,pady=7)
Down=tk.Button(CourseFrame, text="\\/", command=lambda:(UpdateLab(Courses,str(int(Courses.cget("text"))-1))if Courses.cget("text")!="0" else Courses.cget("text"),DelCourse()),fg="red", font=("Arial Black", 18,"bold"),width=3)
Down.grid(row=0, column=0,padx=10,pady=7)
Up=tk.Button(CourseFrame, text="/\\", command=lambda:(UpdateLab(Courses,str(int(Courses.cget("text"))+1)),AddEvent("Fin de course", "+300", "Fin de course")),fg="forest green", font=("Arial Black", 18,"bold"),width=3)
Up.grid(row=0, column=2,padx=10,pady=7)


EssenceFrame=tk.LabelFrame(window, text="Essence", font=("Arial", 12),fg="purple")
EssenceFrame.grid(row=3, column=1)
Essence=tk.Button(EssenceFrame, text="Plein d'essence", command=lambda:(AddEvent("Plein d'essence", "-40", "Plein d'essence")), font=("Arial", 11))
Essence.grid(row=0, column=0,padx=10,pady=7)

ReparationFrame=tk.LabelFrame(window, text="Reparation", font=("Arial", 12),fg="turquoise3")
ReparationFrame.grid(row=3, column=0, padx=10, pady=7)
Bennys=tk.Button(ReparationFrame, text="Rep Benny's", command=lambda:(AddEvent("Reparation Benny's", "", "")), font=("Arial", 11))
Bennys.grid(row=0, column=0,padx=10,pady=7)
Nord=tk.Button(ReparationFrame, text="Rep Nord", command=lambda:(AddEvent("Reparation Nord", "-500", "Reparation Nord")), font=("Arial", 11))
Nord.grid(row=1, column=0,padx=10,pady=7)

ClientFrame=tk.LabelFrame(window,text="Client inconscient", font=("Arial", 10))
ClientFrame.grid(row=4, column=0,columnspan=2,padx=10,pady=7)
Client=tk.Button(ClientFrame, text="Nouvel inconscient", command=lambda:(AddEvent("Inconscient", "", "")), font=("Arial", 11))
Client.grid(row=0, column=0,padx=10,pady=7)

CoursesPrec=tk.Label(window, font=("Arial", 12), fg="black",text="Nombre de courses precemment faites : "+str(VarCP))
CoursesPrec.grid(row=7, column=0,columnspan=2,padx=10,pady=7)

FService=tk.Button(window, text="Fin de Service",command=lambda:(Enable(widgets,False),Enable(PService),AddEvent("Fin de Service", None,"Salaire"),globals().update(VarCP=Courses.cget("text")),UpdateLab(CoursesPrec,"Nombre de courses precemment faites : "+VarCP),UpdateLab(Courses,"0"),Save()), font=("Arial", 18), fg="red")
FService.grid(row=5, column=0,columnspan=2,padx=10,pady=7)
widgets=[Down,Up,Essence,Bennys,Nord,Client,FService]
Enable(widgets,False)

PService=tk.Button(window, text="Prise de Service", command=lambda : (EventTime(),Enable(widgets),Enable(PService,False),AddEvent("Prise de Service","","")), font=("Arial", 18), fg="forest green")
PService.grid(row=1, column=0,columnspan=2,padx=10,pady=7)

if Uservar.get()=="User":
    Enable(PService,False)


update_time()
window.mainloop()

try :
    Save()
except IndexError:
    pass
