
from tkinter import *
from tkinter import ttk
from tkinter.messagebox import showinfo
import requests  # line
import pandas as pd  # export
import xlsxwriter
from tkinter import messagebox
import json

url = 'https://notify-api.line.me/api/notify'
dataframe = pd.read_excel(r'fruits.xlsx')
print(dataframe)
username = []
token = []
for i in dataframe['Username']:
    username.append(i)
for i in dataframe['Token']:
    token.append(i)
print(username)
print(token)
root = Tk()
root.geometry("375x250")
root.resizable(False, False)
root.title('โปรแกรทแจ้งเตือนพัสดุผู้อยุ่อาศัยในหอ')


def action():
    name = entryName.get()
    token = entrytoken.get()
    print(name, token)

    # อ่านข้อมูลที่มีอยู่ในไฟล์เดิม
    readDataframe = pd.read_excel(r'fruits.xlsx')

    # สร้างข้อมูลใหม่เป็นข้อมูลของ orange
    newDataframe = pd.DataFrame({'Username': [name], 'Token':  [token]})

    # นำข้อมูล orange ที่สร้างใหม่รวมเข้ากับข้อมูลเก่าที่อ่านจากไฟล์
    frames = [readDataframe, newDataframe]
    result = pd.concat(frames)

    # สร้าง Writer เหมือนกับตอนเขียนไฟล์
    writer = pd.ExcelWriter('fruits.xlsx', engine='xlsxwriter')

    # นำข้อมูลชุดใหม่เขียนลงไฟล์และจบการทำงาน
    result.to_excel(writer, index=False)
    writer.save()
    dataframe = pd.read_excel(r'fruits.xlsx')
    username = []
    token = []
    for i in dataframe['Username']:
        username.append(i)
    for i in dataframe['Token']:
        token.append(i)
    month_cb['values'] = username
    messagebox.showinfo("รายละเอียดโปรแกรม", "เพิ่มข้อมูลผู้ใช้สำเร็จแล้ว")

def month_changed():
    dataframe = pd.read_excel(r'fruits.xlsx')
    username = []
    token = []
    for i in dataframe['Username']:
        username.append(i)
    for i in dataframe['Token']:
        token.append(i)
    
    for x in range(len(username)):
        if username[x] == month_cb.get():
            headers = {'content-type': 'application/x-www-form-urlencoded',
                       'Authorization': 'Bearer '+token[x]}
            msg = f' พัสดุขอท่าน  {month_cb.get()} อยู่ที่ห้องนิติแล้ว!'
            print(msg)
            r = requests.post(url, headers=headers, data={'message': msg})
            # print(type(r) ,r.json())
            httppsot = (r.json())
            print((httppsot["status"]))
            # print(json.dumps(r.json()))
            if httppsot["status"] == 200:
                messagebox.showinfo("Response 200 ", "แจ้งเตือนสำเร็จแล้ว")
            else :
                messagebox.showinfo("Response 401 ", "แจ้งเตือนผิดพลาด")
            break


label1 = ttk.Label(text="โปรแกรมแจ้งเตือนพัสดุผู้อยู่อาศัยในหอพัก")
label1.place(x=10, y=0)

lblName = Label(root, text="Enter your name : ", width=20)
lblName.place(x=10, y=30)
entryName = Entry(root, width=25)
entryName.place(x=175, y=30)


lbltoken = Label(root, text="Enter your token : ", width=20)
lbltoken.place(x=10, y=60)
entrytoken = Entry(root, width=25)
entrytoken.place(x=175, y=60)


# Create button to execute action
btnAction = Button(root, text="เพิ่มชื่อผู้อยู่", width=25, command=action)
btnAction.place(x=90, y=90)


label2 = ttk.Label(text="โปรแกรมแจ้งเตือนพัสดุผู้อยู่อาศัยในหอพัก")
label2.place(x=10, y=130)

label3 = Label(text="แจ้งเตือนถึง :", width=20)
label3.place(x=10, y=160)
selected_month = StringVar()
month_cb = ttk.Combobox(root, textvariable=selected_month, width=25)
month_cb['values'] = username
month_cb['state'] = 'readonly'  # normal
month_cb.place(x=175, y=160)

btn1 = Button(text="แจ้งเตือนพัสดุ", width=25,
              command=month_changed).place(x=90, y=200)
root.mainloop()
