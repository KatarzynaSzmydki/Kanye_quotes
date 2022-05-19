import win32com.client
import requests
from datetime import datetime
from tkinter import *
import time


MY_LAT = 52.767200 # Your latitude
MY_LONG = 23.191839 # Your longitude
TIME = 10

#Your position is within +5 or -5 degrees of the ISS position.

#
# parameters = {
#     "lat": MY_LAT,
#     "lng": MY_LONG,
#     "formatted": 0,
# }
#
# response = requests.get("https://api.sunrise-sunset.org/json", params=parameters)
# response.raise_for_status()
# data = response.json()
sunrise = 4 #int(data["results"]["sunrise"].split("T")[1].split(":")[0])
sunset = 20 #int(data["results"]["sunset"].split("T")[1].split(":")[0])

time_now = datetime.now()



def countdown(time_span):

    time_print = str(time.strftime('%M:%S', time.gmtime(time_span)))
    canvas.itemconfig(timer, text=time_print)

    if time_span > 0:
        global time_count
        time_count = window.after(1000, countdown, time_span - 1)
    else:
        check_iss()



def check_iss():

    response = requests.get(url="http://api.open-notify.org/iss-now.json")
    response.raise_for_status()
    data = response.json()

    iss_latitude = float(data["iss_position"]["latitude"])
    iss_longitude = float(data["iss_position"]["longitude"])

    if abs(iss_longitude) - 5 < MY_LONG < abs(iss_longitude) + 5 and abs(iss_latitude) - 5 < MY_LAT < abs(iss_latitude) + 5:
        if time_now.hour > sunset:

            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = 'katarzyna.szmydki@pmi.com'
            mail.Subject = "Look up in the sky! ISS is crossing the sky!!"

            attachment1 = r'C:\Users\kszmydki\PycharmProjects\ISS_Overhead_Notifier\OIP.jfif'
            ats = mail.Attachments
            att1 = ats.Add(attachment1, 1, 0)

            Body1 = 'It should be visible up in the sky.'
            mail.HTMLBody = '<img src="cid:{0:}" width=200 height=100> <br/><br/><b>{1:}</b>'.format(
                att1.FileName, Body1)

            mail.Send()
            countdown(TIME)

        else:
            print(f"ISS is close to you but it's sunny outside. \n"
                  f"Prepare for sunset at: {sunset}. Current hour:{time_now.hour}")
            countdown(TIME)

    else:
        print(f"ISS is far away from you. You're at: {MY_LAT};{MY_LONG}. ISS is at: {iss_latitude};{iss_longitude}")
        countdown(TIME)




window = Tk()

canvas = Canvas(width=200, height=224, highlightthickness=0)
canvas.grid(column=1,row=1)
timer = canvas.create_text(100, 130, text='00:00', fill='black', font=('Ariel', 35, 'bold'))

# Label
lbl_timer = Label(text='Timer', fg='black', font=('Ariel', 30, 'bold'))
lbl_timer.grid(column=1,row=0)

countdown(TIME)

window.mainloop()


