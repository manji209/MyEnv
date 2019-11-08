import requests
import datetime


#ITEM Tracker API
base_URL = "http://192.168.1.50:8000/itemtracker/item-cleaner/"

api_key = "5699xLucky90201"
todayDate = datetime.datetime.now()
currentDate = todayDate.strftime("%Y-%m-%d")

URL = base_URL + api_key + "/" + currentDate + "/"

#Get response
r = requests.get(URL)

if r:
    print("True", r)
else:
    print("False", r)
