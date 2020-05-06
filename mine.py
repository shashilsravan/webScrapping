import webbrowser
import requests
import json
from openpyxl import Workbook

# ChIJm5MaboqXyzsR-xPxquRpWss

API_KEY = 'AIzaSyBGBT5ghLBk8nGYDxf6RfMxuNb8P7ojnDg'
from googleplaces import GooglePlaces, types, lang

google_places = GooglePlaces(API_KEY)
pid = input("Enter Place ID: ")
xmlfile = google_places.get_place(place_id=pid)
url = "https://maps.googleapis.com/maps/api/place/details/json?place_id=" + pid + "&fields=name,rating,review&key=" + API_KEY

# webbrowser.open_new(url=url)

response = requests.get(url)
with open('main.json', 'wb') as file:
    file.write(response.content)

with open('main.json', 'rb') as file:
    string = json.load(file)
webpage = string["result"]
reviews = string["result"]["reviews"]

name = string["result"]["name"]
rating = string["result"]["rating"]


workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "Author name"
sheet["B1"] = "Language"
sheet["C1"] = "Rating"
sheet["D1"] = "Review"

for i in range(2, 7):
    sheet["A"+str(i)] = reviews[i-2]["author_name"]
    sheet["B" + str(i)] = reviews[i - 2]["language"]
    sheet["C" + str(i)] = reviews[i - 2]["rating"]
    sheet["D" + str(i)] = reviews[i - 2]["text"]

workbook.save(filename=name+str(rating)+".xlsx")


# webbrowser.open_new(webpage)