'''README      

 This script takes a csv input of addresses to find Longitude and Latitude for each address via Bing Maps API. 
     -Address columns must be labeled as: "AddressLine", "City", "State", "Zip" - input file may contain additional columns.
 Before using this script, please install the following packages: requests, pandas, and xlsxwriter.(IE: pip install requests, etc.)
 The csv, json, and os packages should have already been downloaded when installing Python 3.

 Additional Information:
 Bing Maps is free if you consume fewer than 125,000 transactions in a 12-month period.
 Get a new Bing Maps Key or check your usage here: https://www.bingmapsportal.com/Application#

 Bing API Find a Location by Address documentation: https://msdn.microsoft.com/en-us/library/ff701714.aspx
 Example url: http://dev.virtualearth.net/REST/v1/Locations/US/WA/98052/Redmond/1%20Microsoft%20Way?o=json&key=j8ROx37qaPdF6DV0GkMD~UmHYzVMvhzv6p2z2NZLyzQ~AscsmLIAN6JAxrv4UmtfNo8mUjnUb1kMDhS1tlcGcmNWVjQgHczpkrTRD3oas6vM
 
 Note: If the address is not found, Bing pulls coordinates by zip instead of throwing an error - zip was removed from request url to prevent inaccuracies.

 Additional Logic Needed:
   -Remove periods and extra spaces from address fields - do this manually before processing until incorporated
   -Delete extra files on desktop resulting from output
   -Separate the coordinates into unique columns for latitude and longitude
'''
import requests
import csv
import json
import pandas
from pandas import DataFrame
import os

#set folder path to where the new files should go
os.chdir('C:/Users/ahanway/Desktop')

# Bing Maps Key 
bingMapsKey = "oiFQyicEX3XIUtjLWQgD~DYilAL4jDZeDrX8Xgwg03A~AlWzymK5ge5rw9tffiF4yAl3JYjhRXP7uEhGtPVis02v9OiZFinI3wWkR_XFlij_"

##creates output file as csv
outputfile = open('HA_LatLongOutput.csv', 'w', newline='')
csvwriter = csv.writer(outputfile)
header = ["Bing Coordinates", "Bing Address", "Address Confidence"]
csvwriter.writerow(header)

payload = {'type': 'adminVM', 'pageSize': '100', 'filter': 'status==POWERED_ON'}

# opening input file, pulling data from Bing API json requests and writing to the output file
with open("C:/Users/ahanway/Desktop/to_geocode.csv") as csvfile: 
	rowreader = csv.reader(csvfile)
	totalrows = -1
	for row in rowreader:
		totalrows += 1		
	
with open("C:/Users/ahanway/Desktop/to_geocode.csv") as csvfile: 
	reader = csv.DictReader(csvfile)
	counter = 0
	errorcount = 0
	totalrows = int(totalrows)
	if counter <= totalrows:
		for row in reader:
			
			try:
				counter += 1
				print("Processing Input Line: " + str(counter) + " of " + str(totalrows))
				State = row['State']
				City = row['City']
				#Zip = row['Zip']
				AddressLine = row['AddressLine']
				routeUrl = "http://dev.virtualearth.net/REST/v1/Locations/" + State + "/" + City + "/" + AddressLine + "?o=json&key=" + bingMapsKey  #add zip into the request url if using zip codes
				page = requests.get((routeUrl), params=payload, verify=False).json()
				bingCoordinates = (page['resourceSets'][0]['resources'][0]['point']['coordinates'])
				bingCoordinates = str(bingCoordinates)
				
				#outputs actual address pulled as a result of the Bing request - corresponding to coordinates pulled
				bingAddress = (page['resourceSets'][0]['resources'][0]['address']['formattedAddress'])
				
				#outputs high/medium/low measure of accuracy
				addressConfidence = (page['resourceSets'][0]['resources'][0]['confidence'])
				
			except:
				print("Error processing")
				errorcount +=1
				bingCoordinates = ("Error")
				bingAddress = "Error"
				addressConfidence = "Error"
				
			with open('HA_LatLongOutput.csv') as f:
				values = bingCoordinates, bingAddress, addressConfidence
				csvwriter.writerow(values)
		
outputfile.close()
print("Finished Processing")
print("Creating Files...")

#append output file to inputfile
left = pandas.read_csv("C:/Users/ahanway/Desktop/to_geocode.csv", encoding = "ISO-8859-1")
right = pandas.read_csv('HA_LatLongOutput.csv', encoding = "ISO-8859-1")
result = left.join(right, how="inner")
result.to_csv('HA_LatLongOutputFinal.csv', index=False)

print("Total processing errors: " + str(errorcount))
print("Complete")