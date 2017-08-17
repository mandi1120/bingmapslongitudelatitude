## README ##
## This script takes a csv input of addresses separated by columns for AddressLine, City, State, and Zip to find 
##   Longitude and Latitude for each via Bing Maps.
## Before using this script, please install the following packages: requests, pandas, and xlsxwriter.(IE: pip install requests)
## The csv, json, and os packages should have already been downloaded when installing Python 3.
## Bing Maps is free if you consume fewer than 125,000 transactions in a 12-month period.
## Get a new Bing Maps Key or check your usage here: https://www.bingmapsportal.com/Application#
##
## Additional Information:
## Bing API Find a Location by Address documentation: https://msdn.microsoft.com/en-us/library/ff701714.aspx
##   example url: http://dev.virtualearth.net/REST/v1/Locations/US/WA/98052/Redmond/1%20Microsoft%20Way?o=json&key=j8ROx37qaPdF6DV0GkMD~UmHYzVMvhzv6p2z2NZLyzQ~AscsmLIAN6JAxrv4UmtfNo8mUjnUb1kMDhS1tlcGcmNWVjQgHczpkrTRD3oas6vM
## Note: the zip code was removed from the routeUrl to prevent inaccurate coordinates. 
##   If the address line is wrong, Bing pulls coordinates by zip instead of throwing an error.

import requests
import csv
import json
import pandas
from pandas import DataFrame
import os

#set folder path to where the files should go
os.chdir('C:/Users/ahanway/Desktop')

# Bing Maps Key 
bingMapsKey = "j8ROx37qaPdF6DV0GkMD~UmHYzVMvhzv6p2z2NZLyzQ~AscsmLIAN6JAxrv4UmtfNo8mUjnUb1kMDhS1tlcGcmNWVjQgHczpkrTRD3oas6vM"

##creates output as csv
outputfile = open('LatLongOutput.csv', 'w', newline='')
csvwriter = csv.writer(outputfile)
header = ["Coordinates", "Latitude", "Longitude"]
csvwriter.writerow(header)

payload = {'type': 'adminVM', 'pageSize': '100', 'filter': 'status==POWERED_ON'}

#opening the input file, iterating through each row to create urls with from/to address fields, 
#pulling data from json requests and writing to the output file
with open("C:/Users/ahanway/Desktop/FindingLongLat.csv") as csvfile: 
	rowreader = csv.reader(csvfile)
	totalrows = -1
	for row in rowreader:
		totalrows += 1		
	
with open("C:/Users/ahanway/Desktop/FindingLongLat.csv") as csvfile: 
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
				Zip = row['Zip']
				AddressLine = row['AddressLine']
				routeUrl = "http://dev.virtualearth.net/REST/v1/Locations/" + State + "/" + City + "/" + AddressLine + "?o=json&key=" + bingMapsKey
				#print(routeUrl)
				page = requests.get((routeUrl), params=payload, verify=False).json()
				coordinates = (page['resourceSets'][0]['resources'][0]['point']['coordinates'])
				coordinates = str(coordinates)
				latitude = str(coordinates[0:16])
				longitude = coordinates[1]	

			except:
				print("Error processing")
				errorcount +=1
				coordinates = ("Error")
				latitude = "0"
				longitude = "0"
				
			with open('LatLongOutput.csv') as f:
				values = coordinates, latitude, longitude
				csvwriter.writerow(values)
		
outputfile.close()
print("Finished Processing")
print("Creating Files...")

#append output file to inputfile
left = pandas.read_csv('C:/Users/ahanway/Desktop/FindingLongLat.csv', encoding = "ISO-8859-1")
right = pandas.read_csv('LatLongOutput.csv', encoding = "ISO-8859-1")
result = left.join(right, how="inner")
result.to_csv('LatLongOutput.csv', index=False)

print("Total processing errors: " + str(errorcount))
print("Complete")