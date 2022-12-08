import xml.etree.ElementTree as ET
import csv
from datetime import datetime, timedelta
import matplotlib.pyplot as plt

import os
import glob

# Set the directory where the files are located
directory = 'I:\\uProg\\uSportData\\mp ohl'
# Use the basename function to get the last folder name in the path
folder_name = os.path.basename(directory)
# Set the pattern to match the files you want to process
pattern = '*.tcx'

# Use the glob module to find the files that match the pattern
files = glob.glob(os.path.join(directory, pattern))

strXMLSchemas='{http://www.garmin.com/xmlschemas/TrainingCenterDatabase/v2}'
# import openpyxl
import xlsxwriter
# from openpyxl.chart import LineChart, Reference

headers=[
        'times'
        ,'lats'
        ,'longs'
        ,'dists'
        ,'alts'
        ,'hrs'
        ,'cads'
        ,'durations'
        ]
# Open an Excel file for writing
# wb = openpyxl.Workbook()
workbook = xlsxwriter.Workbook(folder_name+'.xlsx')
ws4Charts = workbook.add_worksheet('charts')
chartAlt = workbook.add_chart({'type': 'scatter'})
chartHR = workbook.add_chart({'type': 'scatter'})
chartCad = workbook.add_chart({'type': 'scatter'})
chartTime = workbook.add_chart({'type': 'scatter'})

# Iterate through the files and process them
for file in files:
    # Use the basename function to get the file name
    file_name = os.path.basename(file)

    # Create a new sheet
    # sheet = wb.create_sheet(file_name)
    worksheet = workbook.add_worksheet(file_name)
    # # Iterate through the data and write each value to a cell in the first row of the sheet
    # for i, value in enumerate(headers):
    #     sheet.cell(row=1, column=i+1).value = value
    worksheet.write_row('A1', headers)


    # Open the file for reading
    tcsFilePath=file
    csvFilePath=tcsFilePath.replace('.tcx','.csv')
    # Parse the XML data from the TCX file
    tree = ET.parse(tcsFilePath)
    root = tree.getroot()

    times=[];lats=[];longs=[];alts=[];dists=[];hrs=[];cads=[]

    # Iterate through the XML data and extract the relevant information
    for lap in root.findall('.//'+strXMLSchemas+\
        'Lap'):
        start_time = lap.get('StartTime')
        total_time = lap.find('.//'+strXMLSchemas+\
        'TotalTimeSeconds').text
        distance = lap.find('.//'+strXMLSchemas+\
        'DistanceMeters').text
        max_speed = lap.find('.//'+strXMLSchemas+\
        'MaximumSpeed').text

        # Print the extracted data
        print('Start Time:', start_time)
        print('Total Time:', total_time)
        print('Distance:', distance)
        print('Max Speed:', max_speed)



        # Iterate through the XML data and extract the relevant information
        for trackpoint in lap.findall('.//'+strXMLSchemas+\
            'Trackpoint'):
            time = trackpoint.find('.//'+strXMLSchemas+\
                    'Time').text
            times.append(time[:-1])
            latitude = trackpoint.find('.//'+strXMLSchemas+\
                    'LatitudeDegrees').text
            lats.append(float(latitude))
            longitude = trackpoint.find('.//'+strXMLSchemas+\
                    'LongitudeDegrees').text
            longs.append(float(longitude))
            altitude = trackpoint.find('.//'+strXMLSchemas+\
                    'AltitudeMeters').text
            alts.append(float(altitude))
            distance = trackpoint.find('.//'+strXMLSchemas+\
                    'DistanceMeters').text
            dists.append(float(distance))
            elemHR=[]
            elemHR = trackpoint.find('.//'+strXMLSchemas+\
                    'Value')
            if elemHR is not None:
                hr=elemHR.text
                hrs.append(float(hr))
            else:
                hrs.append(0)        
            elmCad=[]
            elmCad = trackpoint.find('.//'+strXMLSchemas+\
                    'Cadence')
            if elmCad is not None:
                cad=elmCad.text
                cads.append(float(cad))
            else:
                cads.append(0)
            # Print the extracted data
    # Parse the first time into a datetime object
    start_time = datetime.fromisoformat(times[0])

    # Create an empty list to store the durations
    durations = []

    # Iterate through the times and calculate the durations
    for time in times:
        # Parse the current time into a datetime object
        current_time = datetime.fromisoformat(time)

        # Calculate the difference between the times in seconds
        difference = (current_time - start_time).total_seconds()

        # Convert the difference to a timedelta object
        duration = timedelta(seconds=difference)

        # Add the duration to the list
        durations.append(duration.total_seconds())

    tcxData=[
        times
        ,lats
        ,longs
        ,dists
        ,alts
        ,hrs
        ,cads
        ,durations
        ]
    # Transpose the data so that the rows and columns are switched
    tcxData_transposed = list(zip(*tcxData))
    # Open a file for writing
    with open(csvFilePath, 'w', newline='') as csvfile:
        # Create a CSV writer
        writer = csv.writer(csvfile)

        # Write the data to the file
        writer.writerows(tcxData_transposed)
    # # Iterate through the data and write each row to the sheet
    # for i, row in enumerate(tcxData_transposed):
    #     for j, value in enumerate(row):
    #         # if j>0:
    #         #     sheet.cell(row=i+2, column=j+1,value = value).number_format = '0.00'
    #         # else:
    #         #     sheet.cell(row=i+2, column=j+1).value = value
    #         sheet.cell(row=i+2, column=j+1).value = value
    worksheet.write_column('A2', times)
    worksheet.write_column('B2', lats)
    worksheet.write_column('C2', longs)
    worksheet.write_column('D2', dists)
    worksheet.write_column('E2', alts)
    worksheet.write_column('F2', hrs)
    worksheet.write_column('G2', cads)
    worksheet.write_column('H2', durations)


    # Plot the data
    # plt.plot(dists, alts)

    # Add labels to the axes
    # plt.xlabel('distance')
    # plt.ylabel('alt')

    # Show the plot
    #plt.show()
    #######################################################################
    #
    # Create a new scatter chart.
    #

    # # Configure the first series.
    chartAlt.add_series({
        'name': file_name+' Alt',
        'categories': '='+file_name+'!D2:D'+str(len(times)+1),
        'values': '='+file_name+'!E2:E'+str(len(times)+1),
    })
    chartHR.add_series({
        'name': file_name+' HR',
        'categories': '='+file_name+'!D2:D'+str(len(times)+1),
        'values': '='+file_name+'!F2:F'+str(len(times)+1),
    })
    chartCad.add_series({
        'name': file_name+' Cad',
        'categories': '='+file_name+'!D2:D'+str(len(times)+1),
        'values': '='+file_name+'!G2:G'+str(len(times)+1),
    })
    chartTime.add_series({
        'name': file_name+' Time',
        'categories': '='+file_name+'!D2:D'+str(len(times)+1),
        'values': '='+file_name+'!H2:H'+str(len(times)+1),
    })
    # data = Reference(sheet, min_row=2, max_row=len(times)+1, min_col=4, max_col=5)
    # chart1 = LineChart()
    # chart1.add_data(data)
    # Add the chart to the sheet
    pass

ws4Charts.insert_chart('C1',chartAlt)
ws4Charts.insert_chart('C16',chartHR)
ws4Charts.insert_chart('K1',chartCad)
ws4Charts.insert_chart('K16',chartTime)
workbook.close()

pass