# --------   PRTG-XLSX-Report-Generator.py
# -------------------------------------------------------------------------------
#
#   Pulls sensor & device data from PRTG API and neatly formats the data into a .xmlx (MS Excel) file
#       using python-openpyxl, json, csv, and pandas. 
#   
#       Average runtime: 274 seconds (4.5 - 5.0 minutes)
# -------------------------------------------------------------------------------
from time import time, time_ns
from turtle import width
import requests
import re
import math
import numpy
import getpass
import json

import datetime
import argparse
import openpyxl as opyxl
import os

###########################################################
### [Most primary functions declared here]

### [Getting script's current working directory to be used later to ensure that 
#       no exceptions arise and for safer file IO]
#############
CWD_unsanitized = os.getcwd()
CWD_backslashes = CWD_unsanitized+"/"
CWD = CWD_backslashes.replace("\\","/")

### [Timeframes/Windows For Pulling Historical Data from PRTG]
#############
def timeWindowFrames(timeFrameIDRAW):
    timeFrameID = str(timeFrameIDRAW)
    #############
    ### [Time window declarations]
    ### [A positive time in these comments indicates # of days prior to current (DAY 0)]
    ### [eg: "0d -- 14d" = "From today (0d) through 14 days ago (14d)"]
    #############
    current_sys_datetime = datetime.datetime.now()
    ######    0d -- 14d
    def_s = current_sys_datetime - datetime.timedelta(days = 14)
    def_e = current_sys_datetime - datetime.timedelta(days = 1)
    ######    7d -- 21d
    win1_s = current_sys_datetime - datetime.timedelta(days = 21)
    win1_e = current_sys_datetime - datetime.timedelta(days = 7)
    ######    14d -- 28d
    win2_s = current_sys_datetime - datetime.timedelta(days = 28)
    win2_e = current_sys_datetime - datetime.timedelta(days = 14)
    ######    21d -- 35d
    win3_s = current_sys_datetime - datetime.timedelta(days = 35)
    win3_e = current_sys_datetime - datetime.timedelta(days = 21)

    # return [def_s,def_e],[win1_s,win1_e],[win2_s,win2_e],[win3_s,win3_e]
    if timeFrameID == "0":
        return def_s,def_e
    elif timeFrameID == "1":
        return win1_s,win1_e
    elif timeFrameID == "2":
        return win2_s,win2_e
    elif timeFrameID == "3":
        return win3_s,win3_e
    else: 
        print("Error: timeWindowFrames -- Invalid arguments passed")
### [Defining path to Temporary file]
#############

### [CLI argument parser; --username is required]
#############
def cliArgumentParser():

    default_start,default_end = timeWindowFrames("0")

    parser = argparse.ArgumentParser()
    parser.add_argument('--username', required=False, default="agriffin", ###### !! CHANGE FOR PROD !! #####
                    help='PRTG username for API call')
    parser.add_argument('--start', default=default_start.strftime('%Y-%m-%d'),
                    help='Historic data start date (yyyy-mm-dd)')
    parser.add_argument('--end', default=default_end.strftime('%Y-%m-%d'),
                    help='Historic data end date (yyyy-mm-dd)')
    parser.add_argument('--avgint', default="3600",
                    help='Averaging interval. Smaller numbers increase api call time!'
                        ' Default is 3600')
    parser.add_argument('--debug', action='store_true', dest='debug',
                    help='add additional debugging fields to output')
    parser.add_argument('--percentile', default="99",
                    help='set the percentile for reporting (default is 90)')
    parser.add_argument('--output',
                    help='specify output file directory (default is cwd)')
    parser.add_argument('--sensorid',
                    help='Pull data from only one sensorid')
    cliargs = parser.parse_args()
    return cliargs

### [User credential prompts and opts]
#############
PRTG_PASSWORD = "M9y%23asABUx9svvs"  ###### !! CHANGE FOR PROD !! ##### ----------
PRTG_HOSTNAME = 'nanm.smartaira.net'   ###!! Static, domain/URL to PRTG server
#PRTG_PASSWORD = getpass.getpass('Password: ')


### [!] [Defining cliargs from returned param of cliArgumentParser() {To be stored globally}]
#############
global cliargs ### I know, globals are bad, but it saves a lot of typing in this situation
global default_start
global default_end
global current_sys_datetime
### [!] [Assigning values to global vars]
#############
cliargs = cliArgumentParser()
#default_start,default_end = timeWindowFrames(0)

### [Timeframes/Windows For Pulling Historical Data from PRTG]
#############
current_sys_datetime = datetime.datetime.now()
### [Defining path to Complete/Summary file]
############


####################################################################################################
'''----------------------------------------------------------------------------------------------'''
'''-------------------------------------------- MAIN --------------------------------------------'''
'''----------------------------------------------------------------------------------------------'''
####################################################################################################
global outputMainSheet

outputWorkbook = opyxl.Workbook()
outputMainSheet = outputWorkbook.active

t0BACK_headers_u,t14BACK_headers_u = timeWindowFrames("0")
t7BACK_headers_u,t21BACK_headers_u = timeWindowFrames("1")
t14BACK_headers_u,t28BACK_headers_u = timeWindowFrames("2")
t21BACK_headers_u,t35BACK_headers_u = timeWindowFrames("3")

t0BACK_headers = t0BACK_headers_u.strftime('%m/%d')
t7BACK_headers = t7BACK_headers_u.strftime('%m/%d')
t14BACK_headers = t14BACK_headers_u.strftime('%m/%d')
t21BACK_headers = t21BACK_headers_u.strftime('%m/%d')
t28BACK_headers = t28BACK_headers_u.strftime('%m/%d')
t35BACK_headers = t35BACK_headers_u.strftime('%m/%d')


sheetHeaders = ['Location','Max Traffic (Mb/s)','Choke Point','Choke Point Limit (Mb/s)','Circuit Max Limit (Mb/s)','Circuit Utilization (%)',
    f'Choke Utilization ({t0BACK_headers} - {t14BACK_headers}) (%)',f'Choke Utilization ({t7BACK_headers} - {t21BACK_headers}) (%)',f'Choke Utilization ({t14BACK_headers} - {t28BACK_headers}) (%)',f'Choke Utilization ({t21BACK_headers} - {t35BACK_headers}) (%)',
    'Max Usage Plan','Notes','Action']

coreUtilSummaryHeaders = ['Core Utilization Summary','Bandwidth (Mb/s)',
    'Gross Utilization (Current) (%)','Gross Utilization (7 Days Ago - 14 Days Ago) (%)','Gross Utilization (14 - 21 Days Ago) (%)']

alphabetArray = ['A','B','C','D','E','F','G','H','I','J','K','L','M']

for letter in alphabetArray:
    outputMainSheet.column_dimensions[str(letter)].width = '25'
outputMainSheet.column_dimensions['A'].width = '25'

for i in range(0,len(sheetHeaders)):
    outputMainSheet[str(alphabetArray[i])+'7']=sheetHeaders[i]
    outputMainSheet[str(alphabetArray[i])+'7'].alignment = opyxl.styles.Alignment(horizontal='general', vertical='bottom', text_rotation=0, wrap_text=True, shrink_to_fit=False, indent=0)


for i in range(0,len(coreUtilSummaryHeaders)):
    outputMainSheet[str(alphabetArray[i])+'1']=coreUtilSummaryHeaders[i]
    outputMainSheet[str(alphabetArray[i])+'1'].alignment = opyxl.styles.Alignment(horizontal='general', vertical='bottom', text_rotation=0, wrap_text=True, shrink_to_fit=False, indent=0)
outputMainSheet['A5']='Total: '

### [Query/Call PRTG API]
#############
def get_kpi_sensor_ids(username, password):
    """
    Returns a dict with the following format:
    {'prtg-version': '22.1.74.1869',
    'treesize': 2,
    'sensors': [{'objid': 14398,
    'objid_raw': 14398,
    'device': 'ACA Edge (160.3.214.2)',
    'device_raw': 'ACA Edge (160.3.214.2)'},
    """
    response = requests.get(
            f'https://{PRTG_HOSTNAME}/api/table.json?content=sensors&output=json'
            f'&columns=objid,device,tags&filter_tags=kpi_bandwidth'
            f'&username={username}&password={password}&sortby=device', verify=False
            )
    ### If 200 OK HTTP response is not seen, raise error and print cause to terminal
    if response.status_code == 200:
        response_tree = json.loads(response.text)
        if cliargs.sensorid:
            for i in response_tree.get('sensors'):
                if i['objid'] == int(cliargs.sensorid):
                    return [i]
        else:
            return response_tree.get('sensors')
    else:
        print("Error making API call to nanm.smartaira.net (PRTG)")
        print("HTTP response 200: OK was not received")
        print("Received response code: "+str(response.status_code))
        quit()


### [Converts all speed values to Mb/s]
#############
def normalize_traffic(data, label):
    """
    Takes an input of raw PRTG historic data and the label (ie 'Traffic In (speed)')
    and multiplies the speeds by 0.00008 to get values in mbits/sec ( ((8 / 10) / 100) / 1000)
    """
    traffic_list = []
    for i in data['histdata']:
        if i[label] != '':
            traffic_list.append(i[label] * 0.000008)
    return traffic_list

### [Extract desired information via tags from PRTG API]
#   [Delimits incoming PRTG JSON tags to avoid conflictions and exceptions]
#############
def extract_tags(sensor):
    """
    Takes a sensor dictionary and extacts a properties dictionary from the tags
    string returned by PRTG. The tags string will have a format something like this:
    'kpi_bandwidth kpi_seg=DIA kpi_choke=Circuit kpi_chokelimit=10000 kpi_cktmaxlimit=10000'
    
    """
    target_tags = ['kpi_seg', 'kpi_choke', 'kpi_cktmaxlimit', 'kpi_siteid']
    def filter_tags(tags_list, property_string):
        """
        Takes a list of tags and splits them into a properties dict:
        ['kpi_choke=Circuit', kpi_chokelimit=10000'] becomes
        {'kpi_choke':'Circuit', 'kpi_chokelimit':'10000'}
        
        This function exists because we might have tags like kpi_choke and kpi_chokelimit
        that will both be found by the _in_ function
        """
        properties = {}
        property_list = list(filter(lambda a: property_string in a, tags_list))
        for property in property_list:
            key, value = property.split('=')
            properties.update({key:value})
        return properties

    device_properties = {}
    tag_string = sensor['tags'].split()

    for tag in target_tags:
        device_properties.update(filter_tags(tag_string, tag))

    return device_properties


### [Beginning initial PRTG API call and assigning data to "sensors" var for manipulation]
#############
sensorsMainCall = get_kpi_sensor_ids(cliargs.username, PRTG_PASSWORD)     

def extraChokeUtilCalc(PRTG_HOSTNAME,cliargs,PRTG_PASSWORD,sensorsMainCall,sensor,i_index):
    for k in range(1,4):
        Tstart,Tend = timeWindowFrames(k)
        response = requests.get(
                    f'https://{PRTG_HOSTNAME}/api/historicdata.json?id={sensor["objid"]}'
                    f'&avg={cliargs.avgint}&sdate={Tstart}-00-00&edate={Tend}-23-59'
                    f'&usecaption=1'
                    f'&username={cliargs.username}&password={PRTG_PASSWORD}', verify=False
                    )
        if response.status_code == 200:
            data = json.loads(response.text)
            properties = extract_tags(sensor)
            traffic_in = normalize_traffic(data, 'Traffic In (speed)')
            traffic_out = normalize_traffic(data, 'Traffic Out (speed)')
            device_name = sensor['device'].split(' (')[0]

            ### [dec] - [MAX TRAFFIC (Mb/s)]
            #############
            max_traffic = 0
            if properties.get('kpi_trafficdirection') == 'up':
                if traffic_out == []:
                    max_traffic_san = 0
                else:
                    max_traffic = math.ceil(numpy.percentile(traffic_out, int(cliargs.percentile)))
                    max_traffic_san = max_traffic
            else:
                if traffic_in == []:
                    max_traffic_san = 0
                else:
                    max_traffic = math.ceil(numpy.percentile(traffic_in, int(cliargs.percentile)))
                    max_traffic_san = max_traffic

            ### [dec] - [CHOKE POINT UTILIZATION (%)]
            #############
            from openpyxl.formatting.rule import ColorScaleRule
            rule = ColorScaleRule(start_type='percentile', start_value=0, start_color='FFAA0000',
                end_type='percentile', end_value=100, end_color='FF00AA00')

            if properties.get('kpi_chokelimit'):
                outputMainSheet.cell(row=int(i_index),column=7+int(k)).value=(float(max_traffic) / float(properties['kpi_chokelimit']))
                outputMainSheet.cell(row=int(i_index),column=7+int(k)).style='Percent'
            else:
                outputMainSheet.cell(row=int(i_index),column=7+int(k)).value='NA'
            
             
        else:
            print("Error making API call to nanm.smartaira.net (PRTG)")
            print("HTTP response 200: OK was not received")
            print("Received response code: "+str(response.status_code))
            exit(1)
    k += 1
    outputWorkbook.save("test.xlsx")

def sensorsFrameCall(PRTG_HOSTNAME,cliargs,PRTG_PASSWORD,sensorsMainCall):
    i = 8
    s_count = 1
    kpi_seg_arr = []

    def outputToSheet(yIndex,xIndex,vIndex):
        outputMainSheet.cell(row=int(yIndex),column=int(xIndex)).value=vIndex

    for sensor in sensorsMainCall:
        response = requests.get(
                f'https://{PRTG_HOSTNAME}/api/historicdata.json?id={sensor["objid"]}'
                f'&avg={cliargs.avgint}&sdate={cliargs.start}-00-00&edate={cliargs.end}-23-59'
                f'&usecaption=1'
                f'&username={cliargs.username}&password={PRTG_PASSWORD}', verify=False
                )
        if response.status_code == 200:

            data = json.loads(response.text)
            properties = extract_tags(sensor)
            traffic_in = normalize_traffic(data, 'Traffic In (speed)')
            traffic_out = normalize_traffic(data, 'Traffic Out (speed)')
            device_name = sensor['device'].split(' (')[0]
            

            ### [dec] - [LOCATION (Location)]
            #############
            if properties.get('kpi_siteid'):
                outputToSheet(i,1,re.sub("#", " ", properties['kpi_siteid']))
            else:
                outputToSheet(i,1,'NA')

            if cliargs.debug:
                # Device name (debug)
                outputToSheet(i,1,device_name)
                
                # Device id (debug)
                #outputToSheet(i,1,sensor['objid'])

            ### [dec] - [MAX TRAFFIC (Mb/s)]
            #############
            max_traffic = 0
            if properties.get('kpi_trafficdirection') == 'up':
                if traffic_out == []:
                    outputToSheet(i,2,'NA')
                else:
                    max_traffic = math.ceil(numpy.percentile(traffic_out, int(cliargs.percentile)))
                    outputToSheet(i,2,max_traffic)
            else:
                if traffic_in == []:
                    outputToSheet(i,2,'NA')
                else:
                    max_traffic = math.ceil(numpy.percentile(traffic_in, int(cliargs.percentile)))
                    outputToSheet(i,2,max_traffic)

            if "Core" in properties.get('kpi_seg'):
                if properties.get('kpi_seg') in kpi_seg_arr:
                    pass
                else:
                    outputToSheet(1+s_count,1,properties.get('kpi_seg'))
                    outputToSheet(1+s_count,2,max_traffic)
                    outputToSheet(1+s_count,3,(float(max_traffic)/float(properties.get('kpi_cktmaxlimit'))))
                    outputMainSheet.cell(row=int(1+s_count),column=3).style='Percent'
                    s_count += 1
                    kpi_seg_arr.append(properties.get('kpi_seg'))

            ### [dec] - [CHOKE POINT (Device)]
            #############
            if properties.get('kpi_choke'):
                outputToSheet(i,3,properties['kpi_choke'])
            else:
                outputMainSheet.cell(row=int(i),column=3).value='NA'
                outputToSheet(i,3,'NA')

            ### [dec] - [CHOKE POINT LIMIT (Mb/s)]
            #############
            if properties.get('kpi_chokelimit'):
                outputToSheet(i,4,float(properties['kpi_chokelimit']))
                outputMainSheet.cell(row=int(i),column=int(6)).style='Percent'
            else:
                outputToSheet(i,4,'NA')
            


            ### [dec] - [CIRCUIT MAX LIMIT (Mb/s)]
            #############
            if properties.get('kpi_cktmaxlimit'):
                outputToSheet(i,5,float(properties['kpi_cktmaxlimit']))
            else:
                outputToSheet(i,5,'NA')

            ### [dec] - [CIRCUIT UTILIZATION (%)]
            #############
            if properties.get('kpi_cktmaxlimit'):
                outputToSheet(i,6,(float(max_traffic) / float(properties['kpi_cktmaxlimit'])))
                outputMainSheet.cell(row=int(i),column=int(6)).style='Percent'
            else:
                outputToSheet(i,6,'NA')


            ### [dec] - [CHOKE POINT UTILIZATION (%)]
            #############
            if properties.get('kpi_chokelimit'):
                outputToSheet(i,7,(float(max_traffic) / float(properties['kpi_chokelimit'])))
                outputMainSheet.cell(row=int(i),column=int(7)).style='Percent'
            else:
                outputToSheet(i,7,'NA') 

            extraChokeUtilCalc(PRTG_HOSTNAME,cliargs,PRTG_PASSWORD,sensorsMainCall,sensor,i)

            i += 1
        else:
            print("Error making API call to nanm.smartaira.net (PRTG)")
            print("HTTP response 200: OK was not received")
            print("Received response code: "+str(response.status_code))
            exit(1)

### [PRTG API call and assigning data to "sensors" var]
#############
sensorsFrameCall(PRTG_HOSTNAME,cliargs,PRTG_PASSWORD,sensorsMainCall)



### [Inserting/appending temp file data into Complete/Main output file]
#    [Two separate files (TMP & COMP) are used because Brant wants the summary table at the top of the document, but 
#   the summary table is generated last. The only workaround with the CSV module in Python that I've found works without issue is to
#   create two documents, one stores the normal device data temporarily. Once the workload is done the table is generated as normal, but is placed into a new document
#   so it is at the top. All TMP data is then filed in below the table.]
#############