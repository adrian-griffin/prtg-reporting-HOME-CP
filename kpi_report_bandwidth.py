#############
### [Imports]
#############
import requests
import re
import math
import numpy
import getpass
import json
import csv
import datetime
import argparse
import openpyxl as opyxl
import pandas as pd
import os
import subprocess



#############
### [Efficiency Checking Imports]
######################
import cProfile
import pstats
import timeit

def effCalc(FUNCTIONASARG):
    effCheckProf = cProfile.Profile()
    effCheckProf.runcall(FUNCTIONASARG)
    effPS = pstats.Stats(effCheckProf)
    effPS.print_stats()
######################


#############
### [Defining GLOBAL params for easier use of functions]
#############
global cliargs

#############
### [CLI argument parser]
#############
def cliArgumentParser():
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


#############
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
    ### PRTG API call
    response = requests.get(
            f'https://{PRTG_HOSTNAME}/api/table.json?content=sensors&output=json'
            f'&columns=objid,device,tags&filter_tags=kpi_bandwidth'
            f'&username={username}&password={password}&sortby=device'
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
        print("Error making API call to nanm.bluerim.net (PRTG)")
        print("HTTP response 200: OK was not received")
        print("Received response code: "+str(response.status_code))
        exit(1)

#############
### [Output to temp file (file gets deleted later)]
#############
def out_tmp_csv(row, outfile, mode):
    """ Output to a csv file

    Arguments:
    row -- list with output for row
    outfile -- destination output filename
    mode -- 'w' or 'a' (overwrite or append)
    """
    with open(outfile, mode, newline='') as csvfile:
        csvout = csv.writer(csvfile)
        csvout.writerow(row)

#############
### [Output to final/completed file. (Summary table goes here after all data is gathered, 
#     into temporary file. It is then joined to under the summary table into the completed
#       file below)]
#############
def out_complete_csv(row, outfile, mode):
    """ Output to the smaller Summary.csv file to be joined with 
        main file after completion.

        Ensures that the Summary portion remains on top without 
        deleting/overwriting any other data on the primary sheet
    """
    with open(outfile, mode, newline='') as csvfile:
        csvout = csv.writer(csvfile)
        csvout.writerow(row)

#############
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

#############
### [Extract desired information via tags from PRTG API]
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

#############
### [Writes each line of temp file into complete file in order to keep
#     summary table on toop and ensure that nothing is overwritten/deleted]
#############
def csv_joiner(output_file_TMP,complete_file):
    complete_array = []
    complete_array.clear()
    with open(CWD+"/"+str(output_file_TMP)) as temporaryFile:
        temporaryFileLines = temporaryFile.readlines()

        for line in temporaryFileLines:
            complete_array.append(line)
    out_complete_csv(complete_array, complete_file, 'a')
    os.remove(str(CWD)+"/"+str(output_file_TMP))
            

#############
### [Defining headers to be inserted into CSV/XLSX file]
#############
def create_headers(complete_file):
    headers = ['Location', f'Max Traffic (Mb/s) ({int(cliargs.percentile)}%)',
            'Choke Point', 'Choke Point Limit (Mb/s)', 'Circuit Max Limit (Mb/s)', 'Circuit Utilization (%)',
            'Choke Point Utilization (%) T[0,-14]d','Choke Point Utilization (%) T[-7,-21]d',
            'Choke Point Utilization (%) T[-14,-28]d','Choke Point Utilization (%) T[-21,-35]d',
            'Max Usage Plan','Notes','Action']
    out_tmp_csv(headers, complete_file, 'w')
    if cliargs.debug:
        headers.insert(1, 'Device id')
        headers.insert(1, 'Device')



#############
### [CREATION OF COMPLETE FILE (CURRENLT JUST THE SUMMARY)]
#############
def summary_out(output_file_TMP,complete_file):
    #out_tmp_csv([], output_file_TMP, 'a')
    out_complete_csv(['Core Utilization Summary'], complete_file, 'a')
    out_complete_csv(['Core', 'Bandwidth', 'Max Capacity', 'Utilization'], complete_file, 'a')
    segments = set()
    for data in summary_data:
        # Create set from all the segment values (creates a unique list)
        segments.add(data.get('segment'))

    segment_bandwidth_total = 0
    segment_capacity_total = 0

    for segment in segments:
        segment_bandwidth = 0
        segment_limit = 0
        for data in summary_data:
            if data['segment'] == segment:
                segment_bandwidth += int(data['bandwidth'])
                segment_limit += int(data['limit'])
        saturation = segment_bandwidth / segment_limit
        segment_bandwidth_total += segment_bandwidth
        segment_capacity_total += segment_limit
        out_complete_csv([segment, segment_bandwidth, segment_limit, saturation,], complete_file, 'a')

    segment_saturation = segment_bandwidth_total / segment_capacity_total

    #############
    ### [WRITING SUMMARY DATA TO COMPLETE FILE]
    #############
    out_complete_csv(['Total:', segment_bandwidth_total, segment_capacity_total, segment_saturation], complete_file, 'a')

    #############
    ### [JOINING ALL DATA FROM TEMP INTO COMPLETE FILE (APPENDING @ BOTTOM)]
    #############
    csv_joiner(output_file_TMP,complete_file)

'''
####################################### MAIN #######################################
#----------------------------------------------------------------------------------#
####################################################################################
'''




#############
### [User credential prompts]
#############
###### !! CHANGE FOR PROD !! #####
#PRTG_PASSWORD = getpass.getpass('Password: ')
PRTG_PASSWORD = "M9y%23asABUx9svvs"
###!! Static, domain/URL to PRTG server
PRTG_HOSTNAME = 'nanm.bluerim.net'

#############
### [Getting script's current working directory to be used later for safer file IO]
#############
CWD_unsanitized = os.getcwd()
CWD = CWD_unsanitized.replace("\\","/")


#############
### [Time window declarations]
### [A positive time in these comments indicates # of days prior to current (DAY 0)]
### [eg: "0d -- 14d" = "From today (0d) through 14 days ago (14d)"]
#############
now = datetime.datetime.now()
###  0d -- 14d
default_start = now - datetime.timedelta(days = 14)
default_end = now - datetime.timedelta(days = 1)
###  7d -- 21d
window_roll_1_start = now - datetime.timedelta(days = 21)
window_roll_1_end = now - datetime.timedelta(days = 7)
###  14d -- 28d
window_roll_2_start = now - datetime.timedelta(days = 28)
window_roll_2_end = now - datetime.timedelta(days = 14)
###  21d -- 35d
window_roll_2_start = now - datetime.timedelta(days = 35)
window_roll_2_end = now - datetime.timedelta(days = 21)

cliargs = cliArgumentParser()
summary_data = []
summary_data.clear()

#############
### [Beginning PRTG API call and assigning data to "sensors" var]
### [Declaring paths for 'temp' and 'complete' files to be written to]
#############
sensors = get_kpi_sensor_ids(cliargs.username, PRTG_PASSWORD)
output_file_TMP = 'output_file_temporary.csv'
complete_file = f'output_{cliargs.start}--{cliargs.end}.csv'
if cliargs.output:
    if cliargs.output.endswith('/'):
        output_file_TMP = cliargs.output + output_file_TMP
    else:
        output_file_TMP_TMP = cliargs.output + '/' + output_file_TMP



### [Writing headers into summary/complete file]
#############
create_headers(complete_file)



#############
### [Calling and iterating through sensors data from PRTG]
### [Assigning incoming data to 'properties','traffic_IO',and 'device_name']
### [Selecting values to be written on each row for respective headers]
#############
for sensor in sensors:
    response = requests.get(
            f'https://{PRTG_HOSTNAME}/api/historicdata.json?id={sensor["objid"]}'
            f'&avg={cliargs.avgint}&sdate={cliargs.start}-00-00&edate={cliargs.end}-23-59'
            f'&usecaption=1'
            f'&username={cliargs.username}&password={PRTG_PASSWORD}'
            )
    if response.status_code == 200:

        data = json.loads(response.text)

        properties = extract_tags(sensor)

        traffic_in = normalize_traffic(data, 'Traffic In (speed)')

        traffic_out = normalize_traffic(data, 'Traffic Out (speed)')
        
        device_name = sensor['device'].split(' (')[0]

        out_array = []

        #############
        ### [LOCATION (Location)]
        #############
        if properties.get('kpi_siteid'):
            out_array.append(re.sub("#", " ", properties['kpi_siteid']))
        else:
            out_array.append('NA')

        if cliargs.debug:
            # Device name (debug)
            out_array.append(device_name)

            # Device id (debug)
            out_array.append(sensor['objid'])

        #############
        ### [MAX TRAFFIC (Mb/s)]
        #############
        max_traffic = 0
        if properties.get('kpi_trafficdirection') == 'up':
            if traffic_out == []:
                out_array.append("NA")
            else:
                max_traffic = math.ceil(numpy.percentile(traffic_out, int(cliargs.percentile)))
                out_array.append(max_traffic)
        else:
            if traffic_in == []:
                out_array.append("NA")
            else:
                max_traffic = math.ceil(numpy.percentile(traffic_in, int(cliargs.percentile)))
                out_array.append(max_traffic)

        if "Core" in properties.get('kpi_seg'):
            summary_data.append({'segment': properties.get('kpi_seg'),
                                 'bandwidth': max_traffic,
                                 'limit': properties.get('kpi_cktmaxlimit')})

        #############
        ### [CHOKE POINT (Device)]
        #############
        if properties.get('kpi_choke'):
            out_array.append(properties['kpi_choke'])
        else:
            out_array.append('NA')

        #############
        ### [CHOKE POINT LIMIT (Mb/s)]
        #############
        if properties.get('kpi_chokelimit'):
            out_array.append(properties['kpi_chokelimit'])
        else:
            out_array.append('NA')

        #############
        ### [CIRCUIT MAX LIMIT (Mb/s)]
        #############
        if properties.get('kpi_cktmaxlimit'):
            out_array.append(properties['kpi_cktmaxlimit'])
        else:
            out_array.append('NA')

        #############
        ### [CHOKE POINT UTILIZATION (%)]
        #############
        if properties.get('kpi_chokelimit'):
            out_array.append(max_traffic / int(properties['kpi_chokelimit']))
        else:
            out_array.append('NA')

        #############
        ### [CIRCUIT UTILIZATION (%)]
        #############
        if properties.get('kpi_cktmaxlimit'):
            out_array.append(max_traffic / int(properties['kpi_cktmaxlimit']))
        else:
            out_array.append('NA')
        ###!! Writing (appending) array to temp file
        out_tmp_csv(out_array, output_file_TMP, 'a')
    else:
        print("Error making API call to nanm.bluerim.net (PRTG)")
        print("HTTP response 200: OK was not received")
        print("Received response code: "+str(response.status_code))
        exit(1)




### [Calling summary_out to initiate]
#############
summary_out(output_file_TMP,complete_file)



