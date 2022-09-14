# --------   PRTG-XLSX-Report-Generator.py
# -------------------------------------------------------------------------------
#
#   Pulls sensor & device data from PRTG API and neatly formats the data into a .xmlx (MS Excel) file
#       using python-openpyxl, json, csv, and pandas. 
#   
#       Average runtime: 274 seconds (4.5 - 5.0 minutes)
# -------------------------------------------------------------------------------
from time import time
import requests
import re
import math
import numpy
import getpass
import json

import datetime
import argparse
import openpyxl as opyxl
import pandas as pd
import os
import subprocess
import logging

'''---------------------------------------------------------------------------------------------'''
'''---------------------------------------- FUNCTIONS ------------------------------------------'''
'''---------------------------------------------------------------------------------------------'''
###########################################################
### [Most primary functions declared here]

#############
### [Efficiency Checking Imports]
######################
import cProfile
import pstats
import timeit   


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
PRTG_HOSTNAME = 'nanm.bluerim.net'   ###!! Static, domain/URL to PRTG server
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
default_start,default_end = timeWindowFrames(0)

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

global primary_df
global secondary_df
global summary_df
primary_df = pd.DataFrame(
   {'Location':["","","","","",],
    'MaxTraffic':["","","","","",],
    'ChokePoint':["","","","","",],
    'ChokePointLimit':["","","","","",],        
    'CircuitMaxLimit':["","","","","",],        
    'CircuitUtilization':["","","","","",],        
    'ChokePointUtilization0':["","","","","",],        
    'ChokePointUtilization1':["","","","","",],        
    'ChokePointUtilization2':["","","","","",],        
    'ChokePointUtilization3':["","","","","",],          
    'MaxPlanUsage':["","","","","",],            
    'Notes':["","","","","",],    
    'Action':["","","","","",]})

secondary_df = pd.DataFrame(
   {'kpi_choke':["","","","","",],
    'kpi_chokelimit':["","","","","",],
    'kpi_cktmaxlimit':["","","","","",],
    'ChokePointLimit':["","","","","",]})

summary_df = pd.DataFrame(
   {'Core':["","","","","",],
    'Bandwidth':["","","","","",],
    'MaxCapacity':["","","","","",],
    'Utilization0':["","","","","",],
    'Utilization1':["","","","","",],
    'Utilization2':["","","","","",]})



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



#     summary table on top and ensure that nothing is overwritten/deleted]
#############

            

### [Beginning initial PRTG API call and assigning data to "sensors" var for manipulation]
#############
sensorsMainCall = get_kpi_sensor_ids(cliargs.username, PRTG_PASSWORD)

### [CREATION OF COMPLETE FILE (CURRENLT JUST THE SUMMARY)]
#############
def summary_out():
    segments = set()
    for data in summary_data:
        # Create set from all the segment values (creates a unique list)
        segments.add(data.get('segment'))

    segment_bandwidth_total = (1*10^(-25))
    segment_capacity_total = (1*10^(-25))

    i = 0
    for segment in segments:
        segment_bandwidth = (1*10^(-25))
        segment_limit = (1*10^(-25))
        for data in summary_data:
            if data['segment'] == segment:
                segment_bandwidth += int(data['bandwidth'])
                segment_limit += int(data['limit'])
        saturation = segment_bandwidth / segment_limit
        segment_bandwidth_total += segment_bandwidth
        segment_capacity_total += segment_limit
        segment_saturation = segment_bandwidth_total / segment_capacity_total

        summary_df['Core'][i].append(segment)
        summary_df['Bandwidth'][i].append(segment_bandwidth)
        summary_df['MaxCapacity'][i].append(segment_limit)
        summary_df['Utilization'][i].append(saturation)

        summary_df['Core'][i].append('Total')
        summary_df['Bandwidth'][i].append(segment_bandwidth_total)
        summary_df['MaxCapacity'][i].append(segment_bandwidth_total)
        summary_df['Utilization'][i].append(segment_saturation)        



def extraChokeUtilCalc(PRTG_HOSTNAME,cliargs,PRTG_PASSWORD,summary_data,output_file_TMP,sensorsMainCall,sensor,i_index):
    for k in range(1,3):
        Tstart,Tend = timeWindowFrames(k)
        response = requests.get(
                    f'https://{PRTG_HOSTNAME}/api/historicdata.json?id={sensor["objid"]}'
                    f'&avg={cliargs.avgint}&sdate={Tstart}-00-00&edate={Tend}-23-59'
                    f'&usecaption=1'
                    f'&username={cliargs.username}&password={PRTG_PASSWORD}'
                    )
        if response.status_code == 200:
            data = json.loads(response.text)
            properties = extract_tags(sensor)
            traffic_in = normalize_traffic(data, 'Traffic In (speed)')
            traffic_out = normalize_traffic(data, 'Traffic Out (speed)')
            device_name = sensor['device'].split(' (')[0]



            ### [dec] - [CHOKE POINT LIMIT (Mb/s)]
            #############
            if properties.get('kpi_chokelimit'):
                secondary_df['kpi_chokelimit'][i_index].append(properties['kpi_chokelimit'])
            else:
                secondary_df['kpi_chokelimit'][i_index].append('NA')


            ### [dec] - [MAX TRAFFIC (Mb/s)]
            #############
            max_traffic = 0
            if properties.get('kpi_trafficdirection') == 'up':
                if traffic_out == []:
                    primary_df['MaxTraffic'][i_index].append("NA")
                else:
                    max_traffic = math.ceil(numpy.percentile(traffic_out, int(cliargs.percentile)))
                    primary_df['MaxTraffic'][i_index].append(max_traffic)
            else:
                if traffic_in == []:
                    primary_df['MaxTraffic'][i_index].append("NA")
                else:
                    max_traffic = math.ceil(numpy.percentile(traffic_in, int(cliargs.percentile)))
                    primary_df['MaxTraffic'][i_index].append(max_traffic)

            if "Core" in properties.get('kpi_seg'):
                summary_data.append({'segment': properties.get('kpi_seg'),
                                    'bandwidth': max_traffic,
                                    'limit': properties.get('kpi_cktmaxlimit')})
        

            ### [dec] - [CHOKE POINT UTILIZATION (%)]
            #############
            if properties.get('kpi_chokelimit'):
                secondary_df['kpi_chokelimit'][i_index].append(max_traffic / int(properties['kpi_chokelimit']))
            else:
                secondary_df['kpi_chokelimit'][i_index].append('NA')
            
             
        else:
            print("Error making API call to nanm.smartaira.net (PRTG)")
            print("HTTP response 200: OK was not received")
            print("Received response code: "+str(response.status_code))
            exit(1)


def sensorsFrameCall(PRTG_HOSTNAME,cliargs,PRTG_PASSWORD,summary_data,output_file_TMP,sensorsMainCall):
    i = 5
    for sensor in sensorsMainCall:
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
            

            ### [dec] - [LOCATION (Location)]
            #############
            if properties.get('kpi_siteid'):
                primary_df['Location'][i].append(re.sub("#", " ", properties['kpi_siteid']))
            else:
                primary_df['Location'][i].append('NA')

            if cliargs.debug:
                # Device name (debug)
                primary_df['Location'][i].append(device_name)

                # Device id (debug)
                primary_df['Location'][i].append(sensor['objid'])

            ### [dec] - [MAX TRAFFIC (Mb/s)]
            #############
            max_traffic = 0
            if properties.get('kpi_trafficdirection') == 'up':
                if traffic_out == []:
                    primary_df['MaxTraffic'][i].append("NA")
                else:
                    max_traffic = math.ceil(numpy.percentile(traffic_out, int(cliargs.percentile)))
                    primary_df['MaxTraffic'][i].append(max_traffic)
            else:
                if traffic_in == []:
                     primary_df['MaxTraffic'][i].append("NA")
                else:
                    max_traffic = math.ceil(numpy.percentile(traffic_in, int(cliargs.percentile)))
                    primary_df['MaxTraffic'][i].append(max_traffic)

            if "Core" in properties.get('kpi_seg'):
                summary_data.append({'segment': properties.get('kpi_seg'),
                                    'bandwidth': max_traffic,
                                    'limit': properties.get('kpi_cktmaxlimit')})

            ### [dec] - [CHOKE POINT (Device)]
            #############
            if properties.get('kpi_choke'):
                secondary_df['kpi_choke'][i].append(properties['kpi_choke'])
            else:
                secondary_df['kpi_choke'][i].append('NA')

            ### [dec] - [CHOKE POINT LIMIT (Mb/s)]
            #############
            if properties.get('kpi_chokelimit'):
                secondary_df['kpi_chokelimit'][i].append(properties['kpi_chokelimit'])
            else:
                secondary_df['kpi_chokelimit'][i].append('NA')
            


            ### [dec] - [CIRCUIT MAX LIMIT (Mb/s)]
            #############
            if properties.get('kpi_cktmaxlimit'):
                secondary_df['kpi_cktmaxlimit'][i].append(properties['kpi_cktmaxlimit'])
            else:
                secondary_df['kpi_cktmaxlimit'][i].append('NA')

            ### [dec] - [CIRCUIT UTILIZATION (%)]
            #############
            if properties.get('kpi_cktmaxlimit'):
                secondary_df['kpi_cktmaxlimit'][i].append(max_traffic / int(properties['kpi_cktmaxlimit']))
            else:
                secondary_df['kpi_cktmaxlimit'][i].append('NA')


            ### [dec] - [CHOKE POINT UTILIZATION (%)]
            #############
            if properties.get('kpi_chokelimit'):
                secondary_df['kpi_chokelimit'][i].append(max_traffic / int(properties['kpi_chokelimit']))
            else:
                secondary_df['kpi_chokelimit'][i].append('NA')         

            out_array_w_extras = extraChokeUtilCalc(PRTG_HOSTNAME,cliargs,PRTG_PASSWORD,summary_data,output_file_TMP,sensorsMainCall,sensor,i)

            i += 1
            ### [io] - [Writing newly modified array data to 'output_file_TMP' file.]
            #############
        else:
            print("Error making API call to nanm.smartaira.net (PRTG)")
            print("HTTP response 200: OK was not received")
            print("Received response code: "+str(response.status_code))
            exit(1)



### [Array to be used for piping TMP file lines into final file]
#############
summary_data = []
summary_data.clear() # Flushing array values just in case -- I dont feel like exception handling and this is easier


### [PRTG API call and assigning data to "sensors" var]
#############
sensorsFrameCall(PRTG_HOSTNAME,cliargs,PRTG_PASSWORD,summary_data,output_file_TMP,sensorsMainCall)

### [Calling summary_out to analyze data from TMP file and create Summary table in COMP file]



### [Inserting/appending temp file data into Complete/Main output file]
#    [Two separate files (TMP & COMP) are used because Brant wants the summary table at the top of the document, but 
#   the summary table is generated last. The only workaround with the CSV module in Python that I've found works without issue is to
#   create two documents, one stores the normal device data temporarily. Once the workload is done the table is generated as normal, but is placed into a new document
#   so it is at the top. All TMP data is then filed in below the table.]
#############


































##########################
###################################  [!!] The block of docstring-commented code below was primarily just for testing
################################### though it can still be used to run performance tests on this program/script
##########################          otherwise it can be deleted as well. I left it commented out just in case.
'''###################


# Nothing below this line is very important 
# Mostly testing tools   
----------------------------

############
### [cProfile profiling. Outputs to stdout by default. To save output to a file ...
###     ... run python3 kpi_report_bandwidth.py > performanceOutput.txt]
def programPerformanceProfile():
    profiler = cProfile.Profile()
    profiler.enable()
    main()
    profiler.disable()
    stats = pstats.Stats(profiler).sort_stats('cumtime')
    stats.print_stats() 

if __name__ == "__main__":
    
### Though this script doesn't need to be fast or hyper-efficient, it does take ~5-7 minutes on average to finish,
# and I gotfrom several performance tests of the script

    import cProfile, pstats
    import time

    ############
    ### [Testing total runtime (to check efficiency and look for improvements)]
    ###     [Can be safely ignored or removed]
    def runtimeTimer():
        effTimeBegin = time.time()
        main()
        print("--- %s seconds ---" % (time.time() - effTimeBegin))
        return effTimeBegin

    ############
    ### [Only one performance test can be run at a time in order to not stress the hardware or ...
    ###     ... flood PRTG with API calls.]

    performanceTesterVAL = 2   ###  Can be 1 or 2
                               ### 1 = cProfile
                               ### 2 = runtimeTimer
    if performanceTesterVAL == 1:
        perfOut_cProfile = programPerformanceProfile()   ##   Running cProfile and saving results to 'perfOut_cProfile' var.
    elif performanceTesterVAL == 2:
        runtimeTimer()              ##   Running time-to-execute timer. 
    else:
        main()

'''