import requests

import re
import math
import numpy
import getpass
import json
import csv
import datetime
import calendar
import argparse
global prtg_password
prtg_password = 'M9y%23asABUx9svvs'
now = datetime.datetime.now()
default_start = now - datetime.timedelta(days = 14)
default_end = now - datetime.timedelta(days = 1)

parser = argparse.ArgumentParser()
parser.add_argument('--username', default='agriffin', required=False,
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
args = parser.parse_args()

PRTG_HOSTNAME = 'nanm.smartaira.net'


# Prompt the user for credentials
#prtg_password = getpass.getpass('Password: ')

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
    if response.status_code == 200:
        response_tree = json.loads(response.text)
        if args.sensorid:
            for i in response_tree.get('sensors'):
                if i['objid'] == int(args.sensorid):
                    return [i]
        else:
            return response_tree.get('sensors')
    else:
        print("Error in sensor list request")
        exit(1)

def out_csv(row, outfile, mode):
    """ Output to a csv file

    Arguments:
    row -- list with output for row
    outfile -- destination output filename
    mode -- 'w' or 'a' (overwrite or append)
    """
    with open(outfile, mode, newline='') as csvfile:
        csvout = csv.writer(csvfile)
        csvout.writerow(row)

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

summary_data = []

sensors = get_kpi_sensor_ids(args.username, prtg_password)
output_file = f'output_{args.start}--{args.end}.csv'
if args.output:
    if args.output.endswith('/'):
        output_file = args.output + output_file
    else:
        output_file = args.output + '/' + output_file

headers = ['Location', f'Max Traffic (mbps) ({int(args.percentile)}%)',
         'Choke Point', 'Choke Point Limit (mbps)', 'Circuit Max Limit (mbps)', 'Choke Point Utilization',
         'Circuit Utilization', 'Action']
if args.debug:
    headers.insert(1, 'Device id')
    headers.insert(1, 'Device')

out_csv(headers, output_file, 'w')


for sensor in sensors:
    response = requests.get(
            f'https://{PRTG_HOSTNAME}/api/historicdata.json?id={sensor["objid"]}'
            f'&avg={args.avgint}&sdate={args.start}-00-00&edate={args.end}-23-59'
            f'&usecaption=1'
            f'&username={args.username}&password={prtg_password}', verify=False
            )
    if response.status_code == 200:

        data = json.loads(response.text)

        properties = extract_tags(sensor)

        traffic_in = normalize_traffic(data, 'Traffic In (speed)')

        traffic_out = normalize_traffic(data, 'Traffic Out (speed)')
        
        device_name = sensor['device'].split(' (')[0]

        out_array = []

        # Location
        if properties.get('kpi_siteid'):
            out_array.append(re.sub("#", " ", properties['kpi_siteid']))
        else:
            out_array.append('NA')

        if args.debug:
            # Device name (debug)
            out_array.append(device_name)

            # Device id (debug)
            out_array.append(sensor['objid'])

        # Max Traffic
        max_traffic = 0
        if properties.get('kpi_trafficdirection') == 'up':
            if traffic_out == []:
                out_array.append("NA")
            else:
                max_traffic = math.ceil(numpy.percentile(traffic_out, int(args.percentile)))
                out_array.append(max_traffic)
        else:
            if traffic_in == []:
                out_array.append("NA")
            else:
                max_traffic = math.ceil(numpy.percentile(traffic_in, int(args.percentile)))
                out_array.append(max_traffic)

        if "Core" in properties.get('kpi_seg'):
            summary_data.append({'segment': properties.get('kpi_seg'),
                                 'bandwidth': max_traffic,
                                 'limit': properties.get('kpi_cktmaxlimit')})

        # Choke point 'kpi_choke'
        if properties.get('kpi_choke'):
            out_array.append(properties['kpi_choke'])
        else:
            out_array.append('NA')

        # Choke point limit 'kpi_chokelimit'
        if properties.get('kpi_chokelimit'):
            out_array.append(properties['kpi_chokelimit'])
        else:
            out_array.append('NA')

        # Circuit max limit 'kpi_cktmaxlimit'
        if properties.get('kpi_cktmaxlimit'):
            out_array.append(properties['kpi_cktmaxlimit'])
        else:
            out_array.append('NA')

        # Choke point utilization
        # max down / choke point limit
        if properties.get('kpi_chokelimit'):
            out_array.append(max_traffic / int(properties['kpi_chokelimit']))
        else:
            out_array.append('NA')

        # Circuit utilization
        # max down / circuit max limit
        if properties.get('kpi_cktmaxlimit'):
            out_array.append(max_traffic / int(properties['kpi_cktmaxlimit']))
        else:
            out_array.append('NA')
        
        out_csv(out_array, output_file, 'a')


out_csv([], output_file, 'a')
out_csv(['Core Utilization Summary'], output_file, 'a')
out_csv(['Core', 'Bandwidth', 'Max Capacity', 'Utilization'], output_file, 'a')
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
    out_csv([segment, segment_bandwidth, segment_limit, saturation], output_file, 'a')

segment_saturation = segment_bandwidth_total / segment_capacity_total
out_csv(['Total:', segment_bandwidth_total, segment_capacity_total, segment_saturation], output_file, 'a')

