import argparse
import json
import datetime
import requests

def get_kpi_sensor_ids(username, password, PRTG_HOSTNAME, args):
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
            f'&username={username}&password={password}&sortby=device',verify=False
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

def prtgMainCall(sensordata,PRTG_HOSTNAME,PRTG_PASSWORD,cliargs):
    
    i = 8
    for sensor in sensordata:
        response = requests.get(
            f'https://{PRTG_HOSTNAME}/api/historicdata.json?id={sensor["objid"]}'
            f'&avg={cliargs.avgint}&sdate={cliargs.start}-00-00&edate={cliargs.end}-23-59'
            f'&usecaption=1'
            f'&username={cliargs.username}&password={PRTG_PASSWORD}', verify=False
            )
        if response.status_code == 200:
            print("Initial PRTG API query successful")
            response_j = response.json()
            print(len(response_j['histdata']))
            #storeAPIResponse(response_j,sensor)


        else:
            print("Error making 'main' API call to nanm.smartaira.net (PRTG)")
            print("HTTP response 200: OK was not received")
            print("Received response code: "+str(response.status_code))
            exit(1)

        i += 1


def storeAPIResponse(historicResponseData,sensor):
    boxedJSONData = json.load(historicResponseData)
def cliArgumentParser(currentSystemDatetime):
    now = datetime.datetime.now()
    default_start = now - datetime.timedelta(days = 28)
    default_end = now - datetime.timedelta(days = 0)
    parser = argparse.ArgumentParser()
    parser.add_argument('--username', required=False, default="agriffin", ###### !! CHANGE FOR PROD !! #####
                    help='PRTG username for API call')
    parser.add_argument('--start', default=default_start.strftime('%Y-%m-%d'),
                    help='Historic data start date (yyyy-mm-dd)')
    parser.add_argument('--end', default=default_end.strftime('%Y-%m-%d'),
                    help='Historic data end date (yyyy-mm-dd)')
    parser.add_argument('--avgint', default="21600",
                    help='Averaging interval. Smaller numbers increase api call time!'
                        ' Default is 21600 seconds (6 hours)')
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


PRTG_USERNAME = 'agriffin'
PRTG_PASSWORD = 'M9y%23asABUx9svvs'
PRTG_HOSTNAME = "nanm.smartaira.net"

now = datetime.datetime.now()

cliargs = cliArgumentParser(now)

response_tree = get_kpi_sensor_ids(PRTG_USERNAME,PRTG_PASSWORD,PRTG_HOSTNAME,cliargs)

prtgMainCall(response_tree,PRTG_HOSTNAME,PRTG_PASSWORD,cliargs)