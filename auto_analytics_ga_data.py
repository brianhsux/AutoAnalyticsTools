"""A simple example of how to access the Google Analytics API."""

import argparse
import xlwt
import xlrd
import os, sys
import json

from apiclient.discovery import build
import httplib2
from oauth2client import client
from oauth2client import file
from oauth2client import tools
from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf

def get_service(api_name, api_version, scope, client_secrets_path):
    """Get a service that communicates to a Google API.

    Args:
      api_name: string The name of the api to connect to.
      api_version: string The api version to connect to.
      scope: A list of strings representing the auth scopes to authorize for the
        connection.
      client_secrets_path: string A path to a valid client secrets file.

    Returns:
      A service that is connected to the specified API.
    """
    # Parse command-line arguments.
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        parents=[tools.argparser])
    flags = parser.parse_args([])

    # Set up a Flow object to be used if we need to authenticate.
    flow = client.flow_from_clientsecrets(
        client_secrets_path, scope=scope,
        message=tools.message_if_missing(client_secrets_path))

    # Prepare credentials, and authorize HTTP object with them.
    # If the credentials don't exist or are invalid run through the native client
    # flow. The Storage object will ensure that if successful the good
    # credentials will get written back to a file.
    storage = file.Storage(api_name + '.dat')
    credentials = storage.get()
    if credentials is None or credentials.invalid:
      credentials = tools.run_flow(flow, storage, flags)
    http = credentials.authorize(http=httplib2.Http())

    # Build the service object.
    service = build(api_name, api_version, http=http)

    return service


def get_first_profile_id(service):
    # Use the Analytics service object to get the first profile id.

    # Get a list of all Google Analytics accounts for the authorized user.
    accounts = service.management().accounts().list().execute()

    if accounts.get('items'):
      # Get the first Google Analytics account.
      account = accounts.get('items')[0].get('id')

      # Get a list of all the properties for the first account.
      properties = service.management().webproperties().list(
          accountId=account).execute()

      if properties.get('items'):
        # Get the first property id.
        property = properties.get('items')[0].get('id')

        # Get a list of all views (profiles) for the first property.
        profiles = service.management().profiles().list(
            accountId=account,
            webPropertyId=property).execute()

        if profiles.get('items'):
          # return the first view (profile) id.
          return profiles.get('items')[0].get('id')

    return None


def get_results(service, profile_id):
    # Use the Analytics Service Object to query the Core Reporting API
    # for the number of sessions in the past seven days.
    return service.data().ga().get(
        ids='ga:' + profile_id,
        start_date='7daysAgo',
        end_date='today',
        metrics='ga:sessions').execute()

def get_gaData(service, profile_id, start_date, end_date, metrics):
    # Use the Analytics Service Object to query the Core Reporting API
    # for the number of sessions in the past seven days.
    api_query = service.data().ga().get(
      ids='ga:' + profile_id,
      start_date=start_date,
      end_date=end_date,
      metrics=metrics)

    return api_query.execute()

def get_gaMau(service, profile_id, start_date, end_date, metrics, dimensions):
    # Use the Analytics Service Object to query the Core Reporting API
    # for the number of sessions in the past seven days.
    api_query = service.data().ga().get(
      ids='ga:' + profile_id,
      start_date=start_date,
      end_date=end_date,
      metrics=metrics,
      dimensions=dimensions)

    return api_query.execute()

def get_segmentMau(service, profile_id, start_date, end_date, metrics, dimensions, segment):
    # Use the Analytics Service Object to query the Core Reporting API
    # for the number of sessions in the past seven days.
    api_query = service.data().ga().get(
      ids='ga:' + profile_id,
      start_date=start_date,
      end_date=end_date,
      metrics=metrics,
      dimensions=dimensions,
      segment=segment)

    return api_query.execute()

def print_results(results, type):
    if results:
      if (type == 'general_result'):
        if (results.get('profileInfo').get('profileId') == '103299109'):
            print('View (Profile): 100%')
        else:
            print('View (Profile): 10%')
        
        print ('Total Sessions: %s' % results.get('rows')[0][0])
      elif (type == 'mau_result'):
        if (results.get('profileInfo').get('profileId') == '103299109'):
            print('View (Profile): 100%')
        else:
            print('View (Profile): 10%')
        
        print ('MAU: %s' % results.get('rows')[0][1])
      elif (type == 'segment_mau_result'):
        if (results.get('profileInfo').get('profileId') == '103299109'):
            print('View (Profile): 100%')
        else:
            print('View (Profile): 10%')
        
        print ('Segment MAU: %s' % results.get('rows')[0][1])
    else:
      print ('No results found')

def print_mau(results):
    if results:
      if (results.get('profileInfo').get('profileId') == '103299109'):
          print('View (Profile): 100%')
      else:
          print('View (Profile): 10%')
      
      print ('MAU: %s' % results.get('rows')[0][1])

    else:
      print ('No results found')

def get_results_gaid(results):
    # return data nicely for the user.
    if results:    
      return (results.get('profileInfo').get('profileId'))

    else:
      print ('No results found')  

def get_results_value(results):
    # return data nicely for the user.
    if results:    
      try:
        return (results.get('rows')[0][0])
      except TypeError:
        return 0

    else:
      print ('No results found')  

def get_results_mau(results):
    # return data nicely for the user.
    if results:    
      try:
        return (results.get('rows')[0][1])
      except TypeError:
        return 0

    else:
      print ('No results found')  

def getDataPermonth(year, month, segment_new_mechanism):
    # Define the auth scopes to request.
    scope = ['https://www.googleapis.com/auth/analytics.readonly']

    # Authenticate and construct service.
    service = get_service('analytics', 'v3', scope, 'client_secrets.json')

    start_date_day = '01'
    end_date_day = None

    # gaid with 100%
    gaid_100p = '103299109'
    # gaid with 10%
    gaid_10p = '121989812'

    # segment_new_mechanism_pre = 'sessions::condition::ga:appVersion[]'
    # segment_new_mechanism_content = ('1.6.0.38_161017|1.6.0.39_161027|1.6.0.42_161122|1.6.0.46_161209|'
    #   '1.6.0.52_161227|1.6.0.56_170103|1.6.0.58_170117|1.6.0.59_170120|1.6.0.60_170222')
    # segment_new_mechanism = segment_new_mechanism_pre + segment_new_mechanism_content

    if (month == 1 or month == 3 or month == 5 or month == 7 or month == 8 or month == 10 or month == 12):
        end_date_day = 31
    elif (month == 4 or month == 6 or month == 9 or month == 11):
        end_date_day = 30
    elif (month == 2):
        end_date_day = 28 

    if (month < 10):
        start_date_combine = str(year) + '-0' + str(month) + '-' + start_date_day
        end_date_combine = str(year) + '-0' + str(month) + '-' + str(end_date_day)
        time_str = str(year) + '0' + str(month)
    else:
        start_date_combine = str(year) + '-' + str(month) + '-' + start_date_day
        end_date_combine = str(year) + '-' + str(month) + '-' + str(end_date_day)
        time_str = str(year) + str(month)

    sessions_results_100p = (get_results_value(get_gaData(service, gaid_100p, start_date_combine, 
      end_date_combine, 'ga:sessions')))
    screenviews_results_100p = (get_results_value(get_gaData(service, gaid_100p, start_date_combine, 
        end_date_combine, 'ga:screenviews')))
    screenviewsPerSession_results_100p = (get_results_value(get_gaData(service, gaid_100p, start_date_combine, 
        end_date_combine, 'ga:screenviewsPerSession')))
    avgSessionDuration_results_100p = (get_results_value(get_gaData(service, gaid_100p, start_date_combine, 
        end_date_combine, 'ga:avgSessionDuration')))
    mau_results_100p = (get_results_mau(get_gaMau(service, gaid_100p, end_date_combine, end_date_combine, 
        'ga:30dayUsers', 'ga:date')))
    mau_new_mechenism_results_100p = get_results_mau(get_segmentMau(service, gaid_100p, end_date_combine, 
        end_date_combine, 'ga:30dayUsers', 'ga:date', segment_new_mechanism))

    sessions_results_10p = (get_results_value(get_gaData(service, gaid_10p, start_date_combine, 
        end_date_combine, 'ga:sessions')))
    screenviews_results_10p = (get_results_value(get_gaData(service, gaid_10p, start_date_combine, 
        end_date_combine, 'ga:screenviews')))
    screenviewsPerSession_results_10p = (get_results_value(get_gaData(service, gaid_10p, start_date_combine, 
        end_date_combine, 'ga:screenviewsPerSession')))
    avgSessionDuration_results_10p = (get_results_value(get_gaData(service, gaid_10p, start_date_combine, 
        end_date_combine, 'ga:avgSessionDuration')))
    mau_results_10p = (get_results_mau(get_gaMau(service, gaid_10p, end_date_combine, end_date_combine, 
        'ga:30dayUsers', 'ga:date')))
    mau_new_mechenism_results_10p = (get_results_mau(get_segmentMau(service, gaid_10p, end_date_combine, 
        end_date_combine, 'ga:30dayUsers', 'ga:date', segment_new_mechanism)))

    sample_rate = 10

    return ([time_str, int(sessions_results_100p), int(sessions_results_10p) * sample_rate,
        int(screenviews_results_100p), int(screenviews_results_10p) * sample_rate, 
        float(screenviewsPerSession_results_100p), float(screenviewsPerSession_results_10p),
        float(avgSessionDuration_results_100p), float(avgSessionDuration_results_10p),
        int(mau_results_100p), int(mau_results_10p) * sample_rate,
        int(mau_new_mechenism_results_100p), int(mau_new_mechenism_results_10p) * sample_rate])

def getSecondDecimalPlace(number):
    return float('{:.2f}'.format(number))

def storage2xls(yearAndmonth, data_index, segment_new_mechanism):
    print('storage2xls: {0}'.format(yearAndmonth))
    year_str = str(yearAndmonth)[0:4]
    month_str = str(yearAndmonth)[4:6]    
    if (str(yearAndmonth)[4:5] == 0):
        month_str = str(yearAndmonth)[5:6]
    else:
        month_str = str(yearAndmonth)[4:6]

    year = int(year_str)
    month = int(month_str)
    print('year: {0}, month: {1}'.format(year, month))

    cwd = os.getcwd()
    file_path = cwd + '\\' + 'GA_analytics_result.xls'
    rb = open_workbook(file_path, formatting_info=True)
    r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
    wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
    w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy

    ga_results = getDataPermonth(year, month, segment_new_mechanism)

    #write Sessions data
    print('write Sessions data')
    sessions_index = 0
    sessions_100p = float(ga_results[1])
    sessions_10p = float(ga_results[2])
    w_sheet.write(sessions_index, data_index, ga_results[0])
    w_sheet.write(sessions_index + 1, data_index, sessions_100p)
    w_sheet.write(sessions_index + 2, data_index, sessions_10p)
    w_sheet.write(sessions_index + 3, data_index, sessions_100p + sessions_10p)

    #write Screenviews data
    print('write Screenviews data')
    screenviews_index = sessions_index + 5
    screenviews_100p = float(ga_results[3])
    screenviews_10p = float(ga_results[4])
    w_sheet.write(screenviews_index, data_index, ga_results[0])
    w_sheet.write(screenviews_index + 1, data_index, screenviews_100p)
    w_sheet.write(screenviews_index + 2, data_index, screenviews_10p)
    w_sheet.write(screenviews_index + 3, data_index, screenviews_100p + screenviews_10p)

    #write ScreenviewsPerSession data
    print('write ScreenviewsPerSession data')
    screenviewsPerSession_index = screenviews_index + 5
    screenviewsPerSession_100p = float(ga_results[5])
    screenviewsPerSession_10p = float(ga_results[6])
    w_sheet.write(screenviewsPerSession_index, data_index, ga_results[0])
    w_sheet.write(screenviewsPerSession_index + 1, data_index, getSecondDecimalPlace(screenviewsPerSession_100p))
    w_sheet.write(screenviewsPerSession_index + 2, data_index, getSecondDecimalPlace(screenviewsPerSession_10p))

    #write AvgSessionDuration data
    print('write AvgSessionDuration data')
    avgSessionDuration_index = screenviewsPerSession_index + 4
    avgSessionDuration_100p = float(ga_results[7])
    avgSessionDuration_10p = float(ga_results[8])
    w_sheet.write(avgSessionDuration_index, data_index, ga_results[0])
    w_sheet.write(avgSessionDuration_index + 1, data_index, getSecondDecimalPlace(avgSessionDuration_100p))
    w_sheet.write(avgSessionDuration_index + 2, data_index, getSecondDecimalPlace(avgSessionDuration_10p))

    #write MAU data
    print('write MAU data')
    mau_index = avgSessionDuration_index + 4
    mau_100p = float(ga_results[9])
    mau_10p = float(ga_results[10])
    w_sheet.write(mau_index, data_index, ga_results[0])
    w_sheet.write(mau_index + 1, data_index, mau_100p)
    w_sheet.write(mau_index + 2, data_index, mau_10p)
    w_sheet.write(mau_index + 3, data_index, mau_100p + mau_10p)

    #write MAU new mechenism data
    print('write MAU new mechenism data')
    mau_new_mechenism_index = mau_index + 5
    mau_new_mechenism_100p = float(ga_results[11])
    mau_new_mechenism_10p = float(ga_results[12])
    w_sheet.write(mau_new_mechenism_index, data_index, ga_results[0])
    w_sheet.write(mau_new_mechenism_index + 1, data_index, mau_new_mechenism_100p)
    w_sheet.write(mau_new_mechenism_index + 2, data_index, mau_new_mechenism_10p)
    w_sheet.write(mau_new_mechenism_index + 3, data_index, mau_new_mechenism_100p + mau_new_mechenism_10p)

    wb.save('GA_analytics_result.xls')

def prepare_data_title():
    log_info = "Prepare ga data: "
    print(log_info)

    mau_title = 'GA MAU'
    mau_new_mechenism_title = 'GA new mechenism MAU'
    sessions_title = 'GA Sessions'
    screenviews_title = 'GA Screenviews'
    screenviewsPerSession_title = 'GA ScreenviewsPerSession'
    avgSessionDuration_title = 'GA AvgSessionDuration'
    total_title = 'Total'

    mau_str = '30dayUsers'
    sessions_str = 'sessions'
    screenviews_str = 'screenviews'
    screenviewsPerSession_str = 'screenviewsPerSession'
    avgSessionDuration_str = 'avgSessionDuration'  

    gaid_100p_str = 'ASUS Themes' 
    gaid_10p_str = 'ASUS Themes user activity (10%)'

    wb_CDNDataArrangeTotal = xlwt.Workbook()
    ws_CDNDataPerFile = wb_CDNDataArrangeTotal.add_sheet('GA data', cell_overwrite_ok=True)

    #write Sessions title
    sessions_index = 0
    ws_CDNDataPerFile.write(sessions_index, 0, sessions_title)
    ws_CDNDataPerFile.write(sessions_index + 1, 0, gaid_100p_str)
    ws_CDNDataPerFile.write(sessions_index + 2, 0, gaid_10p_str)
    ws_CDNDataPerFile.write(sessions_index + 3, 0, total_title)

    #write Screenviews title
    screenviews_index = sessions_index + 5
    ws_CDNDataPerFile.write(screenviews_index, 0, screenviews_title)
    ws_CDNDataPerFile.write(screenviews_index + 1, 0, gaid_100p_str)
    ws_CDNDataPerFile.write(screenviews_index + 2, 0, gaid_10p_str)
    ws_CDNDataPerFile.write(screenviews_index + 3, 0, total_title)

    #write ScreenviewsPerSession title
    screenviewsPerSession_index = screenviews_index + 5
    ws_CDNDataPerFile.write(screenviewsPerSession_index, 0, screenviewsPerSession_title)
    ws_CDNDataPerFile.write(screenviewsPerSession_index + 1, 0, gaid_100p_str)
    ws_CDNDataPerFile.write(screenviewsPerSession_index + 2, 0, gaid_10p_str)

    #write AvgSessionDuration title
    avgSessionDuration_index = screenviewsPerSession_index + 4
    ws_CDNDataPerFile.write(avgSessionDuration_index, 0, avgSessionDuration_title)
    ws_CDNDataPerFile.write(avgSessionDuration_index + 1, 0, gaid_100p_str)
    ws_CDNDataPerFile.write(avgSessionDuration_index + 2, 0, gaid_10p_str)    

    #write MAU data title
    mau_index = avgSessionDuration_index + 4
    ws_CDNDataPerFile.write(mau_index, 0, mau_title)
    ws_CDNDataPerFile.write(mau_index + 1, 0, gaid_100p_str)
    ws_CDNDataPerFile.write(mau_index + 2, 0, gaid_10p_str)
    ws_CDNDataPerFile.write(mau_index + 3, 0, total_title)

    #write MAU new mechenism data title
    mau_new_mechenism_index = mau_index + 5
    ws_CDNDataPerFile.write(mau_new_mechenism_index, 0, mau_new_mechenism_title)
    ws_CDNDataPerFile.write(mau_new_mechenism_index + 1, 0, gaid_100p_str)
    ws_CDNDataPerFile.write(mau_new_mechenism_index + 2, 0, gaid_10p_str)
    ws_CDNDataPerFile.write(mau_new_mechenism_index + 3, 0, total_title)

    wb_CDNDataArrangeTotal.save('GA_analytics_result.xls')

def storage_gadata(analytics_month_list, segment_new_mechanism):
    prepare_data_title()

    data_index = 1
    for yearAndmonth in analytics_month_list:
        storage2xls(yearAndmonth, data_index, segment_new_mechanism)
        data_index = data_index + 1

def getJsonFileInfo():
    with open('ga_analytics_setting.json') as data_file:
        data = json.load(data_file)
    return [data['analytics_month_list'], data['appVersion']]

def main():
    print(getJsonFileInfo()[0])
    print(getJsonFileInfo()[1])
    analytics_month_list = getJsonFileInfo()[0]
    segment_conditon_pre = 'sessions::condition::ga:appVersion[]'
    json_compress_mechenism_appVersion = getJsonFileInfo()[1]
    segment_new_mechanism = segment_conditon_pre + json_compress_mechenism_appVersion    
    storage_gadata(analytics_month_list, segment_new_mechanism)

if __name__ == '__main__':
    main()
