import requests
import json
from datetime import datetime as dt
from datetime import timedelta as td

class PowerBIClient:
    
    def __init__(self, client_key, client_secret, workspace_id, tenant_id):
        self.client_key = client_key
        self.client_secret = client_secret
        self.workspace_id = workspace_id
        self.tenant_id = tenant_id
        self.token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        self.api_url = "https://api.powerbi.com/v1.0/myorg"
        self.access_token = self.get_access_token()

# -----------------------------------------------------------------------------------------------------------------------------------------
    
    def get_access_token(self):
        
        """
            Generate access token using client key and client secret.
        """
        
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        payload = {
            'grant_type': 'client_credentials',
            'client_id': self.client_key,
            'client_secret': self.client_secret,
            'scope': 'https://analysis.windows.net/powerbi/api/.default'
        }

        response = requests.post(self.token_url, headers=headers, data=payload)
        if response.status_code == 200:
            token_info = response.json()
            return token_info.get('access_token')
            
        else:
            raise Exception(f"Failed to retrieve access token: {response.text}")

# -----------------------------------------------------------------------------------------------------------------------------------------

    def get_reports(self, short = True):
        
        """
            Get the reports from the specified Power BI workspace.
            SHORT: return full JSON reponse or parsed (shortened) version.
        """
        
        headers = {
            'Authorization': f'Bearer {self.access_token}'
        }

        reports_url = f"{self.api_url}/groups/{self.workspace_id}/reports"
        response = requests.get(reports_url, headers=headers)
        
        if response.status_code == 200:
            if short:
                # filter out Usage metrics reports (hidden reports in a workspace)
                return dict(
                        filter(
                            lambda item: item[0] not in 
                                                ['Report Usage Metrics Report','Usage Metrics Report', 'Reports usage metrics (UM)']
                            , {report['name']: report['id'] for report in response.json().get('value', [])}.items()))
                
            else:
                return response.json()
                
        else:
            return (f"Failed to retrieve reports: {response.text}")

# -----------------------------------------------------------------------------------------------------------------------------------------
    
    def get_datasets(self, short = True):
        
        """
            Get all datasets from the specified Power BI workspace.
            SHORT: return full JSON reponse or parsed (shortened) version.
        """
        
        headers = {
            'Authorization': f'Bearer {self.access_token}'
        }

        datasets_url = f"{self.api_url}/groups/{self.workspace_id}/datasets"
        response = requests.get(datasets_url, headers=headers)
        responseJSON = response.json()

        if response.status_code == 200:
            if short:
                # filter out Usage metrics models (hidden reports in a workspace)
                return dict(
                        filter(
                            lambda item: item[0] not in 
                                                ['Report Usage Metrics Model', 'Usage Metrics Report']
                            ,{report['name']: report['id'] for report in response.json().get('value', [])}.items()))
                
            else:
                return responseJSON
                
        else:
            return f"Failed to retrieve datasets: {response.text}"

# -----------------------------------------------------------------------------------------------------------------------------------------

    def get_report_users(self, report_id, short = True):
        
        """
            Retrieve a list of users who have access to a specific Power BI report.
            SHORT: returns short format (parsed JSON response). 
        """
        
        url = f'https://api.powerbi.com/v1.0/myorg/admin/reports/{report_id}/users'
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
    
        response = requests.get(url, headers=headers)
    
        if response.status_code == 200:

            users = response.json().get('value', [])
            
            if short:
                
                usersShort = []
                
                for user in users:

                    item = {
                            'name': user.get('displayName')
                            ,'email': user.get('emailAddress')
                            ,'rights': user.get('appUserAccessRight')
                    }

                    usersShort.append(item)

                return usersShort
                
            else:
                return users
            
        else:
            return f"Failed to retrieve report users: {response.status_code} - {response.text}"

# -----------------------------------------------------------------------------------------------------------------------------------------

    def get_app_users(self, appId, short=True, file=False):
        """
            Returns a list of users that have access to the specified app.
            SHORT: returns short format (parsed JSON response).
            FILE: if True, writes the results to a file named '__app_users__.txt'.
        """
    
        url = f'https://api.powerbi.com/v1.0/myorg/admin/apps/{appId}/users'
    
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
    
        response = requests.get(url, headers=headers)
    
        if response.status_code == 200:
    
            users = response.json().get('value', [])
    
            if short:
    
                usersShort = []
    
                for user in users:
    
                    item = {
                        'name': user.get('displayName'),
                        'email': user.get('emailAddress'),
                        'rights': user.get('appUserAccessRight')
                    }
    
                    usersShort.append(item)
    
                if file:
                    with open('__app_users__.txt', 'w') as f:
                        f.write("name|email|rights\n")  # Add CSV header
                        for user in usersShort:
                            f.write(f"{user['name']}|{user['email']}|{user['rights']}\n")
                    return f"Results written to __app_users__.txt"

                return usersShort
    
            else:
                if file:
                    with open('__app_users__.txt', 'w') as f:
                        f.write(response.text)
                    return f"Results written to __app_users__.txt"
    
                return users
    
        else:
            return f"Failed to retrieve report users: {response.status_code} - {response.text}"

# -----------------------------------------------------------------------------------------------------------------------------------------

    def get_schedule(self, datasetId = None):
       
        """
            Get refresh schedule of a single report.
        """
        
        headers = {
            'Authorization': f'Bearer {self.access_token}'
        }
            
        datasets_url = f"{self.api_url}/groups/{self.workspace_id}/datasets/{datasetId}/refreshSchedule"
        response = requests.get(datasets_url, headers=headers)

        if response.status_code == 200:
            _res = response.json()
            return {
                            'datasetID': datasetId
                            ,'active': _res['enabled']
                            ,'days': _res['days']
                            ,'time': _res['times']
                            ,'timeZone': _res['localTimeZoneId']
                    }
        else:
            raise Exception(f"Failed to retrieve datasets: {response.text}")

# -----------------------------------------------------------------------------------------------------------------------------------------

    def get_refreshInfo(self, datasetID = None, short = True):
        
        """
            Retreive last record from history of updates (specified dataset).
            Retreives only ONE last update from the history.
            SHORT: returns short format (parsed JSON response).            
        """
        headers = {
            'Authorization': f'Bearer {self.access_token}'
        }

        reports_url = f"https://api.powerbi.com/v1.0/myorg/groups/{self.workspace_id}/datasets/{datasetID}/refreshes?$top={1}"
        response = requests.get(reports_url, headers=headers)
        responseJSON = response.json()
        startTime = responseJSON['value'][0].get('startTime')
        endTime = responseJSON['value'][0].get('endTime')
        status = responseJSON['value'][0].get('status')
        type = responseJSON['value'][0].get('refreshType')
        
        if response.status_code == 200 and endTime != None:
            if short:
                return {
                        'dataset': datasetID
                        ,'start': startTime
                        ,'end': endTime
                        ,'span': self.time(dt.strptime(endTime.split('.')[0], '%Y-%m-%dT%H:%M:%S') - dt.strptime(startTime.split('.')[0], '%Y-%m-%dT%H:%M:%S'))
                        ,'status': status
                        ,'type': type
                        }
            else:
                return responseJSON

        elif response.status_code == 200 and endTime == None:
            return f'Dataset is being refreshed'
            
        else:
            return Exception(f"Failed to retrieve refresh info: {response.text}")

# -----------------------------------------------------------------------------------------------------------------------------------------
        
    def get_refreshInfoAll(self, short = True):
        
        """
            Retreive refresh info of all datasets.
            Retreives only ONE last update from the history of each dataset.
            SHORT: returns short format (parsed JSON response).            
        """
        
        headers = {
            'Authorization': f'Bearer {self.access_token}'
        }

        # get all datasets refresh info

        datasetsAll = self.get_datasets()
        resultAll = []

        # loop datasets
        for _dataset in datasetsAll.values():

            reports_url = f"https://api.powerbi.com/v1.0/myorg/groups/{self.workspace_id}/datasets/{_dataset}/refreshes?$top={1}"
            response = requests.get(reports_url, headers=headers)
            responseJSON = response.json()
            startTime = responseJSON['value'][0].get('startTime')
            endTime = responseJSON['value'][0].get('endTime')
            status = responseJSON['value'][0].get('status')
            type = responseJSON['value'][0].get('refreshType')
            
            if response.status_code == 200 and endTime != None:
                if short:
                    resultAll.append({
                        'dataset': _dataset
                        ,'start': startTime
                        ,'end': endTime
                        ,'span': self.time(dt.strptime(endTime.split('.')[0], '%Y-%m-%dT%H:%M:%S') - dt.strptime(startTime.split('.')[0], '%Y-%m-%dT%H:%M:%S'))
                        ,'status': status
                        ,'type': type
                                })
                else:
                    resultAll.append(response.json())
                    
            elif response.status_code == 200 and endTime == None:
                resultAll.append({
                        'dataset': _dataset
                        ,'start': ''
                        ,'end': ''
                        ,'span': ''
                        ,'status': f'Dataset is being refreshed'
                        ,'type': ''
                                })
                
            else:
                return Exception(f"Failed to retrieve refresh info: {response.text}")
        
        return resultAll

# -----------------------------------------------------------------------------------------------------------------------------------------

    def refresh_dataset(self, dataset_id):
        
        """Refresh specified dataset in the Power BI workspace."""
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }

        refresh_url = f"{self.api_url}/groups/{self.workspace_id}/datasets/{dataset_id}/refreshes"
        response = requests.post(refresh_url, headers=headers)

        if response.status_code == 202:
            return f"Dataset {dataset_id} refresh started."
            
        else:
            return f"Failed to start dataset refresh: {response.text}"

# -----------------------------------------------------------------------------------------------------------------------------------------

    def query_dataset(self, dataset_id, query = None, file = False):
        
        """
            Queries specified dataset in Power BI workspace.
            Response from API either single value or table will be presented in a table view ('|' bar separated CSV ready).

            If file == True: result will be written in a file (__resp.txt).

            Otherwise returns a list with two items: 
                * list[0] - headers
                * list[1] - values              
                    
        """
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }

        body = {
                    "queries": [
                        {
                            "query": query
                        }
                    ]
                }

        refresh_url = f"{self.api_url}/groups/{self.workspace_id}/datasets/{dataset_id}/executeQueries"
        
        response = requests.post(refresh_url, headers = headers, json = body)

        # check response code
        if response.status_code == 200:
            
            respJson = response.json()

            # check if response was OK and DAX query was OK
            if respJson['results'][0].get('error') == None:

                csvHead = '|'.join(respJson['results'][0]['tables'][0]['rows'][0].keys())
                csvBody = ['|'.join(str(value) for value in line.values()) + '\n' for line in respJson['results'][0]['tables'][0]['rows']]

                # return list
                if not file:
                    return [csvHead + '\n', csvBody]

                # write to file
                else:
                    self.write_response([csvHead + '\n', csvBody])
                    return 'file {__resp.txt} was created' 

            else:
                return f"Query failed. {respJson['results'][0]['error']['code']} : {respJson['results'][0]['error']['message']}"
            
        else:
            return f"Failed to query dataset: {response.text}"

# -----------------------------------------------------------------------------------------------------------------------------------------

    def time(self, delta):
        
        '''
            Extract human readable time delta from timedelta
        '''
        
        total_seconds = int(delta.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        
        return f'{hours:02}:{minutes:02}:{seconds:02}'

# -----------------------------------------------------------------------------------------------------------------------------------------

    def write_response(self, txt):

        """
            Write DAX query response into file.
        """

        with open ('__resp.txt', 'w', encoding='utf-8') as _file:
            
            _file.write(txt[0])
            
            for line in txt[1]:
                _file.write(line)
