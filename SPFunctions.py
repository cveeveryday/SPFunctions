import requests

def get_token(client_id, client_secret, tenant_id):
    auth_url = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token".format(tenant_id)
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default',
    }
    
    response = requests.post(auth_url, headers=headers, data=data)
    
    if response.status_code == 200:
        return response.json()['access_token']
    else:
        raise Exception('Failed to retrieve token')


token = get_token(client_id, client_secret, tenant_id)

print(token)


def get_SPSite(token,siteName='default'):
    headers = {'Authorization': 'Bearer'+ token,
               'Content-Type': 'application/json'}
    response = requests.get('https://graph.microsoft.com/v1.0/sites/',headers = headers)
    if response.status_code == 200:
        siteList=response.json()['value']
        matchingSites=[]
        if siteName=='default':
            return siteList
        else:
            for site in siteList:
                if 'displayName' not in site:
                    continue
                if site['displayName'].lower()==siteName.lower():
                    matchingSites.append(site)
            return matchingSites
    else:
        raise Exception('Failed to retrieve site') 
    
    
print(get_SPSite(token,''))