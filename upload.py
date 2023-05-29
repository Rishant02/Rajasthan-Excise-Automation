import requests
import json
from dotenv import load_dotenv
import os
load_dotenv()

def upload_root_file(file_name):
    # Define the Azure AD application credentials
    client_id = os.getenv('API_CLIENT_ID')
    tenant_id = os.getenv('API_TENANT_ID')

    auth_url=f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    api_url = 'https://graph.microsoft.com/v1.0'
    auth_data={
        'grant_type':'password',
        'client_id': client_id,
        'username':os.getenv('API_EMAIL_ADDRESS'),
        'password':os.getenv('API_PASSWORD'),
        'scope': 'Files.ReadWrite.All',
    }
    auth_header={'Content-Type': 'application/x-www-form-urlencoded'}

    r=requests.post(auth_url,data=auth_data,headers=auth_header)
    results=r.json()

    graph_headers={'Authorization':'Bearer {}'.format(results['access_token']),'Content-Type': 'application/json'}
    

    # one_drive_destination='https://graph.microsoft.com/v1.0/me/drive/items/root:/MIS%20STATE/RSBCL/'
    one_drive_destination='https://graph.microsoft.com/v1.0/me/drive/items/root:/MIS%20STATE/'
    file_res = requests.get(one_drive_destination+file_name,headers=graph_headers)
    
    if file_res.status_code == 200:
        item_id = file_res.json()['id']
        url= f'{api_url}/me/drive/items/{item_id}'
        del_res = requests.delete(url,headers=graph_headers)
        if del_res.status_code != 204:
            print('Error deleting file:',del_res.text)
        print(f'{file_name} is deleted. Now uploading...')
    file_path=os.path.join(os.getcwd(),file_name)
    file_size=os.stat(file_path).st_size
    file_data=open(file_path,'rb')
    res={}
    if file_size < 4100000:
        r=requests.put(one_drive_destination+file_name+':/content',data=file_data,headers=graph_headers)
        if r.status_code==201:
            file_data.close()
            print(f'{file_name} has been successfully uploaded.')
            os.remove(file_path)
            res=r.json()
    else:
        upload_session=requests.post(one_drive_destination+file_name+":/createUploadSession",headers=graph_headers).json()
        with open(file_path,'rb') as f:
            total_file_size=os.path.getsize(file_path)
            chunk_size=3932160
            chunk_number=total_file_size//chunk_size
            chunk_leftover=total_file_size-chunk_size*chunk_number
            i=0
            while True:
                chunk_data=f.read(chunk_size)
                start_index=i*chunk_size
                end_index=start_index+chunk_size

                if not chunk_data:
                    break
                if i==chunk_number:
                    end_index=start_index+chunk_leftover

                headers={'Content-Length':'{}'.format(chunk_size),'Content-Range':'bytes {}-{}/{}'.format(start_index,end_index-1,total_file_size)}

                chunk_data_upload=requests.put(upload_session['uploadUrl'],data=chunk_data,headers=headers)
                print(chunk_data_upload)
                print(chunk_data_upload.json())
                res=chunk_data_upload.json()
                i+=1
        file_data.close()
        os.remove(file_path)
    if res:
        request_data={
            'type':'view'
        }
        link_res=requests.post('https://graph.microsoft.com/v1.0/me/drive/items/{}/createLink'.format(res['id']),headers=graph_headers,data=json.dumps(request_data))
        data=link_res.json()
        return data['link']['webUrl']
    else:
        return None
