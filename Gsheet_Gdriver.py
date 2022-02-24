import pygsheets
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import pandas as pd


def access_google_sheets(path, read_type):

    # # ------------------ 讀取google sheet ------------------
    final_sh_data_list = []
    sh_name_set = set()
    # gc = pygsheets.authorize(service_account_file=os.path.join('.', 'qantumbiology-2.json'))
    gc = pygsheets.authorize(client_secret=os.path.join('.', 'client_secrets.json'))

    sh = gc.open_by_url(path)
    sh.worksheets()
    for sheet in sh.worksheets():
        ws = sh.worksheet_by_title(sheet.title)
        df = ws.get_as_df(start='A1', index_colum=0, empty_value='', include_tailing_empty=False, has_header=False)
        df = df.fillna('')
        sh_data_list = df.values.tolist()
        # # read data set
        if read_type == 'data_set':
            sh_name_set.add(sheet.title)
            for sh_d in sh_data_list:
                final_sh_data_list += sh_d
        # # read signal pathway cascades
        if read_type == 'signal_pathway':
            for sig_cas in sh_data_list:
                new_sig_cas = {}
                for signal in sig_cas:
                    if signal:
                        if signal not in new_sig_cas:
                            new_sig_cas[signal] = False
                if new_sig_cas not in final_sh_data_list:
                    final_sh_data_list.append({f"{sheet.title}": new_sig_cas})
                    sh_name_set.add(sheet.title)
    # # ------------------ 讀取google sheet ------------------

    return final_sh_data_list, sh_name_set


def upload(xlsx_file_path):
    gauth = GoogleAuth()
    # Try to load saved client credentials
    gauth.LoadCredentialsFile(os.path.abspath(os.path.join('.', 'mycreds.txt')))
    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()

    # Save the current credentials to a file
    gauth.SaveCredentialsFile(os.path.abspath(os.path.join('.', 'mycreds.txt')))

    drive = GoogleDrive(gauth)
    # # get latest file
    path_list = glob.glob((os.path.abspath(os.path.join('.', 'exported_data', '*.xlsx'))))
    latest_file = sorted(path_list, reverse=True)[0]

    folder_list = drive.ListFile({'q': "'1heqmUz0KeXrALuWOAhd08IHPFVXSbIgk' in parents and trashed=False"}).GetList()

    today = datetime.datetime.now().date()

    file_info_list = []
    if folder_list:
        for gdf in folder_list:
            # # get the info of the file with same name for further replace
            try:
                gdf_file_date = re.search('(20\d{2})(\d{2})(\d{2})', gdf['title']).group()
                if gdf_file_date:
                    gdf_file_date = datetime.datetime.strptime(gdf_file_date, '%Y%m%d').date()
                    if (today - gdf_file_date).days == 0:
                        file_info_list.append(dict(
                            title=gdf['title'], id=gdf['id']))

                    if (today - gdf_file_date).days >= 10:
                        gdf.Delete()

            except Exception as e:
                print(e)

    # # if file name exist, renew file
    if file_info_list:
        for file_info in file_info_list:
            if file_info['title'] == os.path.basename(latest_file):
                file = drive.CreateFile(
                    {'parents': [{'id': '1heqmUz0KeXrALuWOAhd08IHPFVXSbIgk'}], 'title': file_info['title'],
                     'id': file_info['id']})
                file.SetContentFile(latest_file)
                # url = f"https://drive.google.com/thumbnail?id={img_info['id']}&sz=w1920-h1080"
                file.Upload()
                # print(file['title'], url)
                print(file['title'] + '---Renew and Upload success')

    # # upload new file
    else:
        file = drive.CreateFile(
            {'parents': [{'id': '1heqmUz0KeXrALuWOAhd08IHPFVXSbIgk'}], 'title': os.path.basename(xlsx_file_path)})
        file.SetContentFile(xlsx_file_path)
        file.Upload()
        print(file['title'] + '---Upload success')
