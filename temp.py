from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials

gauth = GoogleAuth()
scope = ["https://www.googleapis.com/auth/drive"]
gauth.credentials = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
drive = GoogleDrive(gauth)


file_id = "1wkYF5yZZnEGP_AzticyNndGizKTz_lvv"


file = drive.CreateFile({'id': file_id})
file.GetContentFile('data1.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

