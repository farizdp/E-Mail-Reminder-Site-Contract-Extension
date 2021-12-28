import smtplib, gspread, requests, datetime, time, socket, os, json
from google.oauth2.service_account import Credentials
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']
credentials = Credentials.from_service_account_file('creds.json', scopes = scope)
client = gspread.authorize(credentials)

email = json.load('user.json')
MY_ADDRESS = email.get('email')
PASSWORD = email.get('password')

# Data proxy if needed
socket.getaddrinfo('YourPROXY', YourPort)
proxy = 'yourproxy:yourport'
os.environ['http_proxy'] = proxy
os.environ['https_proxy'] = proxy
os.environ['HTTP_PROXY'] = proxy
os.environ['HTTPS_PROXY'] = proxy

def get_data_agata(bulan, ns = 'Padang', reg = 'Sumbagteng'):
    web = 'http://WEB AGATA/load_data_summary/panel_id_data_nos_network_profile_site?splitData=&filter[0][field]=tower_provider_name&filter[0][data][type]=string&filter[0][data][value]=Non%20TP&filter[1][field]=owner_tower_name&filter[1][data][type]=string&filter[1][data][value]=TELKOMSEL&filter[2][field]=akhir_periode_kontrak&filter[2][data][type]=string&filter[2][data][value]=' + bulan + '&filter[3][field]=regional_name&filter[3][data][type]=string&filter[3][data][value]=' + reg + '&filter[4][field]=departement_name&filter[4][data][type]=string&filter[4][data][value]=' + ns + '&page=1&start=0&limit=1000'
    agata_json = requests.get(web).text
    site = agata_json.replace("success:", "\"success\":")
    site = agata_json.replace("jml:", "\"jml\":")
    site = agata_json.replace("data:", "\"data\":")
    return site['data']

def get_list_user(rtpo) :
    sheet = client.open('KARYAWAN').worksheet('Database')
    cell_rtpo = sheet.find(rtpo)
    list_rtpo = sheet.cell(cell_rtpo.row, 2).value
    cell_sm = sheet.find('SITE MANAGEMENT')
    list_sm = sheet.cell(cell_sm.row, 2).value
    cell_mgr_ns = sheet.find('MANAGER NS')
    list_mgr_ns = sheet.cell(cell_mgr_ns.row, 2).value
    cell_mgr_sm = sheet.find('MANAGER SM')
    list_mgr_sm = sheet.cell(cell_mgr_sm.row, 2).value
    # cell_admin = sheet.find('ADMIN')
    # list_admin = sheet.cell(cell_mgr_sm.row, 2).value
    output = [list_rtpo, list_sm, list_mgr_ns, list_mgr_sm]
    return output

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def kirim_email(site, rtpo, h_minus):
    output = get_list_user(rtpo)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(MY_ADDRESS, PASSWORD) 
    message_template = read_template('message.txt')
    message = message_template.substitute(
        PERSON_NAME = 'Team Site Management',
        SITE_ID = site['site_id'],
        SITE_NAME = site['site_name'],
        ALAMAT_SITE = site['alamat'],
        TGL_KONTRAK = site['akhir_periode_kontrak'],
        LAND_LORD = site['land_lord'],
        NO_KONTRAK = site['no_kontrak'],
        RTPO = site['technical_area_name'],
        )
    msg = MIMEMultipart()
    msg['From'] = MY_ADDRESS
    msg['To'] = output[1]
    msg['Cc'] = output[3] + ', ' + output[2] + ', ' + output[0]
    msg['Subject'] = '[' + h_minus + '] Perpanjangan Kontrak Site ' + site['site_id'] + ' ' + site['site_name']
    msg.attach(MIMEText(message, 'html'))
    server.send_message(msg)
    del msg
    server.quit()
    
def main():
    hari = datetime.datetime.now()
    while 1:
        if hari.strftime('%d') == '01' and hari.strftime('%H:%M') == '09:00':
            waktu_6bulan = hari + datetime.timedelta(days = 150)
            site_6bulan = get_data_agata(waktu_6bulan.strftime('%Y-%m'))
            for x in site_6bulan :
                kirim_email(x, x['technical_area_name'], 'H-6 Bulan')
            
            waktu_3bulan = hari + datetime.timedelta(days = 60)
            site_3bulan = get_data_agata(waktu_3bulan.strftime('%Y-%m'))
            for x in site_3bulan :
                kirim_email(x, x['technical_area_name'], 'H-3 Bulan')

            waktu_1bulan = hari + datetime.timedelta(days = 30)
            site_1bulan = get_data_agata(waktu_1bulan.strftime('%Y-%m'))
            for x in site_1bulan :
                kirim_email(x, x['technical_area_name'], 'H-1 Bulan')

if __name__ == '__main__':
    main()