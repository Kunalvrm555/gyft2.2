import argparse
import getpass
import re
from bs4 import BeautifulSoup as bs
import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
import pandas as pd
import io
# Parsing from commmand line
parser = argparse.ArgumentParser()
parser.add_argument("-u", "--user", help="ERP Username/Login ID")
args = parser.parse_args()
if args.user is None:
    args.user = input("Enter you Roll Number: ")
erp_password = getpass.getpass("Enter your ERP password: ")

# Parsing ends

ERP_HOMEPAGE_URL = 'https://erp.iitkgp.ac.in/IIT_ERP3/'
ERP_LOGIN_URL = 'https://erp.iitkgp.ac.in/SSOAdministration/auth.htm'
ERP_SECRET_QUESTION_URL = 'https://erp.iitkgp.ac.in/SSOAdministration/getSecurityQues.htm'


headers = {
    'timeout': '20',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Ubuntu Chromium/51.0.2704.79 Chrome/51.0.2704.79 Safari/537.36',
}

s = requests.Session()
r = s.get(ERP_HOMEPAGE_URL)
soup = bs(r.text, 'html.parser')
sessionToken = soup.find_all(id='sessionToken')[0].attrs['value']
r = s.post(ERP_SECRET_QUESTION_URL, data={'user_id': args.user},
           headers=headers)
secret_question = r.text
print("Your secret question: " + secret_question)
secret_answer = getpass.getpass("Enter the answer to the security question: ")
login_details = {
    'user_id': args.user,
    'password': erp_password,
    'answer': secret_answer,
    'sessionToken': sessionToken,
    'requestedUrl': 'https://erp.iitkgp.ac.in/IIT_ERP3',
}
r = s.post(ERP_LOGIN_URL, data=login_details,
           headers=headers)
try:
    ssoToken = re.search(r'\?ssoToken=(.+)$',
                         r.history[1].headers['Location']).group(1)
except IndexError:
    print("Error: Please make sure the entered credentials are correct!")

ERP_TIMETABLE_URL = "https://erp.iitkgp.ac.in/Acad/student/view_stud_time_table.jsp"

timetable_details = {
    'ssoToken': ssoToken,
    'module_id': '16',
    'menu_id': '40',
}
# This is just a hack to get cookies. TODO: do the standard thing here
abc = s.post('https://erp.iitkgp.ac.in/Acad/student/view_stud_time_table.jsp',
             headers=headers, data=timetable_details)
cookie_val = None
for a in s.cookies:
    if (a.path == "/Acad/"):
        cookie_val = a.value

cookie = {
    'JSESSIONID': cookie_val,
}
r = s.post('https://erp.iitkgp.ac.in/Acad/student/view_stud_time_table.jsp',
           cookies=cookie, headers=headers, data=timetable_details)

print(r.status_code)
# soup = bs(r.text, 'html.parser')
# table_MN = pd.read_html(soup)
# print(f'Total tables: {len(table_MN)}')
table=pd.read_json(io.StringIO(r.text))
print(f'Total tables: {len(table)}')