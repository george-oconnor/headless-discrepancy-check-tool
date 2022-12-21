from sharepoint_stuff import getCTX, downloadFile
from email_stuff import unattended_send_email as send_email
from office365.sharepoint.client_context import ClientContext
from datetime import datetime
import pandas as pd
import keyring, os, shutil, math, requests, sys, json, logging

#version namimg year.type.major_version.minor_version - type{1 = headless, 0 = gui}
version = "22.1.0.1"
logList = []

logging.captureWarnings(True)
os.makedirs("./logs/", exist_ok=True)
logger = logging.getLogger(__name__)
handler = logging.FileHandler('./logs/headless_discrepancy_check.log')
formatter = logging.Formatter("%(asctime)s | %(name)s | %(levelname)s | %(message)s")
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.INFO)

def getAuth() -> tuple[ClientContext, str, str, str, str, str]:

    def checkCredential(credential_type:str, credential:str) -> None:
        if credential is None: logger.error(f"Failed to get {credential_type}")

    sharepoint_url = keyring.get_password("attendance_sharepoint", "url")
    username = keyring.get_password("attendance_sharepoint", "username")
    password = keyring.get_password("attendance_sharepoint", username)
    try:
        ctx = getCTX(sharepoint_url, username, password)
        logger.debug("Got ctx")
    except Exception as e:
        logger.error("Failed to get ctx")
        logger.error(e, exc_info=True)

    api_key = keyring.get_password("discrep_check", "students_api_key")
    try:
        api_url="https://ioe.isams.cloud/api/batch/1.0/json.ashx?apiKey=" + api_key
        logger.debug("Got api key")
    except Exception as e:
        logger.error("Failed to get api key")
        logger.error(e, exc_info=True)

    auth_server_url = keyring.get_password("discrep_check", "auth_server_url")
    client_id = keyring.get_password("discrep_check", "client_id")
    client_secret = keyring.get_password("discrep_check", "client_secret")

    checkCredential("sharepoint_url", sharepoint_url)
    checkCredential("auth_server_url", auth_server_url)
    checkCredential("client id", client_id)
    checkCredential("client_secret", client_secret)

    return ctx, api_url, auth_server_url, client_id, client_secret

def getDTEMSData(ctx:ClientContext, tempPath:str) -> list[dict]:
    student_li, student_dict_li = [], []
    dtems_filename = "dtems_details.csv"
    try:
        downloadFile(ctx, "/sites/AttendanceTest/Shared%20Documents/dtems_ft_data/FT_ClassDetails.csv", dtems_filename, tempPath)
        logger.debug("Got dtems datafile from sharepoint")
    except Exception as e:
        logger.error("Failed to get dtems datafile from sharepoint")
        logger.error(e, exc_info=True)
        return 0

    try:
        df = pd.read_csv(tempPath+dtems_filename, header=0, encoding="latin-1")
        logger.debug("Read dtems datafile into dataframe")
    except Exception as e:
        logger.error("Failed to read dtems datafile into dataframe")
        logger.error(e, exc_info=True)
    
    for i in range(0, len(df.index)):
        try:
            first_name = df.iloc[i, 6]
            last_name = df.iloc[i, 7]
            student_num = df.iloc[i, 8]
        except Exception as e:
            logger.error(f"Failed to get all data for student {first_name} {last_name} ({student_num}) on line {i}")
            logger.error(e, exc_info=True)
            return 0

        if math.isnan(student_num):
            logger.warning(f"No student number ({student_num}) for {first_name} {last_name} on line {i} so omitting")
        else:
            student_string = str(first_name) + "_" + str(last_name) + "_" + str(int(student_num))
            student_li.append(student_string)

    logger.debug(f"Found {len(student_li)} student details in dtems file")
    unique_student_li = list(dict.fromkeys(student_li))
    logger.info(f"Found {len(unique_student_li)} unique students in dtems file")

    unique_student_dict_li = []

    for student in unique_student_li:
        details = student.split('_')
        student_dict = {
            "firstname": details[0],
            "lastname": details[1],
            "student_id": details[2]
        }
        unique_student_dict_li.append(student_dict)

    logger.info(f"DTEMS: \t {len(unique_student_dict_li)}")
    return unique_student_dict_li

def get_new_token(auth_server_url:str, client_id:str, client_secret:str) -> str:
    """returns the isams access token as a string"""
    token_req_payload = {'grant_type': 'client_credentials'}
    token_response = requests.post(auth_server_url, data=token_req_payload, verify=False, allow_redirects=False, auth=(client_id, client_secret))

    if token_response.status_code !=200:
        logger.error("Failed to obtain token from the OAuth 2.0 server")
        #print(datetime.now().strftime("%H:%M:%S") + '  ' + "Failed to obtain token from the OAuth 2.0 server", file=sys.stderr)
        sys.exit(1)
    else:
        logger.info("Successfully obtained a new token")
        #print(datetime.now().strftime("%H:%M:%S") + '  ' + "Successfuly obtained a new token")
        tokens = json.loads(token_response.text)
        return tokens['access_token']

def callApi(auth_server_url:str, client_id:str, client_secret:str, api_url:str) -> str:
    """returns the api response text"""
    token = get_new_token(auth_server_url, client_id, client_secret)

    api_call_headers = {'Authorization': 'Bearer ' + token,
                        'accept': 'application/json'}
    
    api_call_response = requests.get(api_url, headers=api_call_headers, verify=False)
    
    logger.info("Response " + str(api_call_response.status_code))
    #print(datetime.now().strftime("%H:%M:%S") + '  ' + "Response " + str(api_call_response.status_code))

    if	api_call_response.status_code == 401:
        logger.error("Api response code 401")
    else:
        logger.info("Successful Response")
        #print(datetime.now().strftime("%H:%M:%S") + '  ' + "successful response")
    
    return api_call_response.text

def getListOfStudents(auth_server_url:str, client_id:str, client_secret:str, api_url:str) -> list[dict]:
    
    data = callApi(auth_server_url, client_id, client_secret, api_url)
    list_data = json.loads(data)
    
    export_list, export_dict_list = [], []
    pupil_list = list_data['iSAMS']['PupilManager']['CurrentPupils']['Pupil']
    for i in range(0, len(pupil_list)):
        #print(pupil_list[i]['SchoolId'] + '    ' + pupil_list[i]['NCYear'] + 'th    ' + pupil_list[i]['Surname'].strip() + ', ' + pupil_list[i]['Forename'].strip())
        export_list.append([pupil_list[i]['NCYear']+'th', pupil_list[i]['SchoolId'], pupil_list[i]['Surname'].strip(), pupil_list[i]['Forename'].strip()])
        export_dict_list.append({
            "firstname": pupil_list[i]['Forename'].strip(),
            "lastname": pupil_list[i]['Surname'].strip(),
            "student_id": pupil_list[i]['SchoolId']
        })
    
    logger.info('iSAMS: \t '+str(len(export_dict_list)))
    return export_dict_list

def compareData(dtems_li:list[dict], isams_li:list[dict]) -> list[list[str, str, str], list[str, str, str]]:
    dtems_id_set, isams_id_set = set(), set()
    for student in dtems_li:
        dtems_id_set.add(student["student_id"])

    for student in isams_li:
        isams_id_set.add(student["student_id"])

    on_isams_but_not_dtems = isams_id_set - dtems_id_set
    on_dtems_but_not_isams = dtems_id_set - isams_id_set
    #overall_difference = dtems_id_set.symmetric_difference(isams_id_set)

    print(on_isams_but_not_dtems, len(on_isams_but_not_dtems))
    print("")
    print(on_dtems_but_not_isams, len(on_dtems_but_not_isams))

    on_dtems_but_not_isams_li, on_isams_but_not_dtems_li = [], []

    for student in on_dtems_but_not_isams:
        res = next((sub for sub in dtems_li if sub['student_id'] == student), None)
        on_dtems_but_not_isams_li.append(res)

    for student in on_isams_but_not_dtems:
        res = next((sub for sub in isams_li if sub['student_id'] == student), None)
        on_isams_but_not_dtems_li.append(res)

    return on_dtems_but_not_isams_li, on_isams_but_not_dtems_li

def sendResultsEmail(on_dtems_but_not_isams_li, on_isams_but_not_dtems_li):
    body = f"""
    <html>
    <head></head>
    <body>
        <h2>Found {len(on_dtems_but_not_isams_li)} students on dtems but not on isams:</h2>
        <ul>
    """
    for student in on_dtems_but_not_isams_li:
        body += f"<li>  {student['firstname']} {student['lastname']} ({student['student_id']})</li>"
    
    body += f"""    
        </ul>
        <p>\n</p>
        <h2>Found {len(on_isams_but_not_dtems_li)} student on isams but not on dtems:</h2>
        <ul>
    """
    for student in on_isams_but_not_dtems_li:
        body += f"<li>  {student['firstname']} {student['lastname']} ({student['student_id']})</li>"

    body += """
        </ul>
    </body>
    </html>
    """

    username = keyring.get_password("attendance_sharepoint", "username")
    password = keyring.get_password("attendance_sharepoint", username)

    send_email("Discrepancy Check Results", body, "none", username, password, "goconnor@instituteofeducation.ie")


def main() -> None:
    ctx, api_url, auth_server_url, client_id, client_secret = getAuth()

    tempPath = "./.temp/"
    os.makedirs(tempPath, exist_ok=True)

    dtems_students = getDTEMSData(ctx, tempPath)
    isams_students = getListOfStudents(auth_server_url, client_id, client_secret, api_url)

    on_dtems_but_not_isams_li, on_isams_but_not_dtems_li = compareData(dtems_students, isams_students)

    sendResultsEmail(on_dtems_but_not_isams_li, on_isams_but_not_dtems_li)

    shutil.rmtree(tempPath)

if __name__ == "__main__":
    main()