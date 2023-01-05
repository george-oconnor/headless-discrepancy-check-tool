from sharepoint_stuff import getCTX, downloadFile
from ioe_email_stuff import send_email
from office365.sharepoint.client_context import ClientContext
from datetime import datetime, date
from time import time
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
        sys.exit()

    api_key = keyring.get_password("discrep_check", "students_api_key")
    try:
        api_url="https://ioe.isams.cloud/api/batch/1.0/json.ashx?apiKey=" + api_key
        logger.debug("Got api key")
    except Exception as e:
        logger.error("Failed to get api key")
        logger.error(e, exc_info=True)
        sys.exit()

    auth_server_url = keyring.get_password("discrep_check", "auth_server_url")
    client_id = keyring.get_password("discrep_check", "client_id")
    client_secret = keyring.get_password("discrep_check", "client_secret")

    checkCredential("auth_server_url", auth_server_url)
    checkCredential("client id", client_id)
    checkCredential("client_secret", client_secret)

    return ctx, api_url, auth_server_url, client_id, client_secret

def getDTEMSDataframe(ctx:ClientContext, tempPath:str, dtems_filename:str) -> pd.DataFrame:
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
        return df
    except Exception as e:
        logger.error("Failed to read dtems datafile into dataframe")
        logger.error(e, exc_info=True)
        return 0

def getDTEMSData(ctx:ClientContext, tempPath:str) -> list[dict]:
    student_li, student_dict_li = [], []

    df = getDTEMSDataframe(ctx, tempPath, "dtems_details.csv")

    #lunch_cls_grp_ids = [5774, 5839, 5743]
    lunch_cls_grp_ids = []

    for index, row in df.iterrows():
        if row[1] in lunch_cls_grp_ids:
            logger.debug(f"Removed row {index} : {row[1]} {row[2]} {row[3]} {row[4]} because it is a lunch class")
            df.drop(index, inplace=True)
        elif "nan" in [str(row[0]), str(row[1]), str(row[2]), str(row[6]), str(row[7]), str(row[8])]:
            logger.warning(f"Removed row {index} : {row[1]} {row[2]} {row[3]} {row[4]} because of empty fields")
            df.drop(index, inplace=True)
        else:
            try:
                first_name = row[6]
                last_name = row[7]
                student_num = row[8]
            except Exception as e:
                logger.error(f"Failed to get all data for student {first_name} {last_name} ({student_num}) on line {index}")
                logger.error(e, exc_info=True)
                return 0

            if math.isnan(student_num):
                logger.warning(f"No student number ({student_num}) for {first_name} {last_name} on line {index} so omitting")
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
    return unique_student_dict_li, df

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

def sendResultsEmail(on_dtems_but_not_isams_li:list[dict], on_isams_but_not_dtems_li:list[dict], dtems_df:pd.DataFrame, dtems_num:int, isams_num:int, start_time:datetime, verbose:bool=False) -> None:
    
    if verbose == True: student_details = getStudentDetails(on_dtems_but_not_isams_li, dtems_df, False)
    now = datetime.now()
    today = date.today()
    current_time = now.strftime("%H:%M:%S")
    time_taken = "{0:.2f}".format(time() - start_time)
    time_string =  f"Ran attendance check at {current_time} on {today} on computer {os.environ['COMPUTERNAME']} in {time_taken} seconds"
    logger.info(time_string)
    body = f"""
    <html>
    <head></head>
    <body>
        <h1>DTEMS: {dtems_num} iSAMS: {isams_num}</h1><br></br>
        <p>{time_string}<br></br></p>
        <h2>Found {len(on_dtems_but_not_isams_li)} students on DTEMS but not on iSAMS:</h2>
        <ul>
    """
    for student in on_dtems_but_not_isams_li:
        body += f"<li>  {student['firstname']} {student['lastname']} ({student['student_id']})</li>"
    
    body += f"""
        </ul>
        <br>
        <h2>Found {len(on_isams_but_not_dtems_li)} students on iSAMS but not on DTEMS:</h2>
        <ul>
    """
    for student in on_isams_but_not_dtems_li:
        body += f"<li>  {student['firstname']} {student['lastname']} ({student['student_id']})</li>"

    body += """
        </ul>
        <br>
        <h3>Please be aware that if a student does not have any classes on their timetable they will not show up in the DTEMS check.</h3>
    """

    if verbose == True: body += student_details

    body += """
    </body>
    </html>
    """

    username = keyring.get_password("attendance_sharepoint", "username")
    password = keyring.get_password("attendance_sharepoint", username)

    send_email("Discrepancy Check Results", body, username, password, "goconnor@instituteofeducation.ie")

def getStudentDetails(on_dtems_but_not_isams_li:list[dict], df:pd.DataFrame, standalone:bool=True) -> str:
    #df = getDTEMSDataframe(ctx, tempPath, "FT_ClassDetails.csv")

    body = ""
    row_count = 0

    def check(li:list[dict], body:str, row_count:int) -> tuple[str, int]:
        for student in li:
            body += f"<h4>{student['firstname']} {student['lastname']}</h4><ul>"
            for index, row in df.iterrows():
                missing_student_num = int(student['student_id'])
                student_num = int(row[8])

                if missing_student_num == student_num:
                    details = f"{int(row[1])} {int(row[2])} - {student_num} - {row[6]} {row[7]}"
                    body += f"<li> {details} </li>"
                    row_count += 1
                    
            body += "</ul>"

        return body, row_count

    if standalone == True: body += "<html><head></head><body>"
    body += "<h2>On DTEMS but not on iSAMS:</h2>"
    body, row_count = check(on_dtems_but_not_isams_li, body, row_count)
    if standalone == True: body += "</body></html>"

    username = keyring.get_password("attendance_sharepoint", "username")
    password = keyring.get_password("attendance_sharepoint", username)

    if row_count > 0 and standalone == True: send_email("Discrepancy Details Results", body, username, password, "goconnor@instituteofeducation.ie")
    return body

def main() -> None:
    start_time = time()
    ctx, api_url, auth_server_url, client_id, client_secret = getAuth()

    tempPath = "./.temp/"
    os.makedirs(tempPath, exist_ok=True)

    dtems_students, dtems_df = getDTEMSData(ctx, tempPath)
    isams_students = getListOfStudents(auth_server_url, client_id, client_secret, api_url)

    dtems_num = len(dtems_students)
    isams_num = len(isams_students)

    on_dtems_but_not_isams_li, on_isams_but_not_dtems_li = compareData(dtems_students, isams_students)

    sendResultsEmail(on_dtems_but_not_isams_li, on_isams_but_not_dtems_li, dtems_df, dtems_num, isams_num, start_time, verbose=False)

    shutil.rmtree(tempPath)

if __name__ == "__main__":
    main()