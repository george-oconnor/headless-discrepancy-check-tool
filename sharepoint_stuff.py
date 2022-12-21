from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder
import os, ctypes, logging
from sharepy import connect, SharePointSession
from requests.models import Response

os.makedirs("./logs", exist_ok=True)
logger = logging.getLogger(__name__)
handler = logging.FileHandler('./logs/SharePoint.log')
formatter = logging.Formatter("%(asctime)s | %(name)s | %(levelname)s | %(message)s")
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.INFO)
logger.debug("-------Starting Execution-------")

def getCTX(url:str, username:str, password:str) -> ClientContext:
    ctx_auth = AuthenticationContext(url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(url, ctx_auth)
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        logger.debug("Web title: {0}".format(web.properties['Title']))
        logger.info("Got CTX")
    else:
        logger.error(ctx_auth.get_last_error())
    return ctx

def downloadFile(ctx:ClientContext, relative_url:str, output_filename:str, output_location:str, hidden:bool=False) -> None:
    os.makedirs(output_location, exist_ok=True)
    if hidden == True:
        FILE_ATTRIBUTE_HIDDEN = 0x02
        ret = ctypes.windll.kernel32.SetFileAttributesW(output_location, FILE_ATTRIBUTE_HIDDEN)
        logger.info(f"Hid file location {output_location}")
    try:
        with open(output_location + output_filename, 'wb') as output_file:
                response = File.open_binary(ctx, relative_url)
                output_file.write(response.content)
                logger.debug(f"Downloaded and wrote file: {output_filename}")
    except Exception as e:
        logger.error(f"Error downloading file {output_filename}")
        logger.error(e, exc_info=True)

def returnAllContents(ctx:ClientContext, relativeUrl:str, get_files:bool=True, get_folders:bool=False) -> list:
    file_li, folder_li = [], []
    try:
        libraryRoot = ctx.web.get_folder_by_server_relative_url(relativeUrl)
        ctx.load(libraryRoot)
        ctx.execute_query()
    except Exception as e:
        logger.error('Problem getting directory info')
        logger.error(e, exc_info=True)
        return []

    if get_folders == True:
        try:
            folders = libraryRoot.folders
            ctx.load(folders)
            ctx.execute_query()
            for myfolder in folders:
                logger.debug("Folder name: {0}".format(myfolder.properties["Name"]))
                logger.debug("Folder name: {0}".format(myfolder.properties["ServerRelativeUrl"]))
                #returnAllContents(ctx, relativeUrl + '/' + myfolder.properties["Name"])
                pathList = myfolder.properties["ServerRelativeUrl"].split('/')
                folder_li.append(pathList[-1])
            logger.info(f"Got {len(folder_li)} folders from {relativeUrl}")
        except Exception as e:
            logger.error(f"Problem returning folders from {relativeUrl}")
            logger.error(e, exc_info=True)

    if get_files == True:
        try:
            files = libraryRoot.files
            ctx.load(files)
            ctx.execute_query()
            for myfile in files:
                logger.debug("File name: {0}".format(myfile.properties["ServerRelativeUrl"]))
                pathList = myfile.properties["ServerRelativeUrl"].split('/')
                file_li.append(pathList[-1])
            logger.info(f"Got {len(file_li)} files from {relativeUrl}")
        except Exception as e:
            logger.error(f"Problem returning files from {relativeUrl}")
            logger.error(e, exc_info=True)

    if get_folders == True and get_files == True:
        logger.info(f"Returned {len(folder_li)} folders and {len(file_li)} files from {relativeUrl}")
        return [folder_li, file_li]
    elif get_folders == False and get_files == True:
        logger.info(f"Returned {len(file_li)} files from {relativeUrl}")
        return file_li
    elif get_folders == True and get_files == False:
        logger.info(f"Returned {len(folder_li)} folders from {relativeUrl}")
        return folder_li
    else:
        logger.error(f"neither files nor folders selected for {relativeUrl}")
        return []

def uploadFile(ctx:ClientContext, file:str, filepath:str, rel_path:str) -> None:
    try:
        with open(filepath, 'rb') as content_file:
            file_content = content_file.read()
            logger.debug(f"Opened {file} successfully")
    except Exception as e:
        logger.error(f"Error opening {file}")
        logger.error(e, exc_info=True)
        return None

    try:
        upload = ctx.web.get_folder_by_server_relative_url(rel_path).upload_file(file, file_content).execute_query()
        logger.debug(f"Uploaded {file} successfully")
    except Exception as e:
        logger.error(f"Error uploading {file}")
        logger.error(e, exc_info=True)

def deleteFile(ctx:ClientContext, relativeUrl:str) -> None:
    try:
        ctx.web.get_file_by_server_relative_url(relativeUrl).delete_object().execute_query()
        logger.info(f"Successfully deleted the file at {relativeUrl}")
    except Exception as e:
        logger.error(f"Failed to delete the file at {relativeUrl}")
        logger.error(e, exc_info=True)

def createFolders(ctx:str, relative_url:str) -> None:
    folder = ctx.web.ensure_folder_path(relative_url).execute_query()

def getSharepyAuth(username:str, password:str) -> SharePointSession:
    s = connect("ioedub.sharepoint.com", username, password)
    return s

def getSiteLists(s:SharePointSession, site:str) -> Response:
    r = s.get(f"{site}/_api/web/lists")
    return r

def getSiteList(s:SharePointSession, site:str, list:str) -> Response:
    r = s.get(f"{site}/_api/web/lists/GetByTitle('{list}')")
    return r

def createSiteList(s:SharePointSession, site:str, extra_json:dict) -> Response:
    """{
            "__metadata": 
            {
                "type": "SP.List"
            },
            "AllowContentTypes": true,
            "BaseTemplate": 100,
            "ContentTypesEnabled": true,
            "Description": "My list description",
            "Title": "Test"
        }"""
    r = s.post(f"{site}/_api/web/lists", json=extra_json)
    return r

def deleteSiteList(s:SharePointSession, site:str, list_guid:str) -> Response:
    r = s.post(f"{site}/_api/web/lists(guid'{list_guid}')")
    return r

logger.debug("-------Finished Execution-------")