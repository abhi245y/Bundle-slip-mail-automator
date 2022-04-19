
import ezgmail
import pprint
import glob
import email

import datetime
import json
import eml_parser
import os
import shutil

import img2pdf
from PIL import Image

from PyPDF2 import PdfFileMerger

from imutils.perspective import four_point_transform
import cv2

import pandas as pd


""" [Gmail Search Function Filter Note]

    Returns a list of GmailThread objects that match the search query.

    The ``query`` string is exactly the same as you would type in the Gmail search box, and you can use the search
    operatives for it too:

        * label:UNREAD
        * from:al@inventwithpython.com
        * subject:hello
        * has:attachment
        * after:2004/04/16
        * before:2004/04/16

    More are described at https://support.google.com/mail/answer/7190?hl=en
"""


def json_serial(obj):
    if isinstance(obj, datetime.datetime):
        serial = obj.isoformat()
        return serial


def cropImage():
    print("\n## Cropping Images to Right Size ##")
    files = glob.glob("jpgs/*.jpg")

    for imgs in files:
        imgFile = open(imgs)
        image = cv2.imread(imgFile.name)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (5, 5), 0)
        thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

        # Find contours and sort for largest contour
        cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cnts = cnts[0] if len(cnts) == 2 else cnts[1]
        cnts = sorted(cnts, key=cv2.contourArea, reverse=True)
        displayCnt = None

        for c in cnts:
            # Perform contour approximation
            peri = cv2.arcLength(c, True)
            approx = cv2.approxPolyDP(c, 0.02 * peri, True)
            if len(approx) == 4:
                displayCnt = approx
                break

        # Obtain birds' eye view of image
        warped = four_point_transform(image, displayCnt.reshape(4, 2))

        cv2.imwrite("./cropped/" + os.path.basename(imgFile.name), warped)
        cv2.waitKey()


def imgToPDF():
    print("\n## Converting Images To PDF ##")

    files = glob.glob("cropped/*.jpg")

    pdf_path = r"./pdfToCombine/"
    for croppedImgs in files:
        image = Image.open(croppedImgs)
        pdf_bytes = img2pdf.convert(image.filename)
        fileName = pdf_path+os.path.basename(open(croppedImgs).name).split(".")[0]+".pdf"
        file = open(fileName, "wb")
        file.write(pdf_bytes)
        image.close()
        file.close()


def extractPayloadFromEml():
    print("\n## Extracting attachments from .eml files ##")
    files = glob.glob("files/*.eml")
    for each in files:
        msg = email.message_from_file(open(each))
        attachments = msg.get_payload()
        for attachment in attachments:
            try:
                fnam = attachment.get_filename()
                f = open("./extracted/" + fnam, 'wb').write(attachment.get_payload(decode=True))
                f.close()
            except Exception as detail:
                pass


def mergePDF():
    print("\n## Merging All PDF Files to one ##")

    merger = PdfFileMerger()

    files = glob.glob("pdfToCombine/*.pdf")

    for pdfs in files:
        merger.append(open(pdfs, "rb"))

    merger.write("./output/merged_pdf.pdf")
    merger.close()


def mergeXlsx():
    # not read to be executed

    xlFiles = glob.glob("xls/*.xlsx")

    excl_merged = pd.DataFrame()

    for xlFile in xlFiles:
        df = pd.concat(pd.read_excel(xlFile, sheet_name=None),
                       ignore_index=True, sort=False)
        excl_merged = excl_merged.append(
            df, ignore_index=True)

    excl_merged.to_excel('./output/combined_excel.xlsx', index=False)


def moveFiles():
    print("\n## Moving files to respective folders ##")

    folderToCheck = ['files', 'extracted']

    for folder in folderToCheck:
        jpgFiles = glob.glob(folder + "/*.jpg")
        pdfFiles = glob.glob(folder + "/*.pdf")
        xlFiles = glob.glob(folder + "/*.xlsx")

        for jpgFile in jpgFiles:
            jpgFileSrc = open(jpgFile)
            shutil.copy(jpgFileSrc.name, "./jpgs/")

        for pdfFile in pdfFiles:
            pdfFileSrc = open(pdfFile)
            shutil.copy(pdfFileSrc.name, "./pdfToCombine/")

        for xlFile in xlFiles:
            xlFileSrc = open(xlFile)
            shutil.copy(xlFileSrc.name, "./output/")


def main(filters):
    ezgmail.init()
    threads = ezgmail.search(filters)
    print("\n## Gathering Mails ##")
    for thread in threads:
        for message in thread.messages:
            try:
                pprint.pprint(message.downloadAllAttachments(downloadFolder="./files"))
                print("\n## Downloaded ##")
            except Exception as e:
                print(e)
                pass


def clearOldFiles():
    print("\n## Clearing Old Files ##")
    folderToCheck = ['./files', './extracted', './cropped', './jpgs', './pdfToCombine', './xls', './temp']

    for folder in folderToCheck:
        for f in os.listdir(folder):
            os.remove(os.path.join(folder, f))

def extraFunction():
    """
    dump mail address from .eml file to college_emails.json in config folder

    """
    files = glob.glob("files/*.eml")
    emailAddress = []
    for each in files:
        with open(each, 'rb') as fhdl:
            raw_email = fhdl.read()
        ep = eml_parser.EmlParser()
        parsed_eml = ep.decode_email_bytes(raw_email)
        current = json.loads(json.dumps(parsed_eml, default=json_serial))["header"]["from"]
        emailAddress.append(current)

    emailAddress = list(set(emailAddress))

    seen = set()
    result = []
    for item in emailAddress:
        if item not in seen:
            seen.add(item)
            result.append(item)

    emailListJson = {"emailIds":result}

    try:
        with open('config/college_emails.json', 'r+', encoding='utf-8') as f:
            json.dump(emailListJson, f, ensure_ascii=False, indent=4)
    except:
        with open('config/college_emails.json', 'a', encoding='utf-8') as f:
            json.dump(emailListJson, f, ensure_ascii=False, indent=4)

    pprint.pprint(result)


if __name__ == '__main__':
    with open('./config/college_emails.json', 'r') as college_emails:
        college_emails = json.load(college_emails)

    emailsList = college_emails["emailIds"]

    clearOldFiles()

    # for emailId in emailsList:
    #     query = 'from:'+emailId
    #     main(query)

    main('from:socdunit@keralauniversity.ac.in')
    extractPayloadFromEml()
    moveFiles()
    cropImage()
    imgToPDF()
    mergePDF()

    print("***  All Done Out can be found at folder names 'output' ")