# coding=utf-8

import pyimgur
import keyboard
import os.path
import pickle
import pyscreenshot as ImageGrab
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import os
import datetime
from time import sleep
from tkinter import Tk
# imgur key
CLIENT_ID = "815429ae05941e1"

print("  .oooooo.             o8o               .             oooo   o8o  ")
print(" d8P'  `Y8b            `'"'"'"             .o8             `888   `'"'"'"         ")
print("888          oooo d8b oooo   .oooo.o .o888oo  .oooo.    888  oooo  oooo    ooo")
print("888          `888'"''"'8P `888  d88(  '"'8   888   `P  )88b   888  `888   `88b..8P"'"  ")
print("888           888      888  `'"'Y88b.    888    .oP"888   888   888     Y888"'" ")
print("`88b    ooo   888      888  o.  )88b   888 . d8(  888   888   888   .o8'"'"'"88b ")
print(" `Y8bood8P'  d888b    o888o 8'""'888P'   '"'888'"' `Y888'""'8o o888o o888o o88'   888o ")
print('\n1v\t\t\t\tSTAFF HELPER\t\t\t\tBy VAFI')
print('\t\t\t\t\t\t\thttps://github.com/VAFI\n')

if not os.path.exists("key.json"):
    print("Key for google sheets not found")
    error = input()
    exit()
try:
    if not os.path.exists("settings.cfg"):
        print("Crating config file...")
        conf = open('settings.cfg', 'w')
        confLines = {'email': '', 'tableName': '', 'resolutionX': 1920, 'resolutionY': 1080, "tableID": 'none', 'count': 1}
        email = input("Почта для доступа к таблице: ")
        confLines["email"] = email
        tableName = input("Название таблицы: ")
        confLines["tableName"] = tableName
        conf.write(str(confLines))
        conf.close()
        print("Done! Place, restart program")

    conf = open('settings.cfg', 'r')
    confValues = eval(conf.read())
    conf.close()
    print("name: " + confValues['tableName'])
    print("email: " + confValues['email'])
    if confValues['tableID'] == 'none':

        CREDENTIALS_FILE = 'key.json'  # Имя файла с закрытым ключом, вы должны подставить свое
        # Читаем ключи из файла
        credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                    ['https://www.googleapis.com/auth/spreadsheets',
                                                                        'https://www.googleapis.com/auth/drive'])

        httpAuth = credentials.authorize(httplib2.Http())  # Авторизуемся в системе
        service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)  # Выбираем работу с таблицами и 4 версию API

        spreadsheet = service.spreadsheets().create(body={
            'properties': {'title': 'CristalixStaffTweak$' + confValues['email'], 'locale': 'ru_RU'},
            'sheets': [{'properties': {'sheetType': 'GRID',
                                    'sheetId': 0,
                                    'title': confValues['tableName'],
                                    'gridProperties': {'rowCount': 100, 'columnCount': 15}}}]
        }).execute()
        spreadsheetId = spreadsheet['spreadsheetId']  # сохраняем идентификатор файла

        driveService = apiclient.discovery.build('drive', 'v3',
                                                http=httpAuth)  # Выбираем работу с Google Drive и 3 версию API
        access = driveService.permissions().create(
            fileId=spreadsheetId,
            body={'type': 'user', 'role': 'writer', 'emailAddress': confValues['email']},
            # Открываем доступ на редактирование
            fields='id'
        ).execute()
        confValues['tableID'] = spreadsheetId
        conf = open('settings.cfg', 'w')
        conf.write(str(confValues))
        conf.close()
        print('Таблица создана')
    print('url: ' + 'https://docs.google.com/spreadsheets/d/' + confValues['tableID'])
    print('\nCtrl + G - пред \n Ctrl + M - мут \n Ctrl + B - бан')
    spreadsheetId = confValues['tableID']
    CREDENTIALS_FILE = 'key.json'  # Имя файла с закрытым ключом, вы должны подставить свое
    # Читаем ключи из файла
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                    ['https://www.googleapis.com/auth/spreadsheets',
                                                                        'https://www.googleapis.com/auth/drive'])

    httpAuth = credentials.authorize(httplib2.Http())  # Авторизуемся в системе
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)  # Выбираем работу с таблицами и 4 версию API
    # Получаем список листов, их Id и название
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheetList = spreadsheet.get('sheets')

    def screenshot(resolX, resolY):
        keyboard.send('t')
        sleep(0.1)
        screenshot = ImageGrab.grab(bbox=(0, 0, resolX, resolY))
        keyboard.send('esc')
        screenshot.save('screenshots.png')
        print("Загрузка на скриншота")
        PATH = "screenshots.png"
        im = pyimgur.Imgur(CLIENT_ID)
        uploaded_image = im.upload_image(PATH, title="CristalixStaffTweaks")

        return uploaded_image.link
    def get_name(index):
        keyboard.send('t')
        sleep(0.1)
        keyboard.send('UP')
        sleep(0.1)
        keyboard.press('Ctrl')
        sleep(0.1)
        keyboard.send('a')
        sleep(0.1)
        keyboard.send('x')
        sleep(0.1)
        keyboard.release('Ctrl')
        sleep(0.01)
        keyboard.send('esc')
        c = Tk()
        c.withdraw()
        clip = c.clipboard_get()
        c.update()
        c.destroy()
        clip = clip.replace('!', '').split()
        return str(clip[index])




    def join_time():
        sheetId = sheetList[0]['properties']['sheetId']
        count = confValues['count'] + 1
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                        body={
                                                            "valueInputOption": "USER_ENTERED",
                                                            "data": [
                                                                {"range": confValues['tableName'] + "!" + "C" + str(
                                                                    count),
                                                                "majorDimension": "ROWS",
                                                                "values": [
                                                                    [datetime.datetime.now().strftime("%H:%M:%S")]]}
                                                            ]
                                                        }).execute()
        print("Время захода учтено")
        confValues['count'] = confValues['count'] + 1
        conf = open('settings.cfg', 'w')
        conf.write(str(confValues))
        conf.close()


    def leave_time():
        sheetId = sheetList[0]['properties']['sheetId']
        count = confValues['count'] + 1
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                    body={
                                                        "valueInputOption": "USER_ENTERED",
                                                        "data": [
                                                            {"range": confValues['tableName'] + "!" + "D" + str(
                                                                count),
                                                            "majorDimension": "ROWS",
                                                            "values": [
                                                                [datetime.datetime.now().strftime("%H:%M:%S")]]}
                                                        ]
                                                    }).execute()
        print("Время выхода")
        confValues['count'] = confValues['count'] + 1
        conf = open('settings.cfg', 'w')
        conf.write(str(confValues))
        conf.close()
        exit()


    def upload_mute():
        sheetId = sheetList[0]['properties']['sheetId']
        nickname = get_name(1)
        screenshot_link = screenshot(confValues['resolutionX'], confValues['resolutionY'])



        count = confValues["count"]
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                        body={
                                                            "valueInputOption": "USER_ENTERED",
                                                            "data": [
                                                                {"range": confValues['tableName'] + "!" + "G" + str(
                                                                    count),
                                                                "majorDimension": "ROWS",
                                                                "values": [[screenshot_link]]}
                                                            ]
                                                        }).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                        body={
                                                            "valueInputOption": "USER_ENTERED",
                                                            "data": [
                                                                {"range": confValues['tableName'] + "!" + "E" + str(
                                                                    count),
                                                                "majorDimension": "ROWS",
                                                                "values": [['мут']]}
                                                            ]
                                                        }).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                        body={
                                                            "valueInputOption": "USER_ENTERED",
                                                            "data": [
                                                                {"range": confValues['tableName'] + "!" + "F" + str(
                                                                    count),
                                                                "majorDimension": "ROWS",
                                                                "values": [[nickname]]}
                                                            ]
                                                        }).execute()
        print("Мут " + nickname + ' ' + screenshot_link)
        confValues['count'] = confValues['count'] + 1
        conf = open('settings.cfg', 'w')
        conf.write(str(confValues))
        conf.close()

    def upload_ban():
        sheetId = sheetList[0]['properties']['sheetId']
        nickname = get_name(1)
        screenshot_link = screenshot(confValues['resolutionX'], confValues['resolutionY'])



        count = confValues["count"]
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                        body={
                                                            "valueInputOption": "USER_ENTERED",
                                                            "data": [
                                                                {"range": confValues['tableName'] + "!" + "G" + str(
                                                                    count),
                                                                "majorDimension": "ROWS",
                                                                "values": [[screenshot_link]]}
                                                            ]
                                                        }).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                        body={
                                                            "valueInputOption": "USER_ENTERED",
                                                            "data": [
                                                                {"range": confValues['tableName'] + "!" + "E" + str(
                                                                    count),
                                                                "majorDimension": "ROWS",
                                                                "values": [['бан']]}
                                                            ]
                                                        }).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                        body={
                                                            "valueInputOption": "USER_ENTERED",
                                                            "data": [
                                                                {"range": confValues['tableName'] + "!" + "F" + str(
                                                                    count),
                                                                "majorDimension": "ROWS",
                                                                "values": [[nickname]]}
                                                            ]
                                                        }).execute()
        print("Бан " + nickname + ' ' + screenshot_link)
        confValues['count'] = confValues['count'] + 1
        conf = open('settings.cfg', 'w')
        conf.write(str(confValues))
        conf.close()



    def upload_warn():
        sheetId = sheetList[0]['properties']['sheetId']
        count = confValues["count"]
        nickname = get_name(0)
        screenshot_link = screenshot(confValues['resolutionX'], confValues['resolutionY'])


        count = confValues["count"]
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                        body={
                                                            "valueInputOption": "USER_ENTERED",
                                                            "data": [
                                                                {"range": confValues['tableName'] + "!" + "G" + str(
                                                                    count),
                                                                "majorDimension": "ROWS",
                                                                "values": [[screenshot_link]]}
                                                            ]
                                                        }).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                        body={
                                                            "valueInputOption": "USER_ENTERED",
                                                            "data": [
                                                                {"range": confValues['tableName'] + "!" + "E" + str(
                                                                    count),
                                                                "majorDimension": "ROWS",
                                                                "values": [['пред']]}
                                                            ]
                                                        }).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                        body={
                                                            "valueInputOption": "USER_ENTERED",
                                                            "data": [
                                                                {"range": confValues['tableName'] + "!" + "F" + str(
                                                                    count),
                                                                "majorDimension": "ROWS",
                                                                "values": [[nickname]]}
                                                            ]
                                                        }).execute()
        print("Пред " + nickname + ' ' +screenshot_link)
        confValues['count'] = confValues['count'] + 1
        conf = open('settings.cfg', 'w')
        conf.write(str(confValues))
        conf.close()


    join_time()
    while True:
        keyboard.add_hotkey('Ctrl + m', lambda: upload_mute())
        keyboard.add_hotkey('Ctrl + b', lambda: upload_ban())
        keyboard.add_hotkey('Ctrl + g', lambda: upload_warn())
        keyboard.add_hotkey('Ctrl + c', lambda: leave_time())
        keyboard.wait('Ctrl+alt+l')
except:
    error = input()
    exit()
