# coding=utf-8

import pyimgur
import keyboard
import os.path
import pickle
from PIL import ImageGrab
import httplib2
from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
import os
import datetime
from time import sleep
from tkinter import Tk
from playsound import playsound

# imgur key
CLIENT_ID = "815429ae05941e1"

print("  .oooooo.             o8o               .             oooo   o8o  ")
print(" d8P'  `Y8b            `'"'"'"             .o8             `888   `'"'"'"         ")
print("888          oooo d8b oooo   .oooo.o .o888oo  .oooo.    888  oooo  oooo    ooo")
print("888          `888'"''"'8P `888  d88(  '"'8   888   `P  )88b   888  `888   `88b..8P"'"  ")
print("888           888      888  `'"'Y88b.    888    .oP"888   888   888     Y888"'" ")
print("`88b    ooo   888      888  o.  )88b   888 . d8(  888   888   888   .o8'"'"'"88b ")
print(" `Y8bood8P'  d888b    o888o 8'""'888P'   '"'888'"' `Y888'""'8o o888o o888o o88'   888o ")
print('\nv2\t\t\t\tSTAFF HELPER\t\t\t\tBy VAFI')
print('\t\t\t\t\t\t\thttps://github.com/VAFI\n')

if not os.path.exists("key.json"):
    print("Key for google sheets not found")
    error = input()
    exit()
if not os.path.exists("ok.mp3"):
    print("Если вы удалили ok.mp3 - не надо так делать. Просто уменьшите громкость в микшере")
    print("Если нет - сообщите разработчику")
    exit()

try:
    if not os.path.exists("settings.cfg"):
        print("Crating config file...")
        conf = open('settings.cfg', 'w')
        confLines = {'email': '', 'tableName': '', 'resolutionX': 1920, 'resolutionY': 1080, "tableID": 'none',
                     'count': 1, 'sound': False}
        email = input("Почта для доступа к таблице: ")
        confLines["email"] = email
        tableName = input("Название таблицы: ")
        user_audio_answer = input("Воспроизводить ли звук при удачной загрузке? [Y/n] : ")
        if user_audio_answer == '' or user_audio_answer == 'y' or user_audio_answer == 'yes' or user_audio_answer == 'да' or user_audio_answer == 'д':
            confLines["sound"] = True
        confLines["tableName"] = tableName
        conf.write(str(confLines))
        conf.close()

    conf = open('settings.cfg', 'r')
    confValues = eval(conf.read())
    conf.close()
    print("screen resolution:" + str((confValues['resolutionX'])) + 'x' + str(confValues['resolutionY']))
    print("name: " + confValues['tableName'])
    print("email: " + confValues['email'])
    if confValues['tableID'] == 'none': 
        CREDENTIALS_FILE = 'key.json'  # Имя файла с закрытым ключом, вы должны подставить свое
        # Читаем ключи из файла
        credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                       ['https://www.googleapis.com/auth/spreadsheets',
                                                                        'https://www.googleapis.com/auth/drive'])

        httpAuth = credentials.authorize(httplib2.Http())  # Авторизуемся в системе
        service = discovery.build('sheets', 'v4', http=httpAuth)  # Выбираем работу с таблицами и 4 версию API

        spreadsheet = service.spreadsheets().create(body={
            'properties': {'title': 'CristalixStaffTweak$' + confValues['email'], 'locale': 'ru_RU'},
            'sheets': [{'properties': {'sheetType': 'GRID',
                                       'sheetId': 0,
                                       'title': confValues['tableName'],
                                       'gridProperties': {'rowCount': 100, 'columnCount': 15}}}]
        }).execute()
        spreadsheetId = spreadsheet['spreadsheetId']  # сохраняем идентификатор файла

        driveService = discovery.build('drive', 'v3',
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
        sheetList = spreadsheet.get('sheets')
        sheetId = sheetList[0]['properties']['sheetId']

        # Make up for sheet
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "K" + "3","majorDimension": "ROWS","values": [["Онлайн"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "L" + "3","majorDimension": "ROWS","values": [["=СУММ(D:D)-СУММ(C:C)"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "K" + "4","majorDimension": "ROWS","values": [["Выдано наказаний"]]}]}).execute()
        results = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId,body={"requests": [
                                                            {'mergeCells': {'range': {'sheetId': sheetId,'startRowIndex': 4,'endRowIndex': 4,'startColumnIndex': 11,'endColumnIndex': 12},'mergeType': 'MERGE_ALL'}}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "K" + "5","majorDimension": "ROWS","values": [["Мутов"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "L" + "5","majorDimension": "ROWS","values": [['=СЧЁТЕСЛИМН(E:E; "мут")']]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "K" + "6","majorDimension": "ROWS","values": [["Банов"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "L" + "6","majorDimension": "ROWS","values": [['=СЧЁТЕСЛИМН(E:E; "бан")']]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "K" + "7","majorDimension": "ROWS","values": [["Предов"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "L" + "7","majorDimension": "ROWS","values": [['=СЧЁТЕСЛИМН(E:E; "пред")']]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "K" + "8","majorDimension": "ROWS","values": [["Кик"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "L" + "8","majorDimension": "ROWS","values": [['=СЧЁТЕСЛИМН(E:E; "кик")']]}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"requests":[{"repeatCell":{"cell":{"userEnteredFormat":{
                                                            "backgroundColor": {"red": 0.0,"green": 0.5,"blue": 0.5,"alpha": 1},}},
                                                            "range":{"sheetId": sheetId,"startRowIndex": 0,"endRowIndex": 99,"startColumnIndex": 0,"endColumnIndex": 2},"fields": "userEnteredFormat"}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"requests":[{"repeatCell":{"cell":{"userEnteredFormat":{
                                                            "backgroundColor": {"red": 0.0,"green": 0.5,"blue": 0.5,"alpha": 1},}},
                                                            "range":{"sheetId": sheetId,"startRowIndex": 0,"endRowIndex": 99,"startColumnIndex": 7,"endColumnIndex": 9},"fields": "userEnteredFormat"}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"requests":[{"repeatCell":{"cell":{"userEnteredFormat":{
                                                            "backgroundColor": {"red": 0.0,"green": 0.5,"blue": 0.5,"alpha": 1},}},
                                                            "range":{"sheetId": sheetId,"startRowIndex": 0,"endRowIndex": 1,"startColumnIndex": 0,"endColumnIndex": 9},"fields": "userEnteredFormat"}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"requests": [{'mergeCells': 
                                                            {'range': {'sheetId': sheetId,'startRowIndex': 0,'endRowIndex': 1,'startColumnIndex': 0,'endColumnIndex': 9},'mergeType': 'MERGE_ALL'}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"requests":[{"repeatCell":{"cell":{"userEnteredFormat":{
                                                            "backgroundColor": {"red": 1,"green": 0.0,"blue": 0.0,"alpha": 1},}},
                                                            "range":{"sheetId": sheetId,"startRowIndex": 2,"endRowIndex": 8,"startColumnIndex": 10,"endColumnIndex": 12},"fields": "userEnteredFormat"}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body={"requests": [{"updateDimensionProperties": {
                                                            "range": {"sheetId": sheetId,"dimension": "COLUMNS","startIndex": 0,"endIndex": 2},
                                                            "properties": {"pixelSize": 10},"fields": "pixelSize"}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body={"requests": [{"updateDimensionProperties": {
                                                            "range": {"sheetId": sheetId,"dimension": "COLUMNS","startIndex": 7,"endIndex": 9},
                                                            "properties": {"pixelSize": 10},"fields": "pixelSize"}}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "L" + "8","majorDimension": "ROWS","values": [['=СЧЁТЕСЛИМН(E:E; "кик")']]}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body={"requests":[{'repeatCell': {
                                                            'range': {'startRowIndex': 0,'endRowIndex': 4,'startColumnIndex': 11,'endColumnIndex': 13,},
                                                            'cell': {'userEnteredFormat': {'numberFormat': {'type': 'TIME','pattern': 'h:mm:ss',},},},'fields': 'userEnteredFormat.numberFormat',}}]}).execute()
    
    print('url: ' + 'https://docs.google.com/spreadsheets/d/' + confValues['tableID'])
    print('\nCtrl + G - записать нарушение \nCtrl + J - записать время выхода')
    spreadsheetId = confValues['tableID']
    CREDENTIALS_FILE = 'key.json'  # Имя файла с закрытым ключом
    # Читаем ключи из файла
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                    ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])

    httpAuth = credentials.authorize(httplib2.Http())  # Авторизуемся в системе
    service = discovery.build('sheets', 'v4', http=httpAuth)  # Выбираем работу с таблицами и 4 версию API
    # Получаем список листов, их Id и название
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheetList = spreadsheet.get('sheets')


    def play_sound():
        if confValues['sound'] == True:
            playsound('ok.mp3', False)
        else:
            return True


    def screenshot(resolX, resolY):
        keyboard.send('t')
        sleep(0.2)
        screenshot = ImageGrab.grab(bbox=(0, 0, resolX, resolY))
        keyboard.send('esc')
        screenshot.save('screenshot.png')
        print("Загрузка скриншота")
        PATH = "screenshot.png"
        im = pyimgur.Imgur(CLIENT_ID)
        uploaded_image = im.upload_image(PATH, title="CristalixStaffTweaks")

        return uploaded_image.link


    def get_text():
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
        return (clip)


    def detect_violation():
        text = get_text()
        if "/mute" in text:
            upload(text[1], "мут")
            return ("мут")  # прерывание
        if "/kick" in text:
            upload(text[1], "кик")
            return ("кик")  # прерывание
        if "/ban" in text:
            upload(text[1], "бан")
            return ("бан")  # прерывание
        upload(text[0], "пред")


    def upload(nickname, type_violation):

        screenshot_link = screenshot(confValues['resolutionX'], confValues['resolutionY'])

        count = confValues["count"]
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "G" + str(count),"majorDimension": "ROWS","values": [[screenshot_link]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "E" + str(count),"majorDimension": "ROWS","values": [[type_violation]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "F" + str(count),"majorDimension": "ROWS","values": [[nickname]]}]}).execute()
        play_sound()
        print(type_violation + ' ' + nickname + ' ' + screenshot_link)
        confValues['count'] = confValues['count'] + 1
        conf = open('settings.cfg', 'w')
        conf.write(str(confValues))
        conf.close()


    def time_join():
        count = confValues['count'] + 1
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "C" + str(count),"majorDimension": "ROWS","values": [[datetime.datetime.now().strftime("%H:%M:%S")]]}]}).execute()
        print("Время захода учтено")
        
        confValues['count'] = confValues['count'] + 1
        conf = open('settings.cfg', 'w')
        conf.write(str(confValues))
        conf.close()


    def time_exit():
        count = confValues['count'] + 1
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": confValues['tableName'] + "!" + "D" + str(count),"majorDimension": "ROWS","values": [[datetime.datetime.now().strftime("%H:%M:%S")]]}]}).execute()
        print("Время выхода учтено")
        confValues['count'] = confValues['count'] + 1
        conf = open('settings.cfg', 'w')
        conf.write(str(confValues))
        conf.close()
        exit()


    time_join()
    while True:
        keyboard.add_hotkey('Ctrl + g', lambda: detect_violation())
        keyboard.add_hotkey('Ctrl + j', lambda: time_exit())
        keyboard.wait('Ctrl+alt+l')

except:
    print("О нет! Произошла ошибка! Скиньте текст ниже https://vk.com/outsider77")
    error = input()
    exit()