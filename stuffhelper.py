# coding=utf-8
from win32api import GetSystemMetrics
import sys
import random
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
import json
import sys
# imgur key
CLIENT_ID = "815429ae05941e1"

print("  .oooooo.             o8o               .             oooo   o8o  ")
print(" d8P'  `Y8b            `'"'"'"             .o8             `888   `'"'"'"         ")
print("888          oooo d8b oooo   .oooo.o .o888oo  .oooo.    888  oooo  oooo    ooo")
print("888          `888'"''"'8P `888  d88(  '"'8   888   `P  )88b   888  `888   `88b..8P"'"  ")
print("888           888      888  `'"'Y88b.    888    .oP"888   888   888     Y888"'" ")
print("`88b    ooo   888      888  o.  )88b   888 . d8(  888   888   888   .o8'"'"'"88b ")
print(" `Y8bood8P'  d888b    o888o 8'""'888P'   '"'888'"' `Y888'""'8o o888o o888o o88'   888o ")
print('\nv3\t\t\t\tSTAFF HELPER\t\t\t\tBy VAFI')
print('\t\t\t\t\t\t\thttps://github.com/VAFI\n')

sys.stdout.reconfigure(encoding='utf-8')
def settins_configurator():
    settings = {'script_hotkey' : 'G' , 'resolutionX' : GetSystemMetrics(0) , 'resolutionY' : GetSystemMetrics(1), 'chat_hotkey' : 'T',
                    'second_screenshot_hotkey' : "H" , 'sound_activate' : True}
    print("Введите желаемую букву для горячей главиши для запуска скрипта. \nЧтобы оставить значение по умолчанию пропустите ответ. По умолчанию это G (CTRL+G)")
    script_hotkey = input("Значение буквы хоткея скрипта: ")
    if script_hotkey != '' : settings['script_hotkey'] = script_hotkey

    print("Если вы используете кнопку для чата не по умолчанию (T) введите данную клавишу. \nЧтобы оставить значение по умолчанию пропустите ответ. По умолчанию это T")
    chat_hotkey = input("Кнопка чата: ")
    if chat_hotkey != '' : settings['chat_hotkey'] = chat_hotkey

    print("Если вам вдруг нужен второй скриншот для наршуения вы можете нажать на CRTL+H.\nХотите поменять хот кей? Введите его ниже!\nЧтобы оставить значение по умолчанию пропустите ответ. По умолчанию это H")
    second_screenshot_hotkey = input("Кнопка второго скриншота: ")
    if second_screenshot_hotkey != '' : settings['second_screenshot_hotkey'] = second_screenshot_hotkey

    print("Вовпроизводить ли звук после успешной загрузки? \nЧтобы оставить значение по умолчанию пропустите ответ. По умолчанию это Y")
    sound_activate = input("[Y/n] : ")
    if sound_activate == 'n' or 'N' : 
        settings['sound_activate'] = False
    settings_file = open('settings.cfg', 'w', encoding='utf8')
    settings_file.write(str(settings))
    settings_file.close()

def account_configurator():
    print("Введите вашу почту для доступа к таблице")
    email = input("Почта: ")
    print("Введите название листа для таблицы")
    list_name = input("Название листа: ")
    account_info = {'email' : email, 'listName' : list_name, 'tableID' : table_creator(email, list_name), 'count' : 1 }
    settings_file = open('account_info.cfg', 'w', encoding='utf8')
    settings_file.write(str(account_info))
    settings_file.close()

def table_creator(email, list_name):
    print("Создание таблицы")
    CREDENTIALS_FILE = 'key.json'  # Имя файла с закрытым ключом, вы должны подставить свое
    # Читаем ключи из файла
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                    ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])

    httpAuth = credentials.authorize(httplib2.Http())  # Авторизуемся в системе
    service = discovery.build('sheets', 'v4', http=httpAuth,cache_discovery=False)  # Выбираем работу с таблицами и 4 версию API

    spreadsheet = service.spreadsheets().create(body={
        'properties': {'title': 'Cristalix Stuff Helper' + email, 'locale': 'ru_RU'},
        'sheets': [{'properties': {'sheetType': 'GRID',
                                    'sheetId': 0,
                                    'title': list_name,
                                    'gridProperties': {'rowCount': 9000, 'columnCount': 15}}}]
    }).execute()
    spreadsheetId = spreadsheet['spreadsheetId']  # сохраняем идентификатор файла

    driveService = discovery.build('drive', 'v3',
                                    http=httpAuth,cache_discovery=False)  # Выбираем работу с Google Drive и 3 версию API
    access = driveService.permissions().create(
        fileId=spreadsheetId,
        body={'type': 'user', 'role': 'writer', 'emailAddress': email},
        # Открываем доступ на редактирование
        fields='id'
    ).execute()
    sheetList = spreadsheet.get('sheets')
    sheet_id = sheetList[0]['properties']['sheetId']
    table_configurator(sheet_id,spreadsheet,service,list_name)
    print("Таблица создана")
    return  spreadsheetId

def table_configurator(sheetId,spreadsheet,service,list_name):
        color = {'red' :  random.uniform(0.1, 1), 'green' : random.uniform(0.1, 1), 'blue' :  random.uniform(0.1, 1)}
        color_stats = {'red' :  random.uniform(0.4, 1), 'green' : random.uniform(0.4, 1), 'blue' :  random.uniform(0.4, 1)}
      # Make up for sheet
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "L" + "3","majorDimension": "ROWS","values": [["Онлайн"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "M" + "3","majorDimension": "ROWS","values": [["=СУММ(D:D)-СУММ(C:C)"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "L" + "4","majorDimension": "ROWS","values": [["Выдано наказаний"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "L" + "5","majorDimension": "ROWS","values": [["Мутов"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "M" + "5","majorDimension": "ROWS","values": [['=СЧЁТЕСЛИМН(E:E; "мут")']]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "L" + "6","majorDimension": "ROWS","values": [["Банов"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "M" + "6","majorDimension": "ROWS","values": [['=СЧЁТЕСЛИМН(E:E; "бан")']]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "L" + "7","majorDimension": "ROWS","values": [["Предов"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "M" + "7","majorDimension": "ROWS","values": [['=СЧЁТЕСЛИМН(E:E; "пред")']]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "L" + "8","majorDimension": "ROWS","values": [["Киков"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "M" + "8","majorDimension": "ROWS","values": [['=СЧЁТЕСЛИМН(E:E; "кик")']]}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"requests":[{"repeatCell":{"cell":{"userEnteredFormat":{
                                                            "backgroundColor": {"red" : color['red'], "green" : color['green'], "blue" : color['blue'],"alpha": 1},}},
                                                            "range":{"sheetId": sheetId,"startRowIndex": 0,"endRowIndex": 9000,"startColumnIndex": 0,"endColumnIndex": 2},"fields": "userEnteredFormat"}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"requests":[{"repeatCell":{"cell":{"userEnteredFormat":{
                                                            "backgroundColor": {"red" : color['red'], "green" : color['green'], "blue" : color['blue'],"alpha": 1},}},
                                                            "range":{"sheetId": sheetId,"startRowIndex": 0,"endRowIndex": 9000,"startColumnIndex": 8,"endColumnIndex": 10},"fields": "userEnteredFormat"}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"requests":[{"repeatCell":{"cell":{"userEnteredFormat":{
                                                            "backgroundColor": {"red" : color['red'], "green" : color['green'], "blue" : color['blue'],"alpha": 1},}},
                                                            "range":{"sheetId": sheetId,"startRowIndex": 0,"endRowIndex": 1,"startColumnIndex": 0,"endColumnIndex": 10},"fields": "userEnteredFormat"}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"requests": [{'mergeCells': 
                                                            {'range': {'sheetId': sheetId,'startRowIndex': 0,'endRowIndex': 1,'startColumnIndex': 0,'endColumnIndex': 10},'mergeType': 'MERGE_ALL'}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"requests":[{"repeatCell":{"cell":{"userEnteredFormat":{
                                                            "backgroundColor": {"red" : color_stats['red'], "green" : color_stats['green'], "blue" : color_stats['blue'],"alpha": 1},}},
                                                            "range":{"sheetId": sheetId,"startRowIndex": 2,"endRowIndex": 8,"startColumnIndex": 11,"endColumnIndex": 13},"fields": "userEnteredFormat"}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={"requests": [{"updateDimensionProperties": {
                                                            "range": {"sheetId": sheetId,"dimension": "COLUMNS","startIndex": 0,"endIndex": 2},
                                                            "properties": {"pixelSize": 10},"fields": "pixelSize"}}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={"requests": [{"updateDimensionProperties": {
                                                            "range": {"sheetId": sheetId,"dimension": "COLUMNS","startIndex": 8,"endIndex": 10},
                                                            "properties": {"pixelSize": 10},"fields": "pixelSize"}}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "M" + "8","majorDimension": "ROWS","values": [['=СЧЁТЕСЛИМН(E:E; "кик")']]}]}).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={"requests":[{'repeatCell': {
                                                            'range': {'startRowIndex': 0,'endRowIndex': 4,'startColumnIndex': 12,'endColumnIndex': 14,},
                                                            'cell': {'userEnteredFormat': {'numberFormat': {'type': 'TIME','pattern': 'h:mm:ss',},},},'fields': 'userEnteredFormat.numberFormat',}}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "L" + "13","majorDimension": "ROWS","values": [["made by VAFI"]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": list_name + "!" + "L" + "14","majorDimension": "ROWS","values": [["https://vk.com/outsider77"]]}]}).execute()

if not os.path.exists("key.json"):
    print("Key for google sheets not found")
    error = input()
    sys.exit()

if not os.path.exists("ok.mp3"):
    print("Если вы удалили ok.mp3 - не надо так делать. Просто уменьшите громкость в микшере")
    print("Если нет - сообщите разработчику")
    sys.exit()

try:
    if not os.path.exists("settings.cfg"):
        settins_configurator()
    if not os.path.exists("account_info.cfg"):
        account_configurator()    
    account_file = open('account_info.cfg', 'r', encoding='utf8')
    account = eval(account_file.read())
    account_file.close()
    settings_file = open('settings.cfg', 'r', encoding='utf8')
    settings = eval(settings_file.read())
    settings_file.close()

    print('url: ' + 'https://docs.google.com/spreadsheets/d/' + account['tableID'])
    print('\nCtrl + ' + settings['script_hotkey'] +' - записать нарушение \nCtrl + ' + settings['second_screenshot_hotkey'] +' - добавить скриншот к нарушению \nCtrl + J - записать время выхода')
    spreadsheetId = account['tableID']
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
        if settings['sound_activate'] == True:
            playsound('ok.mp3', False)
        else:
            return True


    def screenshot(resolX, resolY):
        keyboard.send(str(settings['chat_hotkey']).lower())
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
        keyboard.send(str(settings['chat_hotkey']).lower())
        sleep(0.1)
        keyboard.send('UP')
        sleep(0.1)
        keyboard.press('Ctrl')
        sleep(0.1)
        keyboard.release('Ctrl')
        keyboard.press('Ctrl')
        keyboard.send('a')
        sleep(0.1)
        keyboard.release('Ctrl')
        keyboard.press('Ctrl')
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

        screenshot_link = screenshot(settings['resolutionX'], settings['resolutionY'])

        count = account['count'] + 1
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": account['listName'] + "!" + "G" + str(count),"majorDimension": "ROWS","values": [[screenshot_link]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": account['listName'] + "!" + "E" + str(count),"majorDimension": "ROWS","values": [[type_violation]]}]}).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": account['listName'] + "!" + "F" + str(count),"majorDimension": "ROWS","values": [[nickname]]}]}).execute()
        play_sound()
        account['count'] = account['count'] + 1
        print(type_violation + ' ' + nickname + ' ' + screenshot_link)
        conf = open('account_info.cfg', 'w', encoding='utf8')
        conf.write(str(account))
        conf.close()


    def time_join():
        count = account['count'] + 1
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": account['listName'] + "!" + "C" + str(count),"majorDimension": "ROWS","values": [[datetime.datetime.now().strftime("%H:%M:%S")]]}]}).execute()
        print("Время захода учтено")
        
        account['count'] = account['count'] + 1
        conf = open('account_info.cfg', 'w', encoding='utf8')
        conf.write(str(account))
        conf.close()


    def time_exit():
        count = account['count'] + 1
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": account['listName'] + "!" + "D" + str(count),"majorDimension": "ROWS","values": [[datetime.datetime.now().strftime("%H:%M:%S")]]}]}).execute()
        print("Время выхода учтено")
        account['count'] = account['count'] + 1
        conf = open('account_info.cfg', 'w', encoding='utf8')
        conf.write(str(account))
        conf.close()
        sys.exit()

    def second_screenshot():
        screenshot_link = screenshot(settings['resolutionX'], settings['resolutionY'])
        count = account['count']
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],body={"valueInputOption": "USER_ENTERED","data": [
                                                            {"range": account['listName'] + "!" + "H" + str(count),"majorDimension": "ROWS","values": [[screenshot_link]]}]}).execute()
        play_sound()
        print("Добавлен скриншот")

    time_join()
    while True:
        keyboard.add_hotkey('Ctrl + ' + settings['script_hotkey'], lambda: detect_violation())
        keyboard.add_hotkey('Ctrl + ' + settings['second_screenshot_hotkey'], lambda: second_screenshot())
        keyboard.add_hotkey('Ctrl + j', lambda: time_exit())
        keyboard.wait('Ctrl+alt+l')

except:
    print("О нет! Произошла ошибка! Скиньте текст ниже https://vk.com/outsider77")
    error = input()
    exit()
