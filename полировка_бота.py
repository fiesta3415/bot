import requests
import openpyxl
import vk_api, vk
from vk_api.keyboard import VkKeyboard, VkKeyboardColor
from vk_api.utils import get_random_id



dls = "http://serp-koll.ru/files/rasp_zan_1%20korp.xlsx"
resp = requests.get(dls) #скачивает файл в папку с кодом
with open('rasp.xlsx', 'wb') as output:
    output.write(resp.content)

    wb = openpyxl.reader.excel.load_workbook(r'rasp.xlsx')
wb.active = 0

ws = wb.active

sheet = wb.active
one = sheet['E7'].value
two = sheet['E9'].value
three = sheet['E11'].value
four = sheet['E13'].value

vk_session = vk_api.VkApi(token='e018a437406f88804ca1ee80074588533acb03be57e4fb1713a41e4431e8a9e4d46248319667e9c78cb81')
from vk_api.bot_longpoll import VkBotLongPoll, VkBotEventType
longpoll = VkBotLongPoll(vk_session, 202353833)
vk = vk_session.get_api()
from vk_api.longpoll import VkLongPoll, VkEventType
for event in longpoll.listen():
    if event.type == VkBotEventType.MESSAGE_NEW:
        if 'расписание' in str(event) or 'бобс' in str(event) or 'да' in str(event):
            if event.from_chat:
    
                vk.messages.send(
                key = ('aba27a803d4511153f1b6ac52674b15b3ec3f018'),
                server = ('https://lp.vk.com/wh202353833'),
                ts=('1'),
                random_id = get_random_id(),
                message=('Расписание:' + "\n" + "1." + one + "\n" + "2." + two +"\n" + "3." + three + "\n" + "4." + four),
                chat_id = event.chat_id
                )