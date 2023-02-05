import os
import time
import re
import comtypes
import requests
import telebot
from comtypes.client import CreateObject

# Reading a bot token and chat ID from a file:

BOT_AND_CHAT_DATA = open('bot_and_chat_data.txt', 'r')
BOT_TOKEN: str = BOT_AND_CHAT_DATA.readline().strip()
CHAT_ID: str = BOT_AND_CHAT_DATA.readline().strip()
BOT_AND_CHAT_DATA.close()

# Create objects for speech synthesis and write it
# to a file, create a file to which the generated
# speech will be recorded, and create an instance
# of the TeleBot class:

FILE_ENGINE = CreateObject("SAPI.SpVoice")
FILE_STREAM = CreateObject("SAPI.SpFileStream")
OUTFILE: str = "message.mp3"
BOT = telebot.TeleBot(BOT_TOKEN)


# Implementation of the function of receiving
# a messages from the Telegram channel
# (by sending a request):

def get_updates(BOT_TOKEN: str, offset=0):
    result = requests.get(f'https://api.telegram.org/bot{str(BOT_TOKEN)}/getUpdates?offset={str(offset)}').json()
    return result['result']


# Implementing the function of finding
# English characters in a string
# (checks whether the string contains
# the characters 'a'-'z', 'A'-'Z' and '0'-'9'):

def validate_english_text(message: str):
    for i in range(97, 123):
        if message.find(chr(i)) != -1:
            return True
    for i in range(65, 91):
        if message.find(chr(i)) != -1:
            return True
    for i in range(48, 58):
        if message.find(chr(i)) != -1:
            return True
    return False

# Implementation of the function of
# sending a message by a bot
# (sends a request to send a message):

def telegrambot_sendmsg(BOT_TOKEN: str, CHAT_ID: str, message: list):
    url = f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage?chat_id={CHAT_ID}&parse_mode=Markdown&text={message}'
    try:
        response = requests.get(url)
        return response.json()
    except ConnectionError:
        return 0

telegrambot_sendmsg(BOT_TOKEN, CHAT_ID, "Bot has been started")


# Implementation of the function of receiving
# the last message from the Telegram channel,
# if possible, its voiceover and sending the
# voiced form to the Telegram channel:

def scan_messages(BOT_TOKEN: str, CHAT_ID: str):
    update_id: int = get_updates(BOT_TOKEN)[-1]['update_id']  # Assign the ID of the last sent message to the bot
    while True:
        time.sleep(0.5)
        messages: list = get_updates(BOT_TOKEN, update_id)  # Getting updates
        for message in messages:

            # If the update has an id greater than the id of the last message a new message has arrived

            if update_id < message['update_id']:
                update_id = message['update_id']  # Assign the ID of the last sent message to the bot
                try:
                    if validate_english_text(message['message']['text']):

                        # Open .mp3 file and write a voiced message to it:

                        FILE_STREAM.Open(OUTFILE, comtypes.gen.SpeechLib.SSFMCreateForWrite)
                        FILE_ENGINE.AudioOutputStream = FILE_STREAM
                        FILE_ENGINE.speak(message['message']['text'])
                        FILE_STREAM.Close()

                        # Sending a .mp3 file to the Telegram channel:

                        sended_message = open('message.mp3', 'rb')
                        BOT.send_voice(CHAT_ID, sended_message)
                        sended_message.close()
                        os.remove('message.mp3')
                    else:
                        telegrambot_sendmsg(BOT_TOKEN, CHAT_ID, "Sorry, I can't say that")
                except KeyError:
                    continue

scan_messages(BOT_TOKEN, CHAT_ID)
