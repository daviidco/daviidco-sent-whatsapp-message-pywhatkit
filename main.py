import sys
import time

import numpy as np
import pandas as pd
import pywhatkit as kit
import pyautogui as pg


from templates import *


def get_customers():
    df = pd.read_excel('customers_bodytech.xlsx')
    df_array = np.array(df)
    return df_array


def message_body(name_contact, template):
    body = template.format(name_contact)
    return body


def send_whatsapp_message(phone, name, name_template):
    # Specify the phone number (with country code) and the message
    phone_number = phone
    message = message_body(name, dir_templates[name_template])
    # Send the message instantly using PyWhatKit
    kit.sendwhatmsg_instantly(phone_number, message)
    time.sleep(0.5)
    pg.click(x=1000, y=960)  # Adjust the coordinates based on your screen
    time.sleep(1)
    pg.press("Enter")
    time.sleep(1)
    pg.hotkey('ctrl', 'w')


if __name__ == '__main__':
    print("Process Initiated....")
    try:
        name_template = input("Type template name: ")
        print(f"{dir_templates[name_template]}")
        res_continue = input("Continue (y/n): ")
        if res_continue.lower() != "y":
            sys.exit()



    except Exception as e:
        print(f"Template not found: {name_template}")
        sys.exit()
    # Retrieve profiles from the Excel file
    customers = get_customers()

    # Iterate through each profile and send personalized messages
    for customer in customers:
        send_whatsapp_message(phone=f"{customers[2]}", name=f"{customers[0]}", name_template=name_template)
        time.sleep(1.5)

    print("Process completed....")

