import logging as log
from datetime import date
from email.message import EmailMessage

import pandas as pd
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

import config
import io
import smtplib

# RETRIEVE USER CREDENTIALS
user = config.USERNAME
password = config.PASSWORD


# CHECK EXPIRATION HELPER
def convert_days_remaining(expire_date):
    return (expire_date.date() - date.today()).days


# CHECK PASSWORD DOCUMENT FOR UPCOMING EXPIRATIONS
def check_passwords():
    notify_list = []

    url = "https://pytestdev.sharepoint.com/sites/team"
    path = "/sites/team/Shared%20Documents/files/data-set.xlsx"

    ctx_auth = AuthenticationContext(url)
    if ctx_auth.acquire_token_for_user(user, password):
        ctx = ClientContext(url, ctx_auth)
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        print("Authentication successful")

    response = File.open_binary(ctx, path)

    file_object = io.BytesIO()
    file_object.write(response.content)
    file_object.seek(0)  # set file object to start

    data = pd.read_excel(file_object)

    for index, row in data.iterrows():
        time_to_expiration = convert_days_remaining(row[0])
        if 45 >= time_to_expiration > 0:
            notify_list.append(row)

    return notify_list


def send_email(notification_list):
    s = smtplib.SMTP('smtp.outlook.com', 587)

    s.starttls()
    s.login(user, password)

    for record in notification_list:
        # CAPTURE NOTIFICATION INFORMATION FROM RECORD
        expiration_date = str(record[0].date())
        expire_eta = str(convert_days_remaining(record[0]))
        application = record[1]
        contact = record[4]
        contact_email = record[5]

        # BUILD MESSAGE
        message = "\nGreetings " + contact + ",\n It is now time to renew the database credentials for " + \
                  application + ". \nThe current password is set to expire on " + expiration_date + " [" + \
                  expire_eta + " days]. Please renew this password ASAP. \n\n"

        # BUILD EMAIL
        email = EmailMessage()
        email['Subject'] = '***** PASSWORD EXPIRATION NOTICE FOR: ' + application.upper() + ' *****'
        email['From'] = config.USERNAME
        email['To'] = contact_email
        email['X-Priority'] = '1'
        email.set_content(message)

        # OUTPUT LOG
        log.info('APPLICATION: ' + application + '\n EXPIRES: ' + expire_eta + ' DAYS -- [' + expiration_date + ']' +
                 '\n CONTACT: ' + contact + '\n EMAIL: ' + contact_email + '\n')

        s.send_message(email)

    print("Emails sent successfully!")
    log.info("Emails sent successfully!")
    s.quit()


# EXECUTION
expired_list = check_passwords()

send_email(expired_list)
