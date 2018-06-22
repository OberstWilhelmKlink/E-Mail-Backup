#!/usr/bin/env python3

__author__ = "Georg Vogl"
__license__ = "Public Domain"
__version__ = "1.0"

import sqlite3
import getpass
import imaplib
import email
import os
import sys
import base64
from pathlib import Path

SQLITE_DB = 'MailBackups_2.db'
BASE_DIR = ''
ADDRESSES = []
FOLDERS = []


class MailAddress:
    type = ''
    server = ''
    user = ''
    pw = ''
    id = 0

    def __init__(self, _type, _server, _user, _pw, _id):
        self.type = _type
        self.server = _server
        self.user = _user
        self.pw = _pw
        self.id = _id

    def get_type(self):
        return self.type

    def get_server(self):
        return self.server

    def get_pw(self):
        return self.pw

    def get_id(self):
        return self.id

    def get_user(self):
        return self.user

    def export_as_tuple(self):
        return [self.type, self.server, self.user, self.pw]


def run_application():
    check_for_database()
    get_system_data()
    get_mail_addresses(True)
    process_mails()
    return 0


def main():
    args = sys.argv
    if len(args) == 1:
        main_menu()
    if len(args) > 1:
        if args[1] == 'run':
            return run_application()
        if args[1] == 'init':
            return init_application()
        if args[1] == 'edit':
            return modify_menu()
        print('Unknown argument - use options:[run,init,edit]')
        sys.exit()


def init_application():
    global BASE_DIR
    sql_connection = None
    try:
        print('Create Database ...')
        sql_connection = sqlite3.connect(SQLITE_DB)
        cursor = sql_connection.cursor()
        print('Create Tables ...')
        cursor.execute('CREATE TABLE IF NOT EXISTS Data ( BasePath TEXT )')
        sql_connection.commit()
        cursor.execute('''CREATE TABLE IF NOT EXISTS MailLogin ( 
                            Type TEXT, 
                            Server TEXT, 
                            ID INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, 
                            User TEXT, 
                            PW TEXT )''')
        sql_connection.commit()
        while BASE_DIR == '':
            base_folder = input("Enter base folder for backup:")
            if os.path.exists(base_folder):
                BASE_DIR = base_folder
                cursor.execute('INSERT INTO Data VALUES (?)', [BASE_DIR])
                sql_connection.commit()
        print('Application is ready for use.')
        return 0
    finally:
        if sql_connection is not None:
            sql_connection.close()


def get_mail_addresses(with_password):
    try:
        ADDRESSES.clear()
        sql_connection = sqlite3.connect(SQLITE_DB)
        cursor = sql_connection.cursor()
        cursor.execute('SELECT ID,Type,Server,User,PW FROM MailLogin')
        data_rows = cursor.fetchall()
        for row in data_rows:
            if row[4] == '' and with_password:
                mail_address = MailAddress(row[1],
                                           row[2],
                                           row[3],
                                           getpass.getpass(prompt='Password for ' + row[3] + ': '),
                                           row[0])
            else:
                mail_address = MailAddress(row[1], row[2], row[3], row[4], row[0])
            if mail_address.get_pw() != '' or not with_password:
                ADDRESSES.append(mail_address)
        return 0
    except sqlite3.Error:
        return 1


def process_mails():
    for address in ADDRESSES:
        global MAIL_COUNTER
        MAIL_COUNTER = 0
        imap = None
        try:
            hostname = address.get_server()
            imap = imaplib.IMAP4_SSL(hostname)
            imap.login(address.get_user(), address.get_pw())
            resp, folder_list = imap.list('""', '*')
            if resp == 'OK':
                for folder in folder_list:
                    flags, separator, name = parse_mailbox(bytes.decode(folder))
                    FOLDERS.append(name)
        finally:
            if imap is not None:
                imap.logout()

        for folder in FOLDERS:
            process_mail(address, folder)
        return


def write_to_file(mail_dir, msg, filename):
    file_name = mail_dir + filename + '.eml'
    fw = open(file_name, 'w', encoding="utf-8")
    fw.write(msg)
    fw.close()
    return


def process_mail(address, folder_string):
    imap = None
    try:
        hostname = address.get_server()
        imap = imaplib.IMAP4_SSL(hostname)
        imap.login(address.get_user(), address.get_pw())
        print(folder_string)
        resp, count = imap.select('"' + folder_string + '"')
        if resp == 'OK':
            count = int(count[0].decode())
            if count > 0:
                resp, data = imap.uid('FETCH', '1:*', '(RFC822)')
                if resp == 'OK':
                    messages = [data[i] for i in range(0, len(data), 2)]
                    for mail in messages:
                        try:
                            msg = email.message_from_bytes(mail[1])
                            byte_msg = msg.as_bytes().decode(encoding='ISO-8859-1')
                            folder_string_prepared = imaputf7decode(folder_string)
                            folder_string_prepared = folder_string_prepared.replace('.', '/')
                            mail_dir = BASE_DIR + '/' + address.get_user() + '/' + folder_string_prepared + '/'
                            if not os.path.exists(mail_dir):
                                os.makedirs(mail_dir)
                            os.chdir(mail_dir)
                            date_array = email.utils.parsedate(msg['Date'])
                            date_name = str(date_array[0]) + str(date_array[1]) + str(date_array[2]) + '_' + \
                                        str(date_array[3]) + str(date_array[4]) + str(date_array[5])
                            from_mail = email.utils.parseaddr(msg['From'])[1]
                            write_to_file(mail_dir, byte_msg, date_name + '_' + from_mail)
                            global MAIL_COUNTER
                            MAIL_COUNTER += 1
                            print('Mails saved: ' + str(MAIL_COUNTER))
                        except TypeError:
                            break
    finally:
        if imap is not None:
            imap.logout()
    return


def get_system_data():
    global BASE_DIR
    try:
        sql_connection = sqlite3.connect(SQLITE_DB)
        cursor = sql_connection.cursor()
        cursor.execute('select * from Data')
        sql_connection.commit()
        data_rows = cursor.fetchall()
        BASE_DIR = data_rows[0][0]
        cursor.close()
        return 0
    except sqlite3.Error:
        return 1


def parse_mailbox(data):
    flag_end_index = data.find(')', 1)
    flags = data[1:flag_end_index]
    a, b, c = data[flag_end_index + 1:].partition(' ')
    separator, b, name = c.partition(' ')
    return flags, separator.replace('"', ''), name.replace('"', '')


def b64padanddecode(b):
    """Decode unpadded base64 data"""
    b += (-len(b) % 4)*'='  # base64 padding (if adds '===', no valid padding anyway)
    return base64.b64decode(b, altchars='+,', validate=True).decode('utf-16-be')


def imaputf7decode(s):
    """Decode a string encoded according to RFC2060 aka IMAP UTF7.
    Minimal validation of input, only works with trusted data"""
    lst = s.split('&')
    out = lst[0]
    for e in lst[1:]:
        u, a = e.split('-', 1)  # u: utf16 between & and 1st -, a: ASCII chars following it
        if u == '':
            out += '&'
        else:
            out += b64padanddecode(u)
        out += a
    return out


def modify_menu():
    if check_for_database() > 0:
        print('The application database have to be initialized.\nRun init!')
        sys.exit()
    os.system('cls' if os.name == 'nt' else 'clear')
    menu_input = -1
    while menu_input not in range(0, 5):
        print('''What do you want to do?       
            0 - Exit
            1 - Add E-Mail
            2 - Edit E-Mail
            3 - Delete E-Mail
            4 - Set Base directory for backup
            ''')
        menu_input = input('Choose: ')
        menu_input = int(menu_input)
    if menu_input == 0:
        sys.exit()
    if menu_input == 1:
        return add_mail()
    if menu_input == 2:
        return edit_mail()
    if menu_input == 3:
        return remove_mail()
    if menu_input == 4:
        return set_base_dir()


def main_menu():
    menu_input = -1
    os.system('cls' if os.name == 'nt' else 'clear')
    while menu_input not in range(0, 4):
        print('''What do you want to do?\n       
                0 - Exit
                1 - Init Application
                2 - Run Application
                3 - Modify Application
                ''')
        menu_input = input('Choose: ')
        menu_input = int(menu_input)
    if menu_input == 0:
        sys.exit()
    if menu_input == 1:
        return init_application()
    if menu_input == 2:
        return init_application()
    if menu_input == 3:
        return modify_menu()


def add_mail():
    os.system('cls' if os.name == 'nt' else 'clear')
    protocol = ''
    server = ''
    user = ''
    while protocol not in ['IMAP', 'POP3']:
        protocol = input('Enter protocol (IMAP or POP3):')
    while server == '':
        server = input('Enter mail-server:')
    while user == '':
        user = input('Enter user:')
    pw = getpass.getpass(prompt='Enter password (optional - not encrypted saved):')
    if pw == '':
        print('Password not saved. Always you run the application, you have to enter it.')
    mail = MailAddress(protocol, server, user, pw, 0)
    if pw == '':
        pw = getpass.getpass(prompt='Enter password to check the mail and server:')
    imap = imaplib.IMAP4_SSL(server)
    res, dat = imap.login(user, pw)
    if res != 'OK':
        print('Connection error!')
        sys.exit()
    ADDRESSES.append(mail)
    sql_connection = None
    try:
        sql_connection = sqlite3.connect(SQLITE_DB)
        cursor = sql_connection.cursor()
        cursor.execute('INSERT INTO MailLogin (Type,Server,User,PW ) VALUES (?,?,?,?)', mail.export_as_tuple())
        sql_connection.commit()
        print('E-Mail fpr user {} successfully added.', mail.get_user())
    finally:
        sql_connection.close()
    return 0


def remove_mail():
    get_mail_addresses(False)
    print('Choose the address to delete:')
    ids = []
    print('0 - CANCEL')
    for address in ADDRESSES:
        print('{} - Server:{} User:{}'.format(address.get_id(), address.get_server(), address.get_user()))
        ids.append(address.get_id())
    ids.append(0)
    selected_address = -1
    while selected_address not in ids:
        selected_address = input('Select:')
        if selected_address.isnumeric():
            selected_address = int(selected_address)
    if selected_address == 0:
        sys.exit()
    remove_address = None
    for address in ADDRESSES:
        if address.get_id() == selected_address:
            remove_address = address
            break
    ADDRESSES.remove(remove_address)
    sql_connection = None
    try:
        sql_connection = sqlite3.connect(SQLITE_DB)
        cursor = sql_connection.cursor()
        cursor.execute('DELETE FROM MailLogin WHERE ID={}'.format(remove_address.get_id()))
        sql_connection.commit()
        print('E-Mail with ID {1} for user {2} successfully removed.', remove_address.get_id(), remove_address.get_user())
    finally:
        sql_connection.close()

    return 0


def edit_mail():
    # TODO Ändern von bestehender Mail-Adresse ermöglichen
    return 0


def set_base_dir():
    global BASE_DIR
    sql_connection = None
    try:
        sql_connection = sqlite3.connect(SQLITE_DB)
        cursor = sql_connection.cursor()
        while BASE_DIR == '':
            folder = input("Enter base folder for backup:")
            if os.path.exists(folder):
                BASE_DIR = folder
                cursor.execute('UPDATE Data SET BasePath= ?', [BASE_DIR])
                sql_connection.commit()
    except sqlite3.Error:
        return 1
    finally:
        sql_connection.close()
    return 0


def check_for_database():
    path_to_db = Path(SQLITE_DB)
    if path_to_db.is_file():
        sql_connection = None
        try:
            sql_connection = sqlite3.connect(SQLITE_DB)
            cursor = sql_connection.cursor()
            cursor.execute('select Count(*) from Data')
            sql_connection.commit()
            cursor.execute('select Count(*) from MailLogin')
            sql_connection.commit()
            result = 0
        except sqlite3.Error:
            result = 1
        finally:
            sql_connection.close()
    else:
        return 1
    return result


if __name__ == '__main__':
    main()