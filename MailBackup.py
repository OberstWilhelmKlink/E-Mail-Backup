#!/usr/bin/env python3

__author__ = "Georg Vogl"
__license__ = "Public Domain"
__version__ = "1.0"

import sqlite3, getpass
import imaplib
import email
import os
import sys
import base64
from pathlib import Path

SQLITE_DB = 'MailBackups_2.db'
BASE_DIR = ''
ADRESSES = []
FOLDERS = []


class MailAdress():
    type = ''
    server = ''
    user = ''
    pw = ''
    id = 0

    def __init__(self,_type,_server,_user,_pw,_id):
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
        return self.pw

    def get_user(self):
        return self.user


def run_application():
    check_for_database()
    get_system_data()
    get_mailadresses()
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
    try:
        print('Create Database ...')
        sqlconnection = sqlite3.connect(SQLITE_DB)
        cursor = sqlconnection.cursor()
        print('Create Tables ...')
        cursor.execute('CREATE TABLE IF NOT EXISTS Data ( BasePath TEXT )')
        sqlconnection.commit()
        cursor.execute('''CREATE TABLE IF NOT EXISTS MailLogins ( 
                            Type TEXT, 
                            Server TEXT, 
                            ID INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, 
                            User TEXT, 
                            PW TEXT )''')
        sqlconnection.commit()
        while BASE_DIR == '':
            dir = input("Enter base folder for backup:")
            if os.path.exists(dir):
                BASE_DIR = dir
                cursor.execute('INSERT INTO Data VALUES (?)', [BASE_DIR])
                sqlconnection.commit()
        print('Application is ready for use.')
        return 0
    finally:
        if sqlconnection != None:
            sqlconnection.close()


def get_mailadresses():
    try:
        ADRESSES.clear()
        sqlconnection = sqlite3.connect(SQLITE_DB)
        cursor = sqlconnection.cursor()
        cursor.execute('SELECT ID,Type,Server,User,PW FROM MailLogins')
        data_rows = cursor.fetchall()
        for row in data_rows:
            if row[4] == '':
                mailadress = MailAdress(row[1],
                                        row[2],
                                        row[3],
                                        getpass.getpass(prompt='Passwort für ' + row[3] + ': '),
                                        row[0])
            else:
                mailadress = MailAdress(row[1], row[2], row[3], row[4], row[0])
            if mailadress.get_pw() != '':
                ADRESSES.append(mailadress)
        return 0
    except sqlite3.Error as er:
        return 1



def process_mails():
    for adress in ADRESSES:
        global MAILCOUNTER
        MAILCOUNTER = 0
        try:
            hostname = adress.get_server()
            imap = imaplib.IMAP4_SSL(hostname)
            imap.login(adress.get_user(), adress.get_pw())
            resp, folderlist = imap.list('""', '*')
            if resp == 'OK':
                for folder in folderlist:
                    flags, separator, name = parse_mailbox(bytes.decode(folder))
                    FOLDERS.append(name)
        finally:
            if imap != None:
                imap.logout()

        for folder in FOLDERS:
            process_mail(adress, folder)
        return


def write_to_file(mailDir, msg, filename):
    file_name = mailDir + filename + '.eml'
    fw = open(file_name,'w', encoding="utf-8")
    fw.write(msg)
    fw.close()
    return


def process_mail(adress, folderstring):
    try:
        hostname = adress.get_server()
        imap = imaplib.IMAP4_SSL(hostname)
        imap.login(adress.get_user(), adress.get_pw())
        print(folderstring)
        resp, count = imap.select('"' + folderstring + '"')
        if resp == 'OK':
            count = int(count[0].decode())
            if count > 0:
                resp, data = imap.uid('FETCH', '1:*', '(RFC822)')
                if resp == 'OK':
                    messages = [data[i] for i in range(0, len(data), 2)]
                    for mail in messages:
                        try:
                            msg = email.message_from_bytes(mail[1])
                            smsg = msg.as_bytes().decode(encoding='ISO-8859-1')
                            folderstring_prepared = imaputf7decode(folderstring)
                            folderstring_prepared = folderstring_prepared.replace('.', '/')
                            mailDir = BASE_DIR + '/' + adress.get_user() + '/' + folderstring_prepared + '/'
                            if not os.path.exists(mailDir):
                                os.makedirs(mailDir)
                            os.chdir(mailDir)
                            datearray = email.utils.parsedate(msg['Date'])
                            datename = str(datearray[0]) + str(datearray[1]) + str(datearray[2]) + '_' + str(datearray[3]) + str(datearray[4]) + str(datearray[5])
                            from_mail = email.utils.parseaddr(msg['From'])[1]
                            write_to_file(mailDir, smsg, datename + '_' + from_mail)
                            global MAILCOUNTER
                            MAILCOUNTER += 1
                            print('Mails verarbeitet: ' + str(MAILCOUNTER))
                        except TypeError:
                            break
    finally:
        if imap != None:
            imap.logout()
    return


def get_system_data():
    global BASE_DIR
    try:
        sqlconnection = sqlite3.connect(SQLITE_DB)
        cursor = sqlconnection.cursor()
        cursor.execute('select * from Data')
        datarows = cursor.fetchall()
        BASE_DIR = datarows[0][0]
        cursor.close()
        return 0
    except sqlite3.Error as er:
        return 1


def parse_mailbox(data):
    flag_end_index = data.find(')', 1)
    flags = data[1:flag_end_index]
    a, b, c = data[flag_end_index + 1:].partition(' ')
    separator, b, name = c.partition(' ')
    return (flags, separator.replace('"', ''), name.replace('"', ''))


def b64padanddecode(b):
    """Decode unpadded base64 data"""
    b+=(-len(b)%4)*'=' #base64 padding (if adds '===', no valid padding anyway)
    return base64.b64decode(b,altchars='+,',validate=True).decode('utf-16-be')


def imaputf7decode(s):
    """Decode a string encoded according to RFC2060 aka IMAP UTF7.
    Minimal validation of input, only works with trusted data"""
    lst=s.split('&')
    out=lst[0]
    for e in lst[1:]:
        u,a=e.split('-',1) #u: utf16 between & and 1st -, a: ASCII chars folowing it
        if u=='' : out+='&'
        else: out+=b64padanddecode(u)
        out+=a
    return out


def modify_menu():
    if check_for_database() > 0:
        print('The application database have to be initialized.\nRun init!')
        sys.exit()
    menu_input = -1
    while menu_input not in range(0,5):
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
    #TODO hinzufügen von Mail Adressen ermöglichen (inkl. Test)
    return 0


def remove_mail():
    #TODO Löschen von Mail-Adresse umsetzen
    return 0


def edit_mail():
    #TODO Ändern von bestehender Mail-Adresse ermöglichen
    return 0


def set_base_dir():
    global BASE_DIR
    try:
        sqlconnection = sqlite3.connect(SQLITE_DB)
        cursor = sqlconnection.cursor()
        while BASE_DIR == '':
            dir = input("Enter base folder for backup:")
            if os.path.exists(dir):
                BASE_DIR = dir
                cursor.execute('UPDATE Data SET BasePath= ?', [BASE_DIR])
                sqlconnection.commit()
    except sqlite3.Error as er:
        return 1
    finally:
        sqlconnection.close()

    return 0


def check_for_database():
    path_to_db = Path(SQLITE_DB)
    if path_to_db.is_file():
        try:
            sqlconnection = sqlite3.connect(SQLITE_DB)
            cursor = sqlconnection.cursor()
            cursor.execute('select Count(*) from Data')
            sqlconnection.commit()
            cursor.execute('select Count(*) from MailLogins')
            sqlconnection.commit()
            result = 0
        except sqlite3.Error as er:
            result = 1
        finally:
            sqlconnection.close()
    else:
        return 1
    return result


if __name__ == '__main__':
   main()


