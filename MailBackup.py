#!/usr/bin/env python3

__author__ = "Georg Vogl"
__license__ = "Public Domain"
__version__ = "1.0"

import sqlite3
import getpass
import imaplib
import email
import os
import base64

SQLITE_DB = 'MailBackups.db'
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


def main():
    getsystemdata()
    getmailadresses()
    processmails()
    return


def getmailadresses():
    ADRESSES.clear()
    sqlconnection = sqlite3.connect(SQLITE_DB)
    cursor = sqlconnection.cursor()
    cursor.execute('SELECT ID,Type,Server,User,PW FROM MailLogins')
    datarows = cursor.fetchall()
    for row in datarows:
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
    return


def processmails():
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
            processmail(adress, folder)
        return


def writetofile(mailDir, msg, filename):
    file_name = mailDir + filename + '.eml'
    fw = open(file_name,'w', encoding="utf-8")
    fw.write(msg)
    fw.close()
    return


def processmail(adress, folderstring):
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
                            writetofile(mailDir, smsg, datename + '_' + from_mail)
                            global MAILCOUNTER
                            MAILCOUNTER += 1
                            print('Mails verarbeitet: ' + str(MAILCOUNTER))
                        except TypeError:
                            break
    finally:
        if imap != None:
            imap.logout()
    return


def getsystemdata():
    sqlconnection = sqlite3.connect(SQLITE_DB)
    cursor = sqlconnection.cursor()
    cursor.execute('select * from Data')
    datarows = cursor.fetchall()
    BASE_DIR = datarows[0][0]
    cursor.close()


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


if __name__ == '__main__':
   main()


