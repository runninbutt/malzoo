"""
File is part of Malzoo

Exchange supplier. Attempts to get all e-mails from the account defined in the configuration and 
add them to the exchangeworker queue. Supports O365.
"""
import re
import os
import sys # temporary to catch keyboard error
import time
import email
from exchangelib import Credentials, Configuration, Account, DELEGATE
import ConfigParser
from multiprocessing import Process, Queue

#Parent
from malzoo.common.abstract import Supplier

class Exchange(Supplier):
    def open_connection(self, verbose=False):
        """
        Setup the Exchange connection
        """
        #Read configuration
        hostname = self.conf.get('exchange', 'server')
        username = self.conf.get('exchange', 'username')
        password = self.conf.get('exchange', 'password')
        # Create a Credentials object
        credentials = Credentials(username, password)
        # Connect to Exchange
        config = Configuration(server=hostname, credentials=credentials)
        account = Account(username, config=config, autodiscover=False, access_type=DELEGATE)
        if verbose: 
            self.log('{0} - {1} - {2} '.format('exchangesupplier','open_connect','trying to connect'))
        return account
    
    def list_mailboxes(self):
        """ Retrieve all folders and subfolders """
        try:
            data = self.c.root
        except Exception as e:
            self.log('{0} - {1} - {2} '.format('exchangesupplier','list_mailboxes',e))
        finally:
            return data
    
    def get_ids(self, mailbox):
        """ 
        Get all messages from the given folder and return the ID numbers 
        """
        c = self.c
        msg_ids = []
        try:
            response, data = c.root / mailbox
            num_msgs = int(data[0])
            response, msg_ids = c.search(None, '(UNSEEN)')
        except Exception as e:
            self.log('{0} - {1} - {2} '.format('imapsupplier','get_ids',e))
        finally:
            return msg_ids
    
    def fetch_mail(self, mailbox, batch_ids):
        """ 
        For each mail ID, fetch the whole email and parse it with msgparser 
        """
        try:
            c = self.c
            response, data = c.select(mailbox)
            response, messages = c.fetch(batch_ids,"(RFC822)")
        except Exception as e:
            messages = None
            self.log('{0} - {1} - {2} '.format('imapsupplier','fetch_mail',e))
        finally:
            return messages
    
    def copy_message(self, mailbox, ID, target):
        """ Copy a message to the target folder """
        success = False
        try:
            c = self.c
            response, data = c.select(mailbox)
            c.copy(ID,target)
            success = True
        finally:
            return success
        
    def move_message(self, mailbox, ID, target):
        """ Move a message to the target folder """
        success = False
        try:
            c = self.c
            response, data = c.select(mailbox)
            c.copy(ID,target)
            c.store(ID, '+FLAGS', '\\Deleted')
            success = True
        finally:
            return success

    def markAsRead(self, msg):
        msg.is_read = True
        msg.save()
        return msg

    def run(self, mail_q):
        """ Start the process """
        self.conf = ConfigParser.ConfigParser()
        self.conf.read('config/malzoo.conf')
                 
        folderName = self.conf.get('exchange','folder')

        while True:
            try:
                self.c = self.open_connection()
                connected = True
            except Exception as e:
                self.log('{0} - {1} - {2} '.format('exchangesupplier','run',e))
                connected = False
            
            if connected:
                folder = None
                try:
                    # 1.Get items in the defined folder
                    if folderName.lower() == 'inbox':
                        folder = self.c.inbox
                    else:
                        for f in self.c.root.walk():
                            if f.name == folderName:
                                folder = f
                    if folder is None:
                        raise ValueError('folder %s not found', folderName)

                    # Fetch all unread items
                    unread = folder.filter(is_read = False)
                    # 2. Handle unread emails
                    for msg in unread:
                        #Put email messages to the handling queue
                        exchange_q.put(msg)
            # 2. Get ID's from new messages in mailbox INBOX
                    email_ids   = self.get_ids(mailbox)     
                    if email_ids[0] != '': 
                        batch_ids   = ','.join(email_ids[0].split())
            # 3. Get the messages from defined mailbox and the list of ID's
                        emails      = self.fetch_mail(mailbox, batch_ids)
                        if emails != None:
                            for i in range(0,len(emails),2):
                                email = {'filename':emails[i][1],'tag':self.conf.get('settings','tag')}
                                mail_q.put(email)
                        else:
                            self.log('{0} - {1} - {2} '.format('imapsupplier','run','no emails'))
    
            # 4. Check again in 10 seconds
                    self.c.close()
                    self.c.logout()
                    time.sleep(60) 
    
                except KeyboardInterrupt:
                    sys.exit(0)
            else:
                time.sleep(120)
