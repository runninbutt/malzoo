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

    def run(self, exchange_q):
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

                    # 2. Fetch all unread items
                    unread = folder.filter(is_read = False)
                    # 3. Handle unread emails starting from the oldest
                    for msg in reversed(unread):
                        #Put email messages to the handling queue
                        exchange_q.put(msg)
                        #Set message to read so it won't be processed again
                        readmsg = self.markAsRead(msg)
    
            # 4. Check again in 60 seconds
                    time.sleep(60) 
    
                except KeyboardInterrupt:
                    sys.exit(0)
            else:
                time.sleep(120)
