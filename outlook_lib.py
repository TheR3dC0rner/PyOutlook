import win32com.client


def count_subfolders(folder):
    return len(folder.Folders)

def print_subfolders(folder,start,end):
    for counter in range(start,end):
        print folder.Folders[counter]
        
def get_subfolders(folder,start,end):
         list = []
         count = 0
         for counter in range(start,end):
             list[count] = folder.Folders[counter]
             count = count + 1
         return list

def count_messages(folder):
    return len(folder.Items)

def print_subjects(folder,start,end):
    for counter in range(start,end):
        print inbox.Items[counter]

def get_subjects(folder,start,end):
         list = []
         count = 0
         for counter in range(start,end):
             list[count] = folder.Items[counter]
             count = count + 1
         return list

def print_body(start,end,folder):
    for counter in range(start,end):
        print inbox.Items[counter].body
        
def get_subjects(folder,start,end):
         list = []
         count = 0
         for counter in range(start,end):
             list[count] = folder.Items[counter].body
             count = count + 1
         return list        

def print_emailaddress(session):
    print session.Folder[0]
    
def get_emailaddress(session):
    print sessions.Folder[0]

def getattachment_count(message):
    return Attachments.Count

def print_filenames(message):
    for counter in range(start,end):
        print inbox.Items[counter]


def init_session():
    outlook = win32com.client.Dispatch("Outlook.Application")
    session = outlook.GetNamespace('MAPI')
    return session
#3 is deleted items
#4 is outbox
#5 is sent item
#6 is inbox
        
#this is the initialization
#outlook = win32com.client.Dispatch("Outlook.Application")
#session = outlook.GetNamespace('MAPI')

