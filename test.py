import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
session = outlook.GetNamespace('MAPI')

#6 is the inbox will figure out how to get a list of folders later
inbox = session.GetDefaultFolder(6) 
#3 is deleted items
#6 is inbox
#4 is outbox
#5 is sent items


print inbox.Folders[0]
print session.Folders[0]


print inbox


for counter in range(0,len(inbox.Items)):
	print inbox.Items[counter]
	print inbox.Items[counter].body


address_lists = session.AddressLists
gal = address_lists.Item ("Global Address List")
print gal.AddressEntries.Item.Count
#number of attachtments
print inbox.Items[2].Attachments.Count

#name of attachment
print inbox.Items[2].Attachments[0] 

print inbox.Items[2].Attachments[0].FileName

#inbox.Items[2].Attachments[0].SaveAsFile("a")
#SaveAsFile i think is how you save a file


print inbox.Items[2].Attachments[0].SaveAsFile('c:\temp\test.docx')

#(will save to file location default)

#can be changed have to figure out how

#SaveEmailAttachtmentToFolder


