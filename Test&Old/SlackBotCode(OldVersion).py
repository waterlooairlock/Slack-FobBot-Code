import re
import time
import json
import gspread
from slackclient import SlackClient
from oauth2client.service_account import ServiceAccountCredentials


name_index = 0
id_index = 1
fob_count_index = 2
list_of_values = []

slack_bot_token = "REDACTED"
slack_client = SlackClient(slack_bot_token)

#Store Slack Bot's User ID
user_list = slack_client.api_call("users.list")
for user in user_list.get('members'):
    if user.get('name') == 'fob-boy':
        slack_bot_id = user.get('id')
        break
if slack_client.rtm_connect():
    print('Connected to the beautiful web, ready to track some FOBS!')

###############################################################################
###########################__FUNCTION_DEFINITION__#############################

def get_fob_database():
    #use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds']

    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)
    
    print (client.auth )
    
    #Find a workbook by URL and open the first sheet
    sheet = client.open_by_url("REDACTED").sheet1
    
    #Extract all of the values
    list_of_values = sheet.get_all_values()
    
#------------------------------------------------------------------------------

def take_and_give(take_user_name, take_user_id, give_user_name, give_user_id):
    removed = False
    added = False
    
   
    for i in range(list_of_values.__len__()):
        if i != 0:
            sublist = list_of_values[i]
            if sublist[id_index] == take_user_id and sublist[fob_count_index] > 0:
                sublist[fob_count_index] = sublist[fob_count_index]-1
                #UPDATE SPREADSHEET ROW
                if sublist[fob_count_index] == 0:
                    message_list.append(f'<@{take_user_id}> now has no Fobs!')
                elif sublist[fob_count_index] == 1:
                    message_list.append(f'<@{take_user_id}> now has 1 Fob!')
                else:
                    message_list.append(f'<@{take_user_id}> now has {sublist[fob_count_index]} Fobs!')
                removed = True
                break
                
            elif sublist[id_index] == take_user_id and sublist[fob_count_index] == 0:
                message_list.append(f'<@{take_user_id}> does not have a Fob, one cannot be Taken!')
                removed = False
                break
                
    if not removed:
        #sheet.append_row([take_user_name,take_user_id,0])  #COMMENT BACK IN LATER
        list_of_values.append([take_user_name,take_user_id,0])
        message_list.append(f'<@{take_user_id}> does not have a Fob, one cannot be Taken!')
        
    else:
        for i in range(list_of_values.__len__()):
            if i != 0:
                sublist = list_of_values[i]
                if sublist[id_index] == give_user_id:
                    sublist[fob_count_index] = sublist[fob_count_index]+1
                    list_of_values[i] = sublist.copy()
                    #UPDATE SPREADSHEET ROW
                    if sublist[fob_count_index] == 1:
                        message_list.append(f'<@{give_user_id}> now has 1 Fob!')
                    else:
                        message_list.append(f'<@{give_user_id}> now has {sublist[fob_count_index]} Fobs!')
                    added = True
                    break
        if not added:
            #sheet.append_row([give_user_name,give_user_id,1])  #COMMENT BACK IN LATER
            list_of_values.append([give_user_name,give_user_id,1])
            message_list.append(f'<@{give_user_id}> now has 1 Fob!')

#------------------------------------------------------------------------------
    
def add_fob(user_name,user_id):
    added = False
    
    for i in range(list_of_values.__len__()):
            if i != 0:
                sublist = list_of_values[i]
                if sublist[id_index] == user_id:
                    sublist[fob_count_index] = sublist[fob_count_index]+1
                    list_of_values[i] = sublist.copy()
                    #UPDATE SPREADSHEET ROW
                    if sublist[fob_count_index] == 1:
                        message_list.append(f'<@{user_id}> now has 1 Fob!')
                    else:
                        message_list.append(f'<@{user_id}> now has {sublist[fob_count_index]} Fobs!')
                    added = True
                    break
    if not added:
            #sheet.append_row([user_name,user_id,1])  #COMMENT BACK IN LATER
            list_of_values.append([user_name,user_id,1])
            message_list.append(f'<@{user_id}> now has 1 Fob!')

#------------------------------------------------------------------------------
    
def remove_fob(user_name,user_id):
    removed = False
    
    for i in range(list_of_values.__len__()):
        if i != 0:
            sublist = list_of_values[i]
            if sublist[id_index] == user_id and sublist[fob_count_index] > 0:
                sublist[fob_count_index] = sublist[fob_count_index]-1
                #UPDATE SPREADSHEET ROW
                if sublist[fob_count_index] == 0:
                    message_list.append(f'<@{user_id}> now has no Fobs!')
                elif sublist[fob_count_index] == 1:
                    message_list.append(f'<@{user_id}> now has 1 Fob!')
                else:
                    message_list.append(f'<@{user_id}> now has {sublist[fob_count_index]} Fobs!')
                removed = True
                break
                
            elif sublist[id_index] == user_id and sublist[fob_count_index] == 0:
                message_list.append(f'<@{user_id}> does not have a Fob, one cannot be Removed!')
                removed = False
                break
    if not removed:
        #sheet.append_row([user_name,user_id,0])  #COMMENT BACK IN LATER
        list_of_values.append([user_name,user_id,0])
        message_list.append(f'<@{user_id}> does not have a Fob, one cannot be Removed!')
        
#------------------------------------------------------------------------------
    
def check_fobs():
    for i in range(list_of_values.__len__()):
        if i != 0:
            sublist = list_of_values[i]
            if sublist[fob_count_index]  == 1:
                message_list.append(f'<@{sublist[id_index]}> has 1 Fob!\n')
            elif sublist[fob_count_index] > 1:
                message_list.append(f'<@{sublist[id_index]}> has {sublist[fob_count_index]} Fobs!\n')
    message_list.append("Nobody else currently has a Fob. :)")
                
                
###############################################################################
###########################__MAIN_PROGRAM_BLOCK__##############################
"""  
list_of_values = [['User Name','User ID','Fob Count'],['Adam',987654321,1],['Will',123456789,2]] #COMMENT OUT LATER

#take_and_give('Will',123456789,'Adam',987654321)
    
#add_fob('Jamis',123)

check_fobs()
    
for i in range(message_list.__len__()):
    message = message + " " + message_list[i]
    
print(message)
"""
#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
while True:

    list_of_values = []
    send_message = ""
    message_list = []
    
    check = 0
    transfer = 0
    add = 0
    remove = 0
    _help = 0
    
    sender_id = ""
    message_give_id = ""
    message_take_id = ""
    
    sender_name = ""
    message_give_name = ""
    message_take_name = ""
    #-----------------------------------------------------      
    #Get Message and Such    
    for message in slack_client.rtm_read():
        if 'text' in message:
            print(f'Message received:\n{json.dumps(message, indent=2)}')
            message_text = message['text']
            message_text = message_text.lower()
            sender_id = message['user']
            sender_id = sender_id.lower()
            message_text = message_text.replace('me',f'<@{sender_id}>')
            
            #-----------------------------------------------------    
            #Parse Message actions and ID's
            if re.match(r'.*(help).*', message_text, re.IGNORECASE):
                _help = 1
                
            if re.match(r'.*(check|who).*', message_text, re.IGNORECASE):
                check = 1
                
            if re.match(r'.*(add).*', message_text, re.IGNORECASE):
                add = 1
                message_give_id = message_text[message_text.find('@')+1:message_text.find('>')]
                
            if re.match(r'.*(remove).*', message_text, re.IGNORECASE):
                remove = 1
                message_remove_id = message_text[message_text.find('@')+1:message_text.find('>')]
                
            if re.match(r'.*(transfer|give|hand).*', message_text, re.IGNORECASE):
                transfer = 1
                from_text = message_text[message_text.find('from')+5:]
                message_take_id = from_text[from_text.find('@')+1:from_text.find('>')]
                to_text = message_text[message_text.find('to')+3:]
                message_give_id = to_text[to_text.find('@')+1:to_text.find('>')]
                  
            #-----------------------------------------------------   
            #Get Names for ID's
            
            user_list = slack_client.api_call("users.list")
            
            for user in user_list.get('members'):
                if user.get('id') == message_take_id:
                    message_take_name = user.get('name')
        
            for user in user_list.get('members'):
                if user.get('id') == message_give_id:
                    message_give_name = user.get('name')
                    
            for user in user_list.get('members'):
                if user.get('id') == sender_id:
                    sender_name = user.get('name')
            #-----------------------------------------------------    
            #Run Functions
            
            if check+transfer+add+remove > 1:
                message_list.append("Sorry, I can only understand 1 command at a time. \nPlease send each command individually and wait for my reply. \n Thanks!")
            elif check == 1:
                get_fob_database()
                check_fobs()
            elif transfer == 1:
                get_fob_database()
                take_and_give(message_take_name,message_take_id,message_give_name,message_give_id)
            elif add == 1:
                get_fob_database()
                add_fob(message_give_name,message_give_id)
            elif remove == 1:
                get_fob_database()
                remove_fob(message_take_name,message_take_id)
            elif _help == 1:
                message_list.append('Hello! Im Fob-Boy! \nIm here to help the WatLock team keep track of those pesky security fobs. \n\nIf you want to CHECK who has fobs, just DM me "Who has fobs" or "Check fobs" \n\n If you want to TRANSFER a fob between people, just DM me "Transfer fob from @Giver to @Receiver" or "Transfer from me to @Receiver". \n(Please @ people or use "Me" for proper user recognition). \n\nPlease do tell me when a Fob changes hands so I can inform everyone else on the team. \nThanks! :)')
                
            else:
                message_list.append('Sorry, I dont understand. Please use Keywords "Check", "Who", and "Transfer" in your command messages.\nOr, send "Help" for more information about me and my commands.')
            #-----------------------------------------------------
            #Compile and Send Message
            
            for i in range(message_list.__len__()):
                message = message + " " + message_list[i]
                
            
                
                
                
    time.sleep(1)