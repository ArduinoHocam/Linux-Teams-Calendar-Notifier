# Name
#    lintecan.py
#
# Purpose
#    LINux TEams CAlendar Notifier
#
# Authors
#   Ali Sevindik - Arduino Hocam
# Version
#   1.0.2
#-- 

from os.path import isfile as os_path_isfile\
                ,join   as ospath_join
from os     import getcwd as os_getcwd
from O365 import Account, MSGraphProtocol
from datetime import timezone , timedelta
from time import sleep as time_sleep
from colorama import init, Fore, Style # I like visuality :P
from os.path import isfile
import configparser #ini parser
import notify2
import datetime as dt

# LINux TEams CAlendar Notifier
LINTECAN_HEADER = \
    """
            Welcome to LINux TEams CAlendar Notifier script!
                    ###########v1.0.2###############

    db      Sevindik d8b   db dUbuntub Arduino  .o88b.  .Ali.  Ali   db 
    88        `88'   Alio  88 `~~88~~' 88'     GNU  Y8 d8' `8b 888o  88 
    88         88    88V8o 88    88    8Hocam8 8P      8Alan_8 88V8o 88 
    88         88    88 V8o88    88    88~~~~~ 8b      8Turing 88 Linux 
    Kubuntu   ARM.   88  V888    88    88.     C++  d8 88   88 88  V888 
    Y88888P Assembly VP   Ali    YP    Arduino  `Y88P' YP   YP VP   V8P 

                    ###########v1.0.2###############
    """

#GLOBALS
#TODO: add to config file
ICON_PATH = "images/teams_icon.png"
CONFIG_PATH = "config/key.ini"
message =  "Meeting will start in: "
REMAINING_TIME_LIM = 15 #min
TIMEOUT_LIM = 30000 #millisec
VERSION = "1.0.2"
class Notify():
    _caption = None
    _msg = "None"
    _timeout = -1
    _urgency = None
    _obj = None

    def mess_callback():#TODO: calbacks
        print("closed the notif")
        pass

    def __init__(self, parent, caption, msg, timeout=None, urgency=None):
        self._caption = caption
        self._msg = msg
        self._timeout = timeout
        self._urgency = urgency

    def removeNotification(self):
        self._obj.close()

    def showNotification(self):
        caps = notify2.get_server_caps()    
        mess = notify2.Notification(self._caption, self._msg, ospath_join(os_getcwd(), ICON_PATH))
        mess.set_timeout(self._timeout) #milliseconds
        mess.set_urgency(self._urgency) 
        if self._timeout != 0 and 'actions' in caps:
            mess.add_action("close","Close",self.mess_callback,None) #TODO: callback
        mess.show()
        self._obj = mess
#end class Notify ##############################################################
    
class Event():
    _eventCaption = "None"
    _eventMsg = "None"
    _eventRemainingTimeMin = -1
    _eventIsNotified = False
    def __init__(self, eventCaption="None", eventMsg="None", remainingTimeMin =-1 ):
        self._eventCaption = eventCaption
        self._eventMsg = eventMsg
        self._eventRemainingTimeMin = remainingTimeMin
        self._eventIsNotified = False
#end class Event ###############################################################

def CommandlineArguments():
    """
    This method is used for the command line arguments.
    """
    from argparse import ArgumentParser \
                       , SUPPRESS       \
                       , RawDescriptionHelpFormatter
    parser = ArgumentParser \
        (prog="LINTECAN"
         , description=LINTECAN_HEADER
         , formatter_class=RawDescriptionHelpFormatter
        )
    parser.add_argument \
        ("-c", "--credential_file"
         , default="Dummy!!"
         , help="Specify the name of the file that contains credentials, if\
            no arguments are passed, GUI will open to enter credentials!!!"
        )
    parser.add_argument \
        ("-g", "--gui"
         , action="store_true"
         , default=None
         , help="Use GUI for credentials to type."
        )
    
    parser.add_argument \
        ("-v", "--version"
         , action="store_true"
         , default=None
         , help="Display the version of the program."
        )
    return parser.parse_known_args()
#CommandlineArguments ##########################################################

def getCredentialsFromGUI(client_id, secret_val):
    """
    Request login credentials using a GUI.
    """
    import tkinter
    root = tkinter.Tk()
    root.eval('tk::PlaceWindow . center')
    root.title('LINTECAN Login')
    uv = tkinter.StringVar(root, value=client_id)
    pv = tkinter.StringVar(root, value=secret_val)
    userEntry = tkinter.Entry(root, bd=3, width=35, textvariable=uv)
    passEntry = tkinter.Entry(root, bd=3, width=35, show="*", textvariable=pv)
    btnClose = tkinter.Button(root, text="OK", command=root.destroy)
    userEntry.pack(padx=10, pady=5)
    passEntry.pack(padx=10, pady=5)
    btnClose.pack(padx=10, pady=5, side=tkinter.TOP, anchor=tkinter.NE)
    root.mainloop()
    return [uv.get(), pv.get()]
#getCredentialsFromGUI #########################################################


def main():
    global message
    client_id = ""
    secret_value = ""
    init() # colorama init
    print(Fore.LIGHTCYAN_EX+ Style.BRIGHT+ LINTECAN_HEADER + Fore.RESET + "\n")
    arguments, unknown_args = CommandlineArguments()
    if arguments.version != None:
        print(Fore.YELLOW +  VERSION + Fore.RESET )
        exit(1)
    if arguments.credential_file != "Dummy!!":
        if os_path_isfile(arguments.credential_file) == False:
            print(Fore.RED + "Unable to find File : " + arguments.credential_file + Fore.RESET)
            exit(1)
        else:
            print(Fore.YELLOW + Style.BRIGHT + "Credential File is found!!: " + arguments.credential_file + Fore.RESET)
            config = configparser.ConfigParser()
            config.read(arguments.credential_file)
            client_id = config.get("CREDENTIALS", 'ClientID')
            secret_value = config.get("CREDENTIALS", 'SecretValue')
    elif arguments.gui != None :
        client_id, secret_value = getCredentialsFromGUI("<client id>", "<secret value>")
    else:
        print(Fore.RED +  "Invalid Arguments, Please check the Arguments!!!" + Fore.RESET)
        exit(1)

    credentials = (client_id, secret_value)
    protocol = MSGraphProtocol() 
    scopes = ['Calendars.Read.Shared', 'offline_access']
    account = Account(credentials, protocol=protocol)
    eventList = list()
    
    if account.authenticate(scopes=scopes):
        print(Fore.GREEN + '--------------'+ Fore.RESET)
        print(Fore.GREEN + 'Authenticated!')
        print(Fore.GREEN + '--------------\n'+ Fore.RESET)
    else:
        print(Fore.RED + 'UNABLE TO Authenticate!!!, Check your credentials!!!'+ Fore.RESET)

    while(True):#infinite loop
        schedule = account.schedule()
        calendar = schedule.get_default_calendar()
        now = dt.datetime.now()
        q = calendar.new_query('start').greater_equal(dt.datetime(now.year, now.month, now.day))
        #TODO : check if +1 works fine 
        q.chain('and').on_attribute('end').less_equal(dt.datetime(now.year, now.month, now.day) + dt.timedelta(days=1))
        events = calendar.get_events(query=q, include_recurring=True) 
        for event in events:
            if(not event.is_cancelled):
                start_date  = event.attachments._parent.start
                end_date  = event.attachments._parent.end
                now = dt.datetime.now(timezone.utc)
                duration = start_date - now
                duration_in_s = duration.total_seconds()
                minutes = int(divmod(duration_in_s, 60)[0])
                remaining_sec = int(divmod(duration_in_s, 60)[1])
                if minutes > 0: #if it is not overdued
                    eventList.append(Event(str(event.attachment_name), message, minutes))

        for validEvent in eventList:
            if(validEvent._eventRemainingTimeMin == REMAINING_TIME_LIM and validEvent._eventIsNotified == False):
                validEvent._eventIsNotified = True
                message = "Meeting will start in: " + str(int(validEvent._eventRemainingTimeMin)) + " minutes!!"
                notify2.init("fooApp")
                obj=Notify(None, str(validEvent._eventCaption), message, timeout=TIMEOUT_LIM, urgency=2)
                obj.showNotification()
                time_sleep(30)
                obj.removeNotification()
        eventList.clear()


if __name__ == "__main__":
    main()