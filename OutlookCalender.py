import win32com.client, datetime
import re
import webbrowser
import ctypes

##############################################################
#regular expression to parse URL of zoom meeting host
zmurlre = re.compile(r'(?<=https\:\/\/)\S+(?=\/j\/)')

#regular expression to parse meeting ID
zmidre = re.compile(r'(?P<zoommtgID>(?<=\/j\/)([^?|\s]*))')

#regular express to parse meeting pw
zmpw = re.compile(r'(?P<zoommtgPW>(?<=pwd=)\S+)')

##############################################################


def zoomurlparse(oldzoomurl):
    lzid= str()
    lzurl = str()
    lzpw = str()
    IDmatch = re.finditer(zmurlre, oldzoomurl)
    URLmatch = re.finditer(zmidre, oldzoomurl)
    PWmatch = re.finditer(zmpw, oldzoomurl)
    for x in IDmatch:
        lzid = x.group()  # Local Zoom ID (LZID)
        if lzid.__contains__('>'):
            lzid = str(lzid[:-1])
        else:
            pass
    for x in URLmatch:
        lzurl = x.group()  # Local Zoom URL (lzurl)
        if lzurl.__contains__('>'):
            lzurl = str(lzurl[:-1])
        else:
            pass
    for x in PWmatch:
        lzpw = x.group()  # local zoom PW
        if lzpw.__contains__('>'):
            lzpw = str(lzpw[:-1])
        else:
            pass
    return lzurl, lzid, lzpw

def zoomlbuild(oldzoomurl):
    IDmatch, URLmatch, PWmatch = zoomurlparse(oldzoomurl)

#    print('IDmatch is ' + IDmatch)
#    print('URLMatch is ' + URLmatch)
#    print('PWMatch is ' + PWmatch)
    if PWmatch:
        rawzoomurl = str('zoommtg://' + URLmatch + '/join?action=join&confno=' + IDmatch + '&pwd=' + PWmatch)
    else:
        rawzoomurl = str('zoommtg://' + URLmatch + '/join?action=join&confno=' + IDmatch)
    return rawzoomurl

def zoomparse(msgbody):
    zoomurl = zoomlbuild(msgbody)
    return zoomurl

def getTodayCal():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.Session
    myCalendar = ns.GetDefaultFolder(9)
    items = myCalendar.Items
    begin = datetime.date.today()
    end = begin + datetime.timedelta(days=1)
    restriction= "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
    restrictedItems = items.Restrict(restriction)
    return restrictedItems

todaysCalEvents = getTodayCal()
Meetings = []

for appointmentItem in todaysCalEvents:
    meeting_dict = {}
    if appointmentItem.Subject.startswith("Canceled:"):
        pass
    else: # 'zoom' in appointmentItem.Body:
        subject = appointmentItem.Subject
#        print(subject)
#        print(appointmentItem.Body)
        url = zoomparse(appointmentItem.Body)
#       print(url)
        meeting_dict = {'subject': subject, 'meetingurl': url}
        Meetings.append(meeting_dict)

index = 0
for meeting in Meetings:
    if meeting['meetingurl']:
        print('Index No. ' + str(index) + ' ' + meeting['subject'])
    else:
        pass
    index = index + 1
meetingindextojoin = int(input('Enter meeting index you want to join\n'))

print('Would you like to join the following meeting:\n' + Meetings[meetingindextojoin]['meetingurl'])
response = input('Yes or No\n')
if response == 'Yes':
    webbrowser.open(Meetings[meetingindextojoin]['meetingurl'])
    print(Meetings[meetingindextojoin]['meetingurl'])
else:
    pass