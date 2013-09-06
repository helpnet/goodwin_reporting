# -*- coding: utf-8 -*-
import os
import sys
import csv
import urllib2
import cookielib
import urllib
import webbrowser

from HTMLParser import HTMLParser
#from Reports import *
from config import *
from pprint import pprint

list_stu = {}
list_problem = {}

class MLStripper(HTMLParser):
    def __init__(self):
        self.reset()
        self.fed = []
    def handle_data(self, d):
        self.fed.append(d)
    def get_data(self):
        return ''.join(self.fed)


def logMsg(msg):
    print (msg + "\n")
    return


def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()


def strCleanUp(string):
    temp = strip_tags(string)
    temp.replace("&#039;", "'")
    temp.replace("<p>", "")
    temp.replace("</p>", "")
    return temp


def getProgressBar(total):
    x = total / 10
    return x


def write_data(data, path):
    writer = csv.writer(open(path, 'wb'), dialect='excel')
    writer.writerow(['URL', 'DATE_OPEN', 'NAME', 'EMAIL', 'TOPIC', 'CONTENT', 'OS', 'BROWSER', 'COURSE_ID', 'JAVA'])

    for row in data:
        writer.writerow(row)

    return


def buildRecordList(num_record, interval, opener):
    global ticket_hold

    record = []

    for i in range(0, num_record):
        record.append(pullData(ticket_hold[i], opener))
        if not (i % interval):
            print ("[]"),
    return record
    

def pullData(url, opener):
    global list_stu, list_problem
    
    # opens the website
    resp = opener.open(url)

    #correct page is now open
    content = resp.read()

    garbage = content.split('''</a> (''',1)[1]
    name = garbage.split(''') <span class="ticket_count">''',1)[0]

    garbage = content.split('''<strong>E-Mail:</strong> ''',1)[1]
    email = garbage.split('''\r\n\t\t\t\t\t\t\t<br/>''',1)[0]

    garbage = content.split('''<strong>Publish Date:</strong> ''',1)[1]
    date_open = garbage.split('''\r\n\t\t\t\t\t\t\t<br/>''',1)[0]

    garbage = content.split('''<strong>Topic:</strong> ''',1)[1]
    topic = garbage.split('''\r\n\t\t\t\t\t\t\t<a class="view-ticket-notes"''',1)[0]

    garbage = content.split('''View/Add Notes</a></div> ''',1)[1]
    ticket_content = garbage.split('''<p>&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;''',1)[0]

    garbage = content.split('''<strong>Course ID (ex. EDUC-500-150):</strong> ''',1)[1]
    course_id = garbage.split('''<br />''',1)[0]

    garbage = content.split('''<strong>Java Version (instructions are above):</strong> ''',1)[1]
    java_version = garbage.split('''<br />''',1)[0]

    garbage = content.split('''<strong>Visitor OS:</strong> ''',1)[1]
    os = garbage.split('''<br />''',1)[0]

    garbage = content.split('''<strong>Visitor Browser:</strong> ''',1)[1]
    browser = garbage.split('''<br />''',1)[0]

    if browser == "Chrome" and os == "Safari":
        os = "Macintosh"

    if len(course_id.split("-")) != 3:
        temp = url.split("/")
        edit_url = "http://stream.goodwin.drexel.edu/learningtech/ticketsystem/wp-admin/post.php?post=" + temp[6].split("-")[0] + "&action=edit"
        list_problem[name + "ID"] = ["ERROR: Bad course id: " + course_id + ". Opening page.", edit_url]

    if (len(course_id) != 11) and (len(course_id) != 12) and (len(course_id) != 10) and course_id != "SEE-TICKET-INFO":
        temp = url.split("/")
        edit_url = "http://stream.goodwin.drexel.edu/learningtech/ticketsystem/wp-admin/post.php?post=" + temp[6].split("-")[0] + "&action=edit"
        list_problem[name + "ID"] = ["ERROR: Bad course id: " + course_id + ". Opening page.", edit_url]

    if len(java_version.split(".")) != 3:
        if java_version != "N/P":
            temp = url.split("/")
            edit_url = "http://stream.goodwin.drexel.edu/learningtech/ticketsystem/wp-admin/post.php?post=" + temp[6].split("-")[0] + "&action=edit"
            list_problem[name + "ID"] = ["ERROR: Bad Java version: " + java_version + ". Opening page.", edit_url]

    if not name in list_stu:
        list_stu[name] = [os, browser, [course_id], java_version]

    else:
        if not course_id in list_stu[name][2]:
            list_stu[name][2].append(course_id)

    myList = [url, date_open, name, email, topic, ticket_content, os, browser, course_id, java_version]
    
    for item in myList:
        item = strCleanUp(item)
        
    return myList


def buildTicketURLList(content):
    global ticket_hold
    
    while len(content) > 1:
        # content[0] = garbage content[1] = url + garbage
        unparsed_url = content[1]

        # parsed_url[0] = URL [1] = URL + garbage
        parsed_url = unparsed_url.split('" target="_blank"',1)

        # moving url into holder
        ticket_hold.append(parsed_url[0])

        # prepping for next instance of the loop
        content = parsed_url[1].split('''<span class='view'><a href="''',1)

    return


def pullURLGarbage():
    global uname, pwd
    x = 2
    clean_resp = ""

    cj = cookielib.CookieJar()
    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))
    login_data = urllib.urlencode({'log' : uname, 'pwd' : pwd})

    # and we're in the system! NOTE: this will not throw an exception if the login information is incorrect!
    opener.open(url_login, login_data)
    logMsg("\nWe're in the system now. Please wait while \nthe system processes all of the pages.")
    
    # gets the first page of tickets
    dirty_opener = opener.open(url_ticket_home)
    dirty_resp = dirty_opener.read()
    
    # gets rid of the crap at the end
    dirty_resp = dirty_resp.split('</tbody></table><div class="tablenav bottom">')

    # moves it to clean_resp for safe keeping
    clean_resp = dirty_resp[0]
       
    while x > 0:
        print ("[]"),
        
        # opens the next page
        dirty_opener = opener.open(url_tickets + str(x))

        # IMPORTANT: WordPress likes to be helpful and correct incorrect page calls.
        #            This code makes sure that the URL we sent is the URL we're looking
        #            at. If it's not the same, it assumes we've hit the end and need to
        #            move on.
        #
        if (dirty_opener.geturl() != (url_tickets + str(x))):
            break
        
        dirty_resp = dirty_opener.read()

        # gets rid of the crap at the beginning (for a seamless transition)
        dirty_resp = dirty_resp.split('<tbody id="the-list">')[1]
    
        # gets rid of the crap at the end
        dirty_resp = dirty_resp.split('</tbody></table><div class="tablenav bottom">')
    
        # moves it into the next clean bin
        clean_resp = clean_resp + dirty_resp[0]
    
        #onwards!
        x+=1

    logMsg("\n\nAll done! Processed: " + str(x - 1) + " page(s) from the ticketing system.")
    
    content = clean_resp.split('''Trash</a> | </span><span class='view'><a href="''',1)

    return content, opener


def main():
    global list_stu, list_problem
    
    path = "reporting" + str(now)[:-10] + ".csv"

    logMsg("mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm\n" + 
           "mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm\n" + 
           "mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmdhdmmmmmmmmmmmmm\n" + 
           "mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmds:+ddmmmmmmmmmmmmm\n" + 
           "mmmmmmmmmmmmmmmmmmmddyshhhhmmmss..odhs/sohdmmmmmmm\n" + 
           "mmmmmmmmmmmmmmmmmd+/-.:-.-/oyh.+..+:://-+dmmmmmmmm\n" + 
           "mmmmmmmmmmmmmmmmy+.........-/s:....:oosydmmmmmmmmm\n" + 
           "mmmmmmdyssydmmmd+....-+o+:.../o-...+s+/dmmmmmmmmmm\n" + 
           "mmmmmy:./shydmmh:....hmmmd:...+:...sdddmmmmmmmmmmm\n" + 
           "mmmmm:..ommmdmmh/....smmmmo....-..../oydmmmmmmmmmm\n" + 
           "mmmmm:../dmmmmmds-...-ymmm+....o/+.....:ydo+smmmmm\n" + 
           "mmmmmy-.:sdmmmmmd/....-dmmo....+hh-.....::-oodmmmm\n" + 
           "mmmmmmy-.-+syyyho:-...-dmmd/....--......++s+ymmmmm\n" + 
           "mmmmmmmh+-.-----...../hmmmmds:......-::ohdhydmmmmm\n" + 
           "mmmmmmmmmhs/-.....:oyhhsydmmmdy+:-.-/o:/+ymmmmmmmm\n" + 
           "mmmmmmmmmmmdy.....+o+oo-osmmmmmmdhs--./ohhmdhhhmmm\n" + 
           "mds//+ssoooyhyyyyyyyydyyddmmmmmmmmmyooo+smmh+.ommm\n" + 
           "mmm:.smmdhs-:ymmmdmddmmmddmmmmmmmmmmmmddddmmo.ommm\n" + 
           "mmm/.smmmmmo.-hs+/y//dyos++ys+/odo+shhoso/sdo.ommm\n" + 
           "mmm/.smmmmms..dh.-yyso.shy-:mh:-o+hmy-+hh/-ho.ommm\n" + 
           "mmm/.smmmmd/.+mh.-mmm..yhhhhmmh-.odm/.ohhhhdo.ommm\n" + 
           "mdd:.odddy+:sddh.-ddm+./yhyydy/yo./hs--shhsdo.odmm\n" + 
           "mdyooosysyhdmmdsooydmdyo++shyosdhooshho++shhsoshmm\n" + 
           "dddddddddddddddddddddddddddddddddddddddddddddddddd\n")

    logMsg("Thank you for using the Goodwin College Ticketing System Buddy.\n" +
           "This program is intended to simplify the reporting process for \n" +
           "the college, and to allow for seamless reporting improvements.")

    raw_input("Please press return to continue: ")

    content, opener = pullURLGarbage()

    logMsg("Building initial list of tickets, please wait.")

    buildTicketURLList(content)

    logMsg("List complete. Setting up progress bar.")

    interval = getProgressBar(len(ticket_hold))

    logMsg("Progress bar complete. Pulling student data from system now.")

    record = buildRecordList(len(ticket_hold), interval, opener)

    logMsg("\n\nConsolidation complete. Writing " + str(len(ticket_hold)) + " records \nto file '"
           + path + "'.")

    write_data(record, path)

    ans = raw_input("Success! The file '" + path + "' is ready for review. Would you\n" +
              "like the program to run an error report on the records?\n\n(Y/N): ")

    if ans == 'Y' or 'y':
        for key in list_problem:
            pprint(list_problem[key][0])
            firefox.open(list_problem[key][1])
            raw_input()

##    if ans == 'Y':
##        writer = csv.writer(open(path, 'a'), dialect='excel')
##        writer.writerow(["this is a test"])
    

main()
