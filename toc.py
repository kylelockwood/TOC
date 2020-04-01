#! python3
"""
Commonly used functions for The Oregon Community
"""

import datetime as dt
from datetime import timedelta
import random
import json
import csv
import os, sys
import kxl
import kmail

SCRIPTPATH = os.path.dirname(os.path.realpath(sys.argv[0])) + '\\'

def user_vars():
    """ Return dict of user variables from the JSON file 'uservars.json'
        JSON file must be located in script path"""
    with open('C:\\Users\\Kyle\\Dropbox\\PythonScripts\\TOC\\uservars.json') as f:
        return json.load(f)

def wb():
    user = user_vars()
    wb = user['path'] + user['book']
    return wb

def get_emails():
    """ Returns a list of email addresses from the Conflicts sheet """
    emails = kxl.data(wb(), 'Conflicts', row_range=[3, 36], col_range=[19]).list_of('string')
    return emails

def reformat_str_date(strDate):
    oldDate = strDate.split('-')
    newDate = oldDate[1] + '/' + oldDate[2]
    return newDate

class mail:
    """ userinfo is a dict of user information : 'name' and 'pass' req\n
        emails is a list of string email addresses\n
        content is a dict of 'subject' : string, and 'body' : string
    """
    def __init__(self, userinfo=user_vars(), emails=[], content={}):
        self.userinfo = userinfo
        if not emails:
            self.emails = get_emails()
        else:
            self.emails = emails
        self.content = {'subject': content['subject'], 'body':  self.__snips__(content['body'])}
        kmail.mail(self.userinfo, self.emails, self.content)

    def __snips__(self, content):
        """ Chooses random greeting and valediction, adds tag and signature """
        greetings = [   'Greetings Humans!',
                        'Good day all!',
                        'Hello everyone!',
                        'Top o\' the mornin\' to ya'
        ]
        greeting = random.choice(greetings) + '\n\n'
        tag = f'\nFor further schedule details, please visit www.theoregoncommunity.com/calendar\n\n'
        valedictions = ['Beep-boop-bop',
                        'Have a wonderful day!',
                        'See you Sunday!'
                        'Have a great week!'
        ]
        valediction = random.choice(valedictions) + '\n\n'
        signature = (   'Barry Pi\n'
                        'Digital Assistant to Mr. Lockwood\n'
                        'The Oregon Community'
                    )
        body = greeting + content + tag + valediction + signature
        return body

class quarter:
    """ Get information about quarter based on passed datetime.date
        Default datetime is today() """
    def __init__(self, searchdate=dt.date.today()):
        self.searchdate = searchdate
        self.seasons = ['Spring', 'Summer', 'Fall', 'Winter']
        self.season = self.get_season()
        self.months = self.get_months(self.season)
        self.next_season = self.get_next_season()
        self.next_months = self.get_months(self.next_season)
    
    def get_season(self):
        """ Returns the Season of the date sent """
        new_year_day = dt.date(year=self.searchdate.year, month=1, day=1)
        intday = (self.searchdate - new_year_day).days + 1
        # "day of year" ranges for the northern hemisphere
        spring = range(80, 172)
        summer = range(172, 264)
        fall = range(264, 355)
        # winter = everything else
        if intday in spring:
            return self.seasons[0]
        elif intday in summer:
            return self.seasons[1]
        elif intday in fall:
            return self.seasons[2]
        else:
            return self.seasons[3]

    def get_months(self, season):
        """ Returns the first and last month of the quarterly season """
        if season.lower() == 'spring':
            return ['April', 'May', 'June']
        elif season.lower() == 'summer':
            return ['July', 'August', 'September']
        elif season.lower() == 'fall':
            return ['October', 'November', 'December']
        else:
            return ['January', 'February', 'March']
    
    def get_next_season(self):
        """ Returns the season after the queried season """
        index_season = self.seasons.index(self.season) + 1
        if index_season > 3:
            index_season = 0
        return self.seasons[index_season]

class schedule:
    def __init__(self):
        self.wb = wb()
        self.location = user_vars()['location']
        self.dates = self.get_dates()
        self.services = self.get_services()
        self.descriptions = self.get_descriptions()
        self.schedule_data = self.get_schedule_data()

    def get_dates(self):
        """
        Returns list of datetime.date values from the sheet
        """
        dates = kxl.data(self.wb, 'Conflicts', col_range=[4,17], skip_none=False).list_of('list')
        for i in range(len(dates)):        
            if not dates[i] is None:
                dates[i] = dates[i].date()
        return dates
        
    def get_services(self):
        """ Returns a dict of key self.dates and cooresponding titles as items"""
        services = {}
        for i in range(len(self.dates)):
            name = 'Team Schedule'
            date = self.dates[i]
            if date is None:
                name = 'Empty Service'
                lastdate = self.dates[i-1]
                # If 2 or more None dates are returned, 
                # you are beyond the date scope, remove those items
                if lastdate is None:
                    self.dates.pop(i)
                    continue
                date = lastdate + timedelta(days=7)
            services[date] = name
        return services

    def get_descriptions(self):
        """ Returns a list of concatenated strings comprising the schedule descriptions """
        wbdata = kxl.data(self.wb, 'Web Cal', alerts=False, skip_none=False)
        descriptions = []
        for r in range(2, 41, 3):
            empty_check = wbdata.get_value(row=r, col=2)
            if empty_check is None:
                d = 'Empty Service'
            else:
                d = wbdata.list_of('string', row_range=[r], col_range=[1,40])
            descriptions.append(d)
        return descriptions

    def get_schedule_data(self):
        """ Combines service date, name, description, and location into a list of dicts """
        slist = []
        for i in range(len(self.services)):
            sdict = {}
            sdict['name'] = list(self.services.values())[i]
            sdict['date'] = list(self.services.keys())[i]
            sdict['description'] = self.descriptions[i]
            sdict['location'] = self.location
            slist.append(sdict)
        return slist

    def compare_content(self, fileName='schedule_log.csv', content=None):
        """ Reads the log file for comparing changes in the schedule
            Returns a list of changes
        """
        if content is None:
            content = self.schedule_data
        
        # If log file doesn't exist, create it
        if not os.path.exists(SCRIPTPATH + fileName):
            print('No log file present. ', end='')
            self.update_log(fileName, content)
            return []

        # Read the log file
        print(f'Comparing schedule changes... ', flush=True, end='')
        logList = []
        logDictFile = csv.DictReader(open(SCRIPTPATH + fileName))
        for lines in logDictFile:
            listDict = dict(lines)
            logList.append(listDict.copy())

        # Convert datetime.date to string date for comparison
        for lines in content:
            lines['date'] = lines['date'].strftime('%Y-%m-%d')

        # If dates don't match, i.e. new quarter, create a new log
        if logList[0]['date'] != content[0]['date']:
            print('New date data found.  ', end='')
            self.update_log(fileName, content)
            return []

        # Populate list allDiff with formatted differences
        allDiff = []
        chngCntr = 0
        checkRange = len(logList)
        if len(content) < len(logList): #Always iterate over the smallest of the two
            checkRange = len(content)
        for i in range(checkRange):
            diffList = []
            # If there has been a change in the content:
            if logList[i] != content[i]:
                date = reformat_str_date(logList[i]['date'])
                log_description = list(logList[i]['description'].split('|'))
                current_description = list(content[i]['description'].split('|'))
                for j in range(len(log_description)):
                    if log_description[j] != current_description[j]:
                        chngCntr += 1
                        current_role = current_description[j].split(' - ', 1)[0]
                        current_name = current_description[j].split(' - ', 1)[1]
                        log_name = log_description[j].split(' - ', 1)[1]
                        if current_name == 'None ':
                            current_name = 'No one '
                        elif log_name == 'None ':
                            log_name = 'no one '
                        diffList.append('- ' + current_name + 'is now on' + current_role + ' instead of ' + log_name)
                d = '\nTeam changes for ' + date + ':\n'
                for diff in diffList:
                    d += diff  +'\n'
                allDiff.append(d)
        if chngCntr is not 0:
            print(f'{chngCntr} change(s) found')
        else:
            print('No changes found')
        return allDiff

    def update_log(self, fileName, content):
        """ Writes list of dictionary content to csv file """
        print(f'Updating \'{fileName}\'... ', flush=True, end='')
        keys = content[0].keys()
        with open(SCRIPTPATH + fileName, 'w') as f:
            w = csv.DictWriter(f, keys)
            w.writeheader()
            w.writerows(content)
        print('Done')
        return

    def email_changes(self, content=None):
        """ Sends data to list of emails """
        if content is None:
            content = self.compare_content()
        if len(content) < 2:
            return
        else:
            body =  {   'subject':  'TOC volunteer schedule has changed',
                        'body':     'The Oregon Community volunteer schedule has been updated, here are the changes:\n'
                    }
            for line in content:
                body['body'] += line
            mail(content=body)
        return
