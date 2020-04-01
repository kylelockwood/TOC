#! python3
# Sends quarterly email to all volunteers

import toc
import kxl

def main():
    # Get spreadsheet data
    emails = toc.get_emails()
    emails += ['christine@theoregoncommunity.com']

    wb = toc.wb()
    volDates = format_data(kxl.data(wb, 'Email', 
                                    row_range=[1, 30], col_range=[1, 30]
                                    ).list_of('string', date_format='%m/%d', delimiter=', ')
                           )
    # Get the next season for email content
    first_date_on_sheet = toc.schedule().dates[0]
    season = toc.quarter(first_date_on_sheet).season
    # Send emails
    send_emails(emails, volDates, season)

def format_data(data):
    """ Removes unneeded commas and cleans up the strings """
    dataList = []
    for line in data:
        newline = ''
        lineList = line.split(',  ')
        for l in lineList:
            newline += l
        newline = newline.replace(',', ' : ', 1)
        newline = newline[:newline.rfind(', ')] + ''
        newline = newline.replace('-', ' -')
        dataList.append(newline)
    return dataList

def send_emails(emails, content, season):
    """ Sends data to list of emails """
    body = { 'subject': f'{season} season schedule has been posted',
             'body':    f'Here are the scheduled volunteer shifts for {season} quarter:\n\n'
            }
    for line in content:
        body['body'] += line + '\n'
    toc.mail(emails=emails, content=body)
    return

if __name__ =='__main__':
    main()
