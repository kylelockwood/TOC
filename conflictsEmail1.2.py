#! python3
# Sends quarterly email to all volunteers

import toc

def main():
    # Whom to email
    emails = toc.get_emails()
    emails += ['christine@theoregoncommunity.com']

    # Get email content
    season = toc.quarter().next_season
    months = toc.quarter().next_months

    # Send reminder emails
    send_emails(emails, season, months) # This is turned off in the function while testing

    print('Done')

def send_emails(emails, season, months):
    """ Sends data to list of emails """ 
    body = {'subject':  f'Oregon Community {season} Schedule Time!',
            'body':     f'Time for our {season} season schedule.  '
            f'Please send me your Sunday conflicts for the months of {months[0]} through {months[2]}.  '
            'For our new members, those are the Sundays you are NOT available to volunteer.  '
            'Remember, the last Sunday of the month is Empty service so you don\'t need to send me your conflicts on those dates.\n'
            }
    toc.mail(emails=emails, content=body)
    return

if __name__ =='__main__':
    main()
