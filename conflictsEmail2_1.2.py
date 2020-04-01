#! python3
# Sends second round of conflict emails

import toc

def main():
    # Whom to email
    emails = toc.get_emails()
    emails += ['christine@theoregoncommunity.com']

    # Get email content
    season = toc.quarter().next_season
    months = toc.quarter().next_months

    # Send reminder emails
    send_emails(emails, season, months)

    print('Done')

def send_emails(emails, season, months):
    """ Sends data to list of emails """ 
    body = {'subject':  f'Oregon Community {season} Schedule, Round Two!',
            'body':     f'I\'ve gotten {season} season Sunday conflicts from about half of you so far.  '
                        f'If you haven\'t yet, please send me your Sunday conflicts for the months of {months[0]} through {months[2]}.\n'
            }
    toc.mail(emails=emails, content=body)
    return

if __name__ =='__main__':
    main()
