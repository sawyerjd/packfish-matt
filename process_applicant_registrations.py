from imap_tools import MailBox, AND
from bs4 import BeautifulSoup
import config as cfg

# define results file
results = {}


# Functions
def set_exams(exams):
    # set_exams receives a list of exams the applicant is interested in taking.  The exams list will be parsed and
    # the correct columns for the elements will be set.
    # If element 3 and/or element 4 are selected, these selections should show up on the Applicant's
    # registration form in Session Manager.

    # define form field names
    element2 = 'Element 2 (Technician)'
    element3 = 'Element 3 (General)'
    element4 = 'Element 4 (Amateur Extra)'

    # Define column heading names to be used with the csv file.
    req_element_3 = 'REQUESTED_ELEMENT_3'
    req_element_4 = 'REQUESTED_ELEMENT_4'

    # loop through each exam and set corresponding element name and value.
    for exam in exams:
        if exam == element2:
            # if element 2, set element 3 and 4 to false
            results[cfg.Header.fields[req_element_3]] = False
            results[cfg.Header.fields[req_element_4]] = False
        # if element 3, set element 3 to true
        if exam == element3:
            results[cfg.Header.fields[req_element_3]] = True
        # if element 4, set element 4 to true
        if exam == element4:
            results[cfg.Header.fields[req_element_4]] = True
    return None

# TODO: add logging to script.
# TODO: add logic to output dictionary items to a csv.  Reference: https://pythonguides.com/python-dictionary-to-csv/


# main code
def main():
    print("Starting Script")
    mb = MailBox(cfg.Mail.server).login(cfg.Mail.user, cfg.Mail.password)
    print("Fetching application forms")
    messages = mb.fetch(criteria=AND(from_="burst@emailmeform.com", seen=False), mark_seen=False, bulk=True)

    # Start processing retrieved applications
    for msg in messages:
        print("Processing Messages")
        # parse html email message
        bs = BeautifulSoup(msg.html, 'html.parser')
        # emailme form is output to a table.
        # get all table rows
        table_rows = bs.find('table').findAll('tr')
        # process each row to get the table data.
        for row in table_rows:
            name = ""
            value = ""
            td = row.findAll('td')

            # parse each row to get the field name and field value.
            # field name will contain '*:' and the will need to be stripped out.
            # the first td item should be the field name with *:, ex first name*:
            # the second td item should be the value

            for item in td:
                if item.text.find("*:") != -1:
                    name = item.text.strip().replace("*:", "")
                else:
                    value = item.text.strip()
            # before adding the name/value pair to the results dictionary, check for required modifications.
            # Check for items that need updated before adding to csv file
            if name == 'Middle Initial' and value.upper() == 'NONE':
                # If Middle Initial is NONE, set value to empty string
                value = ''
            elif name == 'Suffix' and value.upper() == 'NONE':
                # If Suffix is NONE, set value to empty string
                value = ''
            elif name == 'Street Address' and value.find('PO') == -1:
                # If not a PO Box, add empty PO Box entry
                results[cfg.Header.fields['PO Box']] = ''
            elif name == 'Street Address' and value.find('PO') != -1:
                # If PO Box, set Street Address(value) to empty string and add PO_BOX
                results[cfg.Header.fields['PO Box']] = value
                value = ''
            elif name == 'Callsign' and value.upper() == 'NOCALL':
                # If callsign is NOCALL, set Callsign value to empty string and set UPGRADE_LICENSE to False
                value = ''
                results[cfg.Header.fields['UPGRADE_LICENSE']] = False
            elif name == 'Callsign' and value.isalnum():
                # If a callsign was entered, set UPGRADE_LICENSE to True
                results[cfg.Header.fields['UPGRADE_LICENSE']] = True
            elif name == 'Exams':
                set_exams(value.split(', '))

            print('=========================')
            print(f'Name: {name}, Value: {value}')

            results[cfg.Header.fields[name]] = value
    print('Print results')
    print(results)

    # Logout of mailbox
    mb.logout()


if __name__ == '__main__':
    main()
