# Useful Links
# https://www.crummy.com/software/BeautifulSoup/bs4/doc/#navigating-the-tree
# https://www.twilio.com/blog/web-scraping-and-parsing-html-in-python-with-beautiful-soup
# https://realpython.com/beautiful-soup-web-scraper-python/
# Get test session email registrations and generate a csv file for import
import yaml
from imap_tools import MailBox, AND
import csv
from os.path import exists
import logging
import datetime

log_filename = datetime.datetime.now().strftime("%m%d%Y_%H%M%S") + "_sessionmanager.log"
logging.basicConfig(filename=log_filename, level=logging.DEBUG)

csv_filename = datetime.datetime.now().strftime("%m%d%Y_%H%M%S") + "_session_import.csv"
# csv_filename = "session_import.csv"
# 04222022 MAB - Added UPGRADE_LICENSE column to the csv to support upgrade pre-registered applicants
fields = ['FIRST_NAME', 'MIDDLE_INITIAL', 'LAST_NAME', 'NAME_SUFFIX', 'STREET_ADDRESS', 'PO_BOX', 'CITY',
          'STATE', 'ZIP_CODE', 'TELEPHONE', 'E_MAIL', 'FRN', 'CALL_SIGN', 'FELONY_CONVICTION', 'NOTES', 'CERTIFYING_VES', 'UPGRADE_LICENSE']


def find_starting_index(email_text, search_string, start_index):
    # Gets the starting index of the search string
    return email_text.find(search_string, start_index)


def find_value(email_text, starting_index, ending_index):
    # finds field value between indexes
    logging.info(f'Found: {email_text[starting_index: ending_index].strip()}')
    # returns a substring within the starting and ending index
    return email_text[starting_index: ending_index]


def set_name(msg_text, field, start_index, end_index):
    if field == 'first':
        logging.info('Setting first name')
        field_start = 14
        first_name = find_value(msg_text, start_index + 2, end_index - field_start)
        logging.info(f'first_name: {first_name.strip()}')
        return first_name.strip()
    elif field == 'middle':
        logging.info('Setting middle initial')
        field_start = 9
        middle_initial = find_value(msg_text, start_index + 2, end_index - field_start)
        logging.info(f'middle_initial: {middle_initial.strip()}')
        return middle_initial.strip()
    elif field == 'last':
        logging.info('Setting last name')
        field_start = 6
        last_name = find_value(msg_text, start_index + 2, end_index - field_start)
        logging.info(f'last_name: {last_name.strip()}')
        return last_name.strip()
    elif field == 'suffix':
        logging.info('Setting suffix')
        field_start = 14
        suffix = find_value(msg_text, start_index + 2, end_index - field_start)
        if suffix.strip() == 'NONE':
            suffix = ''
        else:
            logging.info(f'suffix: {suffix}')
        return suffix.strip()


def set_address(msg_text, field, start_index, end_index):
    if field == 'street':
        field_start = 4
        logging.info('Setting street_address')
        street_address = find_value(msg_text, start_index + 2, end_index - field_start)
        logging.info(f'street_address: {street_address.strip()}')
        return street_address.strip()
    elif field == 'city':
        field_start = 5
        logging.info('Setting city')
        city = find_value(msg_text, start_index + 2, end_index - field_start)
        logging.info(f'city: {city.strip()}')
        return city.strip()
    elif field == 'state':
        logging.info('Setting state')
        field_start = 8
        state = find_value(msg_text, start_index + 2, end_index - field_start)
        logging.info(f'state: {state.strip()}')
        return state.strip()
    elif field == 'zip':
        logging.info('Setting zip')
        field_start = 5
        zip_code = find_value(msg_text, start_index + 2, end_index - field_start)
        logging.info(f'zip code: {zip_code.strip()}')
        return zip_code.strip()


def set_phonenumber(msg_text, phone_index, email_index):
    phone_number = find_value(msg_text, phone_index + 2, email_index - 6)
    logging.info(f'phone number: {phone_number.strip()}')
    return phone_number.strip()


def set_email(msg_text, email_index, callsign_index):
    email_address = find_value(msg_text, email_index + 3, callsign_index - 9)
    logging.info(f'email address: {email_address.strip()}')
    return email_address.strip()


def set_callsign(msg_text, callsign_index, frn_index):
    # 04222022 mab - added global variable upgrade_license to support applicants who are upgrading
    global upgrade_license
    upgrade_license = False
    callsign = find_value(msg_text, callsign_index + 2, frn_index - 15)
    # 04212022 - Added .upper() to string for comparison.  Someone actually typed "Nocall" and we set the call to "Nocall"
    callsign = callsign.strip().upper()
	
    # 04212022 - Added a check for a space in the middle of the call sign.  Someone actually put a space in their call sign
    # if a space isn't found, position will be -1
    position = callsign.find(" ")
    while position != -1:
        callsign = callsign.replace(" ", "")
        position = callsign.find(" ")
	
    if callsign == 'NOCALL':
        callsign = ''
    else: 
        upgrade_license = True
    logging.info(f'callsign: {callsign}')
    return callsign


def set_frn(msg_text, frn_index, exams_index):
    global frn
    frn = find_value(msg_text, frn_index + 2, exams_index - 33)
    logging.info(f'frn: {frn.strip()}')
    return frn.strip()


def set_exams(msg_text, exams_index, felony_index):
    global exams
    exams = find_value(msg_text, exams_index + 2, felony_index - 18)
    logging.info(f'exams: {exams.strip()}')
    return exams.strip()


def set_felony(msg_text, felony_index, end_index):
    global felony
    felony = find_value(msg_text, felony_index + 3, end_index + 6)
    logging.info(f'felony: {felony.strip()}')
    return felony.strip()


def create_import_file(filename, fields):
    if exists(filename):
        logging.info(f'{filename} already exists. ')
    else:
        with open(filename, 'w', newline='') as csv_out:
            # creating a csv writer object
            logging.info(f'Creating {filename}')
            csvwriter = csv.writer(csv_out)
            csvwriter.writerow(fields)


def get_mail_creds(mail_config):
    """Get mail server and credentials from YAML config file."""
    return mail_config['url'], mail_config['uid'], mail_config['pwd']


def yaml_loader(config_file):
    """Loads YAML config file"""
    with open(config_file, "r") as file:
        data = yaml.safe_load(file)
        return data


def get_certifying_ves(ve_data):
    """Get certifying VEs from the config file."""
    ve1 = ve_data['ve1'].upper()
    ve2 = ve_data['ve2'].upper()
    ve3 = ve_data['ve3'].upper()
    return ve1, ve2, ve3


if __name__ == '__main__':

    logging.info(f'{datetime.datetime.now()} ===================================')
    logging.info('Application export starting')

    # load config from yaml file
    config_file_path = "config.yaml"
    logging.info(f'Loading YAML config file, {config_file_path}.')
    config = yaml_loader(config_file_path)

    # Get imap config from config file
    logging.info('Getting Mail server config.')
    mail_config = config.get('server')
    server, user, password = get_mail_creds(mail_config)

    # get Certifying VEs from config file
    logging.info('Getting Certifying VEs from config file.')
    cv_data = config.get('certifying_ves')
    VE1, VE2, VE3 = get_certifying_ves(cv_data)
    logging.info(f'The following VEs were retrieved from the config file: {VE1}, {VE2}, {VE3}.')
    delim = '~'
    # Certifying VEs must be added to the test session before running the import.  Otherwise,
    # the certifying VEs will not be imported into the applicant record.
    # format 'cs1~cs2~cs3'
    certifying_ves = VE1 + delim + VE2 + delim + VE3

    create_import_file(csv_filename, fields)

    mb = MailBox(server).login(user, password)
    messages = mb.fetch(criteria=AND(from_="burst@emailmeform.com", seen=False), mark_seen=True, bulk=True)
    message_count = 0
    for msg in messages:
        logging.info(msg)
        msg_text = msg.text
        logging.info('Printing message\n')
        logging.info(msg_text)

        if msg_text[:15] == 'Full Legal Name':
            logging.info('Skipping message, Full Legal Name invalid, old format\n')
            continue

        # these indexes will be used to find the values
        logging.info('\n\n')
        logging.info(f'Setting indexes')
        first_name_index = find_starting_index(msg_text, '*:', 0)
        middle_initial_index = find_starting_index(msg_text, '*:', first_name_index + 2)

        if msg_text[middle_initial_index - 14: middle_initial_index] != 'Middle Initial':
            logging.info('Skipping message, Middle Initial not valid, old format\n')
            continue
        last_name_index = find_starting_index(msg_text, '*:', middle_initial_index + 2)
        suffix_index = find_starting_index(msg_text, '*:', last_name_index + 2)
        street_address_index = find_starting_index(msg_text, '*:', suffix_index + 2)
        city_index = find_starting_index(msg_text, '*:', street_address_index + 2)
        state_index = find_starting_index(msg_text, '*:', city_index + 2)
        zip_code_index = find_starting_index(msg_text, '*:', state_index + 2)
        phone_index = find_starting_index(msg_text, '*:', zip_code_index + 2)
        email_index = find_starting_index(msg_text, '*:', phone_index + 2)
        callsign_index = find_starting_index(msg_text, '*:', email_index + 2)
        frn_index = find_starting_index(msg_text, '*:', callsign_index + 2)
        exams_index = find_starting_index(msg_text, '*:', frn_index + 2)
        felony_index = find_starting_index(msg_text, '*:', exams_index + 2)

        # set each field using the indexes above
        logging.info(f'Setting fields')
        first_name = set_name(msg_text, 'first', first_name_index, middle_initial_index)
        middle_initial = set_name(msg_text, 'middle', middle_initial_index, last_name_index)
        last_name = set_name(msg_text, 'last', last_name_index, suffix_index)
        suffix = set_name(msg_text, 'suffix', suffix_index, street_address_index)
        address = set_address(msg_text, 'street', street_address_index, city_index)

        if address[:2].lower() != 'po':
            street_address = address
            po_box = ''
        else:
            po_box = address
            street_address = ''

        city = set_address(msg_text, 'city', city_index, state_index)
        state = set_address(msg_text, 'state', state_index, zip_code_index)
        zip_code = set_address(msg_text, 'zip', zip_code_index, phone_index)

        phone_number = set_phonenumber(msg_text, phone_index, email_index)

        email_address = set_email(msg_text, email_index, callsign_index)

        callsign = set_callsign(msg_text, callsign_index, frn_index)

        frn = set_frn(msg_text, frn_index, exams_index)

        exams = set_exams(msg_text, exams_index, felony_index)

        felony = set_felony(msg_text, felony_index, felony_index)

        examinee_info = [
            first_name,
            middle_initial,
            last_name,
            suffix,
            street_address,
            po_box,
            city,
            state,
            zip_code,
            phone_number,
            email_address,
            frn,
            callsign,
            felony,
            exams,
            certifying_ves,
            upgrade_license
        ]

        logging.info(f'Examinee info: {examinee_info}')

        with open(csv_filename, 'a', newline='') as csv_out:
            # creating a csv writer object
            logging.info(f'Appending to {csv_filename}')
            csvwriter = csv.writer(csv_out)
            # write field names
            csvwriter.writerow(examinee_info)
        message_count += 1

        logging.info(f'{first_name} {last_name} added to {csv_filename}')

    if message_count == 0:
        logging.info(f'{message_count} applications processed')
        print(f'{message_count} applications processed')
    else:
        logging.info(f'{message_count} new applications processed.')
        print(f'{message_count} new applications processed.')

# Logout of mailbox
mb.logout()
logging.info('Application export complete! =============================\n\n')
