
# $env:PATH += ';C:\Program Files\wkhtmltopdf\bin'
from datetime import datetime
import logging
import os
import pdb
from pathlib import Path, PurePath
from pprint import pprint
import shutil

import win32com
from msgtopdf import Msgtopdf

DATA_DIR = Path('data').resolve()
INPUT_DIR = DATA_DIR.joinpath('input')
OUTPUT_DIR = DATA_DIR.joinpath('output')

outlook = None


logging.getLogger().setLevel(logging.DEBUG)
OL_DISCARD = 1

CLIENT_EMAIL_TO_NAME: dict[str, str] = {}


DATE_FORMAT = '%d.%m.%y'


def init():
    global outlook

    if outlook is None:
        outlook = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")


def read_clients_file():
    for line in DATA_DIR.joinpath('Clients.txt').read_text('utf8').splitlines():
        line = line.strip()
        pos = line.find(' ')
        if pos == -1:
            continue
        email_suffix = line[0:pos]
        name = line[pos+1:].strip()
        if not (name and email_suffix):
            continue

        CLIENT_EMAIL_TO_NAME[email_suffix] = name


def get_msg_output_name(msg):
    now = datetime.now().strftime(DATE_FORMAT)
    email_address = msg.SenderEmailAddress
    client_name = resolve_client_name(email_address)
    if client_name is None:
        client_name = email_address

    received_date = msg.ReceivedTime.strftime(DATE_FORMAT)
    return f"{now} מאת- {client_name} התקבל ב- {received_date}"


def resolve_client_name(email_address):

    for (email_suffix, name) in CLIENT_EMAIL_TO_NAME.items():
        if email_address.endswith(email_suffix):
            return name


class _Msgtopdf(Msgtopdf):
    def __init__(self, msg):
        self.msg = msg


def handle_msg_file(msg_path: Path):
    global outlook
    msg = outlook.OpenSharedItem(msg_path)
    output_path = None
    try:
        output_name = get_msg_output_name(msg)
        print(output_name)
        output_path = OUTPUT_DIR.joinpath(output_name)
        if output_path.exists():
            shutil.rmtree(output_path)
        o = _Msgtopdf(msg)
        o.file_name = '0 EMAIL'
        o.save_path = output_path
        o.email2pdf()
        o.msg = None

        msg.close(OL_DISCARD)
        msg = None
        msg_path.unlink()
    finally:
        if msg is not None:
            msg.close(OL_DISCARD)
            msg = None

    if output_path is not None:
        counter = 0
        for name in os.listdir(output_path):
            path = output_path.joinpath(name)
            if path.suffix.lower() == '.pdf':
                if name.startswith('0 '):
                    continue
                counter += 1
                path.rename(path.parent.joinpath(
                    f"{counter} {path.name}"
                ))
                continue
            try:
                path.unlink()
            except Exception as e:
                logging.debug(e)


def handle_input_dir():

    for name in os.listdir(INPUT_DIR):
        if not name.endswith('.msg'):
            continue

        try:
            handle_msg_file(INPUT_DIR.joinpath(name))
        except Exception as e:
            logging.error(f"Can't handle input msg-file {name}.")
            logging.debug(e)


def main():
    init()

    read_clients_file()
    handle_input_dir()


if __name__ == '__main__':
    main()
