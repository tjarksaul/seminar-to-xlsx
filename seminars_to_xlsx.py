#!/usr/bin/env python3
import requests
import pandas as pd
import urllib.parse

from config import glied_id, ids, username, password

s = requests.Session()


def load_and_write_data():
    dfs = []
    canceled = []

    for id in ids:
        print(f"Trying {id} for glied {glied_id}")

        xls_content = get_xlsx(glied_id, id)
        dfs.append(parse_excel(xls_content))
        canceled_content = get_xlsx(glied_id, id, status=3)
        canceled.append(parse_excel(canceled_content))

    df = pd.concat([d for d in dfs])
    df.sort_values('Anm-Dat', axis=0, inplace=True)
    df.to_excel('kampfrichter.xlsx')

    cdf = pd.concat([d for d in canceled])
    cdf.sort_values('Anm-Dat', axis=0, inplace=True)
    cdf.to_excel('kampfrichter_abgesagt.xlsx')


def parse_excel(data):
    cdf = pd.read_excel(data, engine="openpyxl")
    count = len(cdf.index)

    print(f"Found {count} entries")

    return cdf


def login():
    parameters = {'auth[user]': username, 'auth[pass]': password}
    r = s.post('https://dlrg.net/', parameters)
    if r.status_code != 200:
        raise Exception("Couldn't log in")


def get_xlsx(glied_id, id, status = 0):
    '''Downloads a seminar XLSX
    Arguments:
    glied_id:   ID of Gliederung
    id:         ID of seminar
    status:     Which participant statuses to download.
                0: active
                3: canceled
    '''
    url = f"https://dlrg.net/apps/seminar?page=loadDokumente&format=pdf&edvnummer={glied_id}&id={id}&noheader=1"
    parameters = {
        'dokumentListeTyp': 'xls',
        'dokumentListeRolleList[]': '1',
        'dokumentListeStatusList[]': f"{status}",
        'dokumentListeTnstatusBestaetigtDurchTeilnehmer': '',
        'dokumentListeTnstatusBestaetigtDurchVerwalter': '',
        'dokumentListeTnstatusBestaetigtDurchGliederung': '',
        'dokumentListeTnstatusBezahlt': '',
        'dokumentListeTnstatusTeilgenommen': '',
        'dokumentListeTnstatusBestanden': '',
        'dokumentListeSortierung': 'anmeldenummer',
    }

    r = s.post(url, parameters)

    filename = '.'.join(
        r.headers['Content-Disposition'].split('=', 1)[1].strip('"').split('.')[:-1])
    filename = urllib.parse.unquote(filename)
    print(f"Downloaded '{filename}'")

    if r.status_code != 200:
        raise (f"Couldn't get id {id} for glied {glied_id}")

    return r.content


def main():
    login()

    load_and_write_data()


if __name__ == '__main__':
    main()
