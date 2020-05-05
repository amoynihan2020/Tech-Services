'''
Concurrent Fuzzy Duplicate Report
Created by Chris Clunie for SilverCloud inc.
'''

import pandas as pd
from fuzzywuzzy import fuzz
from sys import argv
from bs4 import BeautifulSoup
import time
from multiprocessing import Process, Manager
from os import getcwd, listdir


def warning_text(x):
    print('CAUTION: This script will use all available processing power.\nYour computer will be useless for several hours.\nIntended for use outside of operational hours.\nPress ctrl+c at any time to abort.')
    time.sleep(x)


def check_inputname(inputname):
    all_files = listdir(getcwd())
    if inputname not in all_files:
        return True
    else:
        return False


def check_inputname_type(inputname):
    if '.csv' in inputname:
        return False
    else:
        return True


def check_filename(filename):
    all_files = listdir(getcwd())
    if filename in all_files:
        return True
    else:
        return False


def check_filetype(filename):
    if '.xlsx' in filename:
        return False
    else:
        return True


def start_up(argv):
    warning_text(5)
    if len(argv) < 2:
        csv_name = ''
    else:
        csv_name = argv[1]
    if len(argv) < 3:
        output_filename = ''
    else:
        output_filename = argv[2]
    if len(argv) > 3:
        ratio_threshold = int(argv[3])
    else:
        ratio_threshold = 80
    input_name = True
    input_ext = True
    threshold = False
    filename = True
    excel = True
    while input_name == True or input_ext == True or filename == True or excel == True or threshold is False:
        input_name = check_inputname(csv_name)
        input_ext = check_inputname_type(csv_name)
        filename = check_filename(output_filename)
        excel = check_filetype(output_filename)
        if type(ratio_threshold) == int and ratio_threshold > 0 and ratio_threshold <= 100:
            threshold = True
        else:
            ratio_threshold = int(input('Threshold must be integer between 1 and 100\nPlease choose a new threshold:'))
        if filename == True:
            output_filename = str(input('Filename must be unique\nPlease choose a new file name:'))
        if excel == True:
            output_filename = str(input('Filename must end in .xlsx file extension\nPlease choose a new file name:'))
            filename = check_filename(output_filename)
        if input_name == True:
            csv_name = str(input('Input file must be in current directory\nPlease check to make sure the files exists in this location\nor choose a new file name:'))
        if input_ext == True:
            output_filename = str(input('Input file must be a CSV and end in a .csv file extension\nPlease choose a new file name:'))
            filename = check_filename(output_filename)
    print('finished start up')
    return (output_filename, ratio_threshold, csv_name)


def get_text(df):
    body_df = df.loc[df['body'].str.len() > 0]
    for index, row in body_df.iterrows():
        try:
            bs = BeautifulSoup(row['body'], 'html.parser')
        except:
            print('Unable to parse {0} content type: {1}'.format(row['title'], row['content_type']))
            continue
        body_df.loc[index, 'plain_text'] = str(bs.text)
    return body_df


def ratio_check(text1, df, threshold, dfs):
    columns1 = ['title', 'type', 'id', 'procedure_id', 'do_not_display_in_search', 'match_title', 'match_type',
               'match_id', 'match_procedure_id', 'match_do_not_display_in_search', 'match_percentage']
    df2 = pd.DataFrame(columns=columns1)
    for index, row in df.iterrows():
        if row['content_type'] != 'document' and row['content_type'] != 'link' and len(row['plain_text']) > 0:
            if index != text1[0]:
                ratio = fuzz.ratio(text1[1], row['plain_text'])
                if ratio >= threshold:
                    title = text1[2]
                    type = text1[3]
                    ident = text1[4]
                    try:
                        pro_ident = text1[5]
                    except NameError:
                        pro_ident = 'None'
                    not_displayed = text1[6]
                    match_name = row['title']
                    match_type = row['content_type']
                    match_ident = row['id']
                    if row['procedure_id'] == '':
                        match_pro_ident = 'None'
                    else:
                        match_pro_ident = row['procedure_id']
                    match_do_not_display_in_search = row['do_not_display_in_search']
                    df2 = df2.append({'title': title, 'type': type, 'id': ident, 'procedure_id': pro_ident,
                                      'do_not_display_in_search': not_displayed, 'match_title': match_name,
                                      'match_type': match_type, 'match_id': match_ident,
                                      'match_procedure_id': match_pro_ident,
                                      'match_do_not_display_in_search': match_do_not_display_in_search,
                                      'match_percentage': ratio}, ignore_index=True)
    dfs[text1[4]] = df2
    return df2


if __name__ == '__main__':
    start = start_up(argv)
    output_filename = start[0]
    threshold = start[1]
    start_time = time.time()
    manager = Manager()
    dfs = manager.dict()
    raw_df = pd.read_csv(start[2])
    df = get_text(raw_df)
    empty_df = df.loc[df['plain_text'].str.len() == 0]
    df = df.loc[df['plain_text'].str.len() > 0]
    empty_df.to_excel(f"empty_bodies_report_{str(time.time()).replace('.', '_')}.xlsx")
    columns = ['title', 'type', 'id', 'procedure_id', 'do_not_display_in_search', 'match_title', 'match_type',
                'match_id', 'match_procedure_id', 'match_do_not_display_in_search', 'match_percentage']
    final_df = pd.DataFrame(columns=columns)
    counter = 0
    for index, row in df.iterrows():
        arg = [index, row['plain_text'], row['title'], row['content_type'], row['id'], row['procedure_id'], row['do_not_display_in_search']]
        process = Process(target=ratio_check, args=(arg, df, threshold, dfs))
        process.daemon = True
        process.start()
        counter += 1
        percent_complete = counter / len(df)
        print(f"{percent_complete:.2%} Complete, {counter} of {len(df)}, Processing {row['title']}, type: {row['content_type']}")
    process.join()
    for key, value in dfs.items():
        final_df = pd.concat([final_df, value])
    final_df.to_excel(output_filename, index=False)
    elapsed_time = (time.time() - start_time) / 60
    print('\nDuplicate check complete!\nTotal elapsed time: {0:.2f} minutes'.format(elapsed_time))
