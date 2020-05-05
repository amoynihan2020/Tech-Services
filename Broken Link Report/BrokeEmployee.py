# Broken link report, updated 9/25/2019
# created by cclunie
# I probably miss you guys
# Maybe text me idk
# 603-303-5005



'''
    Updated by Adam Moynihan on 12/4/2019
    
    We made changes to add columns in response to scs-10124
    Also had to add a new line of code to determine if the link_type 
    is local. 
'''

'''
Added the find_parent method so we can get the Titles of steps parent procedure. 
and add them to the end report 
'''
import pandas as pd
from bs4 import BeautifulSoup
import requests
import re
from sys import argv
import time
import xlsxwriter as excel


def make_df(json):
    df = pd.read_json(json, dtype=True)
    #added everythin past updated_by on 12/4
    df = df.filter(['content_type', 'title', 'body', 'id', 'category', 'updated_at', 'updated_by', 'source', 'do_not_display_in_search', 
                    'published', 'readable_id', 'procedure_id'])
    id_lst = list(df['id'])
    df = df.loc[df['body'].isnull() == False]
    df['updated_at'].dt.tz_localize(None)
    return df, id_lst


def check_link(link):
    print(link)
    try:
        code = requests.get(link, timeout=5)
    except requests.exceptions.Timeout:
        return 'timeout'
    except requests.exceptions.ConnectionError:
        return 'connection failed'
    except requests.exceptions.MissingSchema:
        return 'invalid URL'
    except requests.exceptions.TooManyRedirects:
        return 'exceeded redirect limit'
    except requests.exceptions.InvalidSchema:
        return 'invalid url schema'
    return code

def format_excel(dataFrame, name_of_sheet,workbook_name):
    
    workSheet = workbook_name.add_worksheet(name_of_sheet)
    cell_format = workbook_name.add_format({'bold': True}) 
    cell_format.set_bg_color('#000066')
    cell_format.set_font_color('#FFFFFF')
    cell_format.set_font_name('Calibri')
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    
    blue = workbook_name.add_format()
    blue.set_bg_color('#ccecff')
    blue.set_border(2)

    white = workbook_name.add_format()
    white.set_border(2)
    

    exRow = 0
    workSheet.write(0,0, 'Content Title', cell_format)
    workSheet.write(0,1, 'Content Type', cell_format)
    workSheet.write(0,2, 'System ID', cell_format)
    workSheet.write(0,3, 'Content ID', cell_format)
    workSheet.write(0,4, 'Procedure ID', cell_format)
    workSheet.write(0, 5, 'Parent Title' , cell_format)
    workSheet.write(0,6, 'Category', cell_format)
    workSheet.write(0,7, 'Updated At', cell_format)
    workSheet.write(0,8, 'Update By', cell_format)
    workSheet.write(0,9, 'Link', cell_format)
    workSheet.write(0,10, 'Link HTML', cell_format)
    workSheet.write(0,11, 'Status', cell_format)
    workSheet.write(0,12, 'Type', cell_format)
    workSheet.write(0,13, 'Source', cell_format)
    workSheet.write(0,14, 'Linked Text', cell_format)
    workSheet.write(0,15, 'Not Shown In Search', cell_format)
    workSheet.write(0,16, 'published', cell_format)
    workSheet.write(0,17, 'Fixed By SilverCloud', cell_format)
    workSheet.write(0,18, 'Comments', cell_format)
    workSheet.write(0,19, 'Updated URL', cell_format)
    
    for index, rows in dataFrame.iterrows():
        if ((exRow+1) % 2 == 1):
           if 'title_parent' not in dataFrame.columns:
                  workSheet.write_string(exRow+1,0, str(dataFrame.loc[index]['Content Title'] ),blue)
                  workSheet.write_string(exRow+1,1, str(dataFrame.loc[index]['Content_Type']  ),blue)
                  workSheet.write_string(exRow+1,2, str(dataFrame.loc[index]['System ID']  ),blue)
                  workSheet.write_string(exRow+1,3, str(dataFrame.loc[index]['Content ID'] ) ,blue)
                  workSheet.write_string(exRow+1,4, str(dataFrame.loc[index]['Procedure ID'] ) ,blue)
                  workSheet.write_string(exRow+1,5, str(dataFrame.loc[index]['Parent Title'] ) ,blue)
                  workSheet.write_string(exRow+1,6, str(dataFrame.loc[index]['Category']),blue)
                  workSheet.write_string(exRow+1,7, str(dataFrame.loc[index]['Updated_At']  ),blue)
                  workSheet.write_string(exRow+1,8, str(dataFrame.loc[index]['Updated_by']  ),blue)
                  workSheet.write_string(exRow+1,9, str(dataFrame.loc[index]['Link']),blue)
                  workSheet.write_string(exRow+1,10, str(dataFrame.loc[index]['Link_HTML']  ),blue)
                  workSheet.write_string(exRow+1,11, str(dataFrame.loc[index]['Status']  ),blue)
                  workSheet.write_string(exRow+1,12, str(dataFrame.loc[index]['type']  ),blue)
                  workSheet.write_string(exRow+1,13, str(dataFrame.loc[index]['Source']  ),blue)
                  workSheet.write_string(exRow+1,14, str(dataFrame.loc[index]['Link_text']  ),blue)
                  workSheet.write_string(exRow+1,15, str(dataFrame.loc[index]['Not Shown in Search']  ),blue)
                  workSheet.write_string(exRow+1,16, str(dataFrame.loc[index]['Published']),blue)
                  workSheet.write_blank(exRow+1,17, None,blue)
                  workSheet.write_blank(exRow+1,18, None,blue)
                  workSheet.write_blank(exRow+1,19, None,blue)
                  exRow+=1
           else:
                      
                      workSheet.write_string(exRow+1,0, str(dataFrame.loc[index]['Content Title'] ),blue)
                      workSheet.write_string(exRow+1,1, str(dataFrame.loc[index]['Content_Type']  ),blue)
                      workSheet.write_string(exRow+1,2, str(dataFrame.loc[index]['System ID']  ),blue)
                      workSheet.write_string(exRow+1,3, str(dataFrame.loc[index]['Content ID'] ) ,blue)
                      workSheet.write_string(exRow+1,4, str(dataFrame.loc[index]['Procedure ID'] ) ,blue)
                      workSheet.write_string(exRow+1,5, str(dataFrame.loc[index]['title_parent'] ) ,blue)
                      workSheet.write_string(exRow+1,6, str(dataFrame.loc[index]['Category']),blue)
                      workSheet.write_string(exRow+1,7, str(dataFrame.loc[index]['Updated_At']  ),blue)
                      workSheet.write_string(exRow+1,8, str(dataFrame.loc[index]['Updated_by']  ),blue)
                      workSheet.write_string(exRow+1,9, str(dataFrame.loc[index]['Link']),blue)
                      workSheet.write_string(exRow+1,10, str(dataFrame.loc[index]['Link_HTML']  ),blue)
                      workSheet.write_string(exRow+1,11, str(dataFrame.loc[index]['Status']  ),blue)
                      workSheet.write_string(exRow+1,12, str(dataFrame.loc[index]['type']  ),blue)
                      workSheet.write_string(exRow+1,13, str(dataFrame.loc[index]['Source']  ),blue)
                      workSheet.write_string(exRow+1,14, str(dataFrame.loc[index]['Link_text']  ),blue)
                      workSheet.write_string(exRow+1,15, str(dataFrame.loc[index]['Not Shown in Search']  ),blue)
                      workSheet.write_string(exRow+1,16, str(dataFrame.loc[index]['Published']),blue)
                      workSheet.write_blank(exRow+1,17, None,blue)
                      workSheet.write_blank(exRow+1,18, None,blue)
                      workSheet.write_blank(exRow+1,19, None,blue)
                      exRow+=1
                     
                
      
        else:
            if 'title_parent' not in dataFrame.columns:
                workSheet.write_string(exRow+1,0, str(dataFrame.loc[index]['Content Title'] ),white)
                workSheet.write_string(exRow+1,1, str(dataFrame.loc[index]['Content_Type']  ),white)
                workSheet.write_string(exRow+1,2, str(dataFrame.loc[index]['System ID']  ),white)
                workSheet.write_string(exRow+1,3, str(dataFrame.loc[index]['Content ID'] ) ,white)
                workSheet.write_string(exRow+1,4, str(dataFrame.loc[index]['Procedure ID'] ) ,white)
                workSheet.write_string(exRow+1,5, str(dataFrame.loc[index]['Parent Title'] ) ,blue)
                workSheet.write_string(exRow+1,6, str(dataFrame.loc[index]['Category']),white)
                workSheet.write_string(exRow+1,7, str(dataFrame.loc[index]['Updated_At']  ),white)
                workSheet.write_string(exRow+1,8, str(dataFrame.loc[index]['Updated_by']  ),white)
                workSheet.write_string(exRow+1,9, str(dataFrame.loc[index]['Link']), white)
                workSheet.write_string(exRow+1,10, str(dataFrame.loc[index]['Link_HTML']  ), white)
                workSheet.write_string(exRow+1,11, str(dataFrame.loc[index]['Status']  ), white)
                workSheet.write_string(exRow+1,12, str(dataFrame.loc[index]['type']  ), white)
                workSheet.write_string(exRow+1,13, str(dataFrame.loc[index]['Source']  ), white)
                workSheet.write_string(exRow+1,14, str(dataFrame.loc[index]['Link_text']  ), white)
                workSheet.write_string(exRow+1,15, str(dataFrame.loc[index]['Not Shown in Search']  ), white)
                workSheet.write_string(exRow+1,16, str(dataFrame.loc[index]['Published']),white)
                workSheet.write_blank(exRow+1,17, None,white)
                workSheet.write_blank(exRow+1,18, None,white)
                workSheet.write_blank(exRow+1,19, None,white)
                exRow+=1
            else:
               
                workSheet.write_string(exRow+1,0, str(dataFrame.loc[index]['Content Title'] ),white)
                workSheet.write_string(exRow+1,1, str(dataFrame.loc[index]['Content_Type']  ),white)
                workSheet.write_string(exRow+1,2, str(dataFrame.loc[index]['System ID']  ),white)
                workSheet.write_string(exRow+1,3, str(dataFrame.loc[index]['Content ID'] ) ,white)
                workSheet.write_string(exRow+1,4, str(dataFrame.loc[index]['Procedure ID'] ) ,white)
                workSheet.write_string(exRow+1,5, str(dataFrame.loc[index]['title_parent'] ) ,blue)
                workSheet.write_string(exRow+1,6, str(dataFrame.loc[index]['Category']),white)
                workSheet.write_string(exRow+1,7, str(dataFrame.loc[index]['Updated_At']  ),white)
                workSheet.write_string(exRow+1,8, str(dataFrame.loc[index]['Updated_by']  ),white)
                workSheet.write_string(exRow+1,9, str(dataFrame.loc[index]['Link']), white)
                workSheet.write_string(exRow+1,10, str(dataFrame.loc[index]['Link_HTML']  ), white)
                workSheet.write_string(exRow+1,11, str(dataFrame.loc[index]['Status']  ), white)
                workSheet.write_string(exRow+1,12, str(dataFrame.loc[index]['type']  ), white)
                workSheet.write_string(exRow+1,13, str(dataFrame.loc[index]['Source']  ), white)
                workSheet.write_string(exRow+1,14, str(dataFrame.loc[index]['Link_text']  ), white)
                workSheet.write_string(exRow+1,15, str(dataFrame.loc[index]['Not Shown in Search']  ), white)
                workSheet.write_string(exRow+1,16, str(dataFrame.loc[index]['Published']),white)
                workSheet.write_blank(exRow+1,17, None,white)
                workSheet.write_blank(exRow+1,18, None,white)
                workSheet.write_blank(exRow+1,19, None,white)
                exRow+=1
   
    workSheet.set_row(0,30)
    workSheet.set_landscape()
    workSheet.set_page_view()
    workSheet.set_header('&C&25&"Bold"Internal Broken Link Report')
    workSheet.set_footer('&L&"Calibri Italic" SilverCloud Inc. - Confidential Proprietary &R&G', {'image_right': 'sclogo.png'})
    workSheet.set_footer('&L&"Calibri Italic" SilverCloud Inc. - Confidential Proprietary &R&G', {'image_right': 'sclogo.png'})
  
    return workSheet
def get_links(body):
    bs = BeautifulSoup(body, 'html.parser')
    a_tags = bs.find_all('a')
    with_hrefs = bs.find_all('a', href=True)
    no_href = [x for x in a_tags if x not in with_hrefs]
    return with_hrefs, no_href

def find_parent(master_json, df_to_merge):
  
    df1 = pd.read_json(master_json)
    
    #only finds procedures 
    find_steps = df1.filter(['title', 'category', 'procedure_id', 'content_type', 'id','body'])
    find_procedures = df1.filter(['title', 'category', 'procedure_id', 'content_type', 'id'])
   
    #finding the steps in the DataFrame
    for index, row  in find_steps.iterrows():
        if row['content_type'] == 'procedure':
           
            find_steps = find_steps.drop(index)
    
    #Finding the Procedures in the Step
    for index, row in find_procedures.iterrows():
        if row['content_type'] == 'step':
           
            #gets rid of emppty rows
            find_procedures = find_procedures.drop(index)
   #mergin the two procedures to get them to match the parent procedure with the procedures id
    if 'procedure_id' in find_steps.columns:
        merge_pls  = pd.merge(find_steps, find_procedures,  left_on ='procedure_id', right_on = 'id', suffixes=('_step','_parent'))
        # filtering to get the parents title and ID
        merge_pls = merge_pls.filter(['title_parent', 'id_parent'])
        #Dropping duplicates  
        merge_pls.drop_duplicates(keep='first',inplace = True)
        # merging the master df to the one that contains ID and titles
        merge_pls = pd.merge(df_to_merge, merge_pls, left_on = 'Procedure ID', right_on = 'id_parent', how = 'left' )
        #filtering the columns to be included in the report
        final_df = merge_pls.filter(['Content Title', 'Content_Type', 'System ID', 'Content ID', 'Procedure ID', 'title_parent', 'Category', 'Updated_At', 'Updated_by',
                      'Link', 'Link_HTML', 'Status', 'type', 'Source', 'Link_text', 'Not Shown in Search', 'Published', 'Fixed By SilverCloud',
                      'Comments', 'Updated URL'])
        return final_df
    else:
        return df_to_merge
    
    
    
    
def process_links(df, intranet, all_ids):
  
    #Determines the order of the columns in the  final report Content Title will be first 
    cols = ['Content Title',
            'Content_Type',
            'System ID',
            'Content ID',
            'Procedure ID',
            'Parent Title',
            'Category',
            'Updated_At',
            'Updated_by',
            'Link',
            'Link_HTML',
            'Status',
            'type',
            'Source',
            'Link_text',
            'Not Shown in Search',
            'Published',
            'Fixed By SilverCloud',
            'Comments',
            'Updated URL']
    df2 = pd.DataFrame(columns=cols)
    visited = {}
    missing_href = []
    pattern = re.compile(r"=[5][0-9a-f]{23}")
    doc_pattern = re.compile(r"[5][0-9a-f]{23}")
    for index, row in df.iterrows():
        links = get_links(row['body'])
        missing_href.append(links[1])
        for link in links[0]:
            lnk = link['href']
            if len(lnk) > 4 or 'vm.getDocuments' in link:
                if hash(lnk) not in visited:
                    search = pattern.search(str(link))
                    if intranet is not None and intranet in lnk.lower():
                        link_type = 'local'
                    elif 'file:' in lnk or 'mailto:' in lnk:
                        link_type = 'local'
                    elif 'vm.getDocuments' in str(link):
                        link_type = 'sc-doc'
                    elif search is not None:
                        link_type = 'sc'
                    else:
                        link_type = 'external'
                    if link_type == 'sc':
                        if search is not None and search.group(0).replace('=', '') not in all_ids:
                            if 'procedure_id' in df.columns:
                               
                                df2 = df2.append({'Content Title': row['title'],
                                              'Content_Type': row['content_type'],
                                              'System ID': row['id'],
                                              'Content ID': row['readable_id'],
                                              'Procedure ID':row['procedure_id'],
                                              
                                             
                                              'Category': ', '.join(row['category']),
                                              'Updated_At': row['updated_at'],
                                              'Updated_by': row['updated_by'],
                                              'Link': lnk,
                                              'Source': row['source'],
                                              'Link_HTML': str(link),
                                              'Link_text': link.text,
                                              'Published':str(row['published']),
                                              'Not Shown in Search':str(row['do_not_display_in_search']),
                                              'Status': 'Not in KB',
                                              'type': link_type}, ignore_index=True)
                            else:
                               
                               df2 = df2.append({'Content Title': row['title'],
                                              'Content_Type': row['content_type'],
                                              'System ID': row['id'],
                                              'Content ID': row['readable_id'],
                                            #  'Procedure ID':row['procedure_id'],
                                             
                                             
                                              'Category': ', '.join(row['category']),
                                              'Updated_At': row['updated_at'],
                                              'Updated_by': row['updated_by'],
                                              'Link': lnk,
                                              'Source': row['source'],
                                              'Link_HTML': str(link),
                                              'Link_text': link.text,
                                              'Published':str(row['published']),
                                              'Not Shown in Search':str(row['do_not_display_in_search']),
                                              'Status': 'Not in KB',
                                              'type': link_type}, ignore_index=True)
                    elif link_type == 'sc-doc':
                        search2 = doc_pattern.search(str(link))
                        if search2 is not None and search2.group(0) not in all_ids:
                            if 'procedure_id' in df.columns:
                                
                                df2 = df2.append({'Content Title': row['title'],
                                              'Content_Type': row['content_type'],
                                              'System ID': row['id'],
                                              'Content ID': row['readable_id'],
                                              'Procedure ID':row['procedure_id'],
                                              
                                              'Category': ', '.join(row['category']),
                                              'Updated_At': row['updated_at'],
                                              'Updated_by': row['updated_by'],
                                              'Link': search2.group(0),
                                              'Source':row['source'],
                                              'Link_HTML': str(link),
                                              'Link_text':link.text,
                                              'Published':str(row['published']),
                                              'Not Shown in Search':str(row['do_not_display_in_search']),
                                              'Status': 'Not in KB',
                                              'type': link_type}, ignore_index=True)
                            else:
                               df2 = df2.append({'Content Title': row['title'],
                                              'Content_Type': row['content_type'],
                                              'System ID': row['id'],
                                              'Content ID': row['readable_id'],
                                            #  'Procedure ID':row['procedure_id'],
                                             
                                              'Category': ', '.join(row['category']),
                                              'Updated_At': row['updated_at'],
                                              'Updated_by': row['updated_by'],
                                              'Link': lnk,
                                              'Source': row['source'],
                                              'Link_HTML': str(link),
                                              'Link_text': link.text,
                                              'Published':str(row['published']),
                                              'Not Shown in Search':str(row['do_not_display_in_search']),
                                              'Status': 'Not in KB',
                                              'type': link_type}, ignore_index=True)
                    elif link_type == 'local':
                        if 'procedure_id' in df.columns:
                            df2 = df2.append({'Content Title': row['title'],
                                          'Content_Type': row['content_type'],
                                          'System ID': row['id'],
                                          'Content ID':row['readable_id'],
                                          'Procedure ID':row['procedure_id'],
                                          
                                          'Category': ', '.join(row['category']),
                                          'Updated_At': row['updated_at'],
                                          'Updated_by': row['updated_by'],
                                          'Link': lnk,
                                          'Source':row['source'],
                                          'Link_HTML': str(link),
                                          'Link_text':link.text,
                                          'Published':str(row['published']),
                                          'Not Shown in Search':str(row['do_not_display_in_search']),
                                          'Status': 'Unknown',
                                          'type': link_type}, ignore_index=True)
                        else:
                           
                           df2 = df2.append({'Content Title': row['title'],
                                              'Content_Type': row['content_type'],
                                              'System ID': row['id'],
                                              'Content ID': row['readable_id'],
                                             # 'Procedure ID':row['procedure_id'],
                                             
                                              'Category': ', '.join(row['category']),
                                              'Updated_At': row['updated_at'],
                                              'Updated_by': row['updated_by'],
                                              'Link': lnk,
                                              'Source': row['source'],
                                              'Link_HTML': str(link),
                                              'Link_text': link.text,
                                              'Published':str(row['published']),
                                              'Not Shown in Search':str(row['do_not_display_in_search']),
                                              'Status': 'Not in KB',
                                              'type': link_type}, ignore_index=True)
                    elif link_type == 'external':
                        if lnk.startswith('#') is False:
                            status = check_link(lnk)
                        else:
                            continue
                        errors = ['timeout', 'connection failed', 'invalid URL', 'exceeded redirect limit', 'URL Invalid' , 'Expired Certificate']
                        if type(status) != str:
                            if status is not None and status not in errors:
                                status = status.status_code
                                if status == '' or type(status) != int:
                                    status = 'No Response'
                        print(f'{lnk}, status: {str(status)}')
                        if 'procedure_id' in df.columns:
                            
                            df2 = df2.append({'Content Title': row['title'],
                                          'Content_Type': row['content_type'],
                                          'System ID': row['id'],
                                          'Content ID': row['readable_id'],
                                          'Procedure ID': row['procedure_id'],
                                         
                                          'Category': ', '.join(row['category']),
                                          'Updated_At': row['updated_at'],
                                          'Updated_by': row['updated_by'],
                                          'Link': lnk,
                                          'Source':row['source'],
                                          'Link_HTML': str(link),
                                          'Link_text':link.text,
                                          'Published':str(row['published']),
                                          'Not Shown in Search':str(row['do_not_display_in_search']),
                                          'Status': status,
                                          'type': link_type}, ignore_index=True)
                        else:
                            df2 = df2.append({'Content Title': row['title'],
                                              'Content_Type': row['content_type'],
                                              'System ID': row['id'],
                                              'Content ID': row['readable_id'],
                                              #'Procedure ID':row['procedure_id'],
                                             
                                              'Category': ', '.join(row['category']),
                                              'Updated_At': row['updated_at'],
                                              'Updated_by': row['updated_by'],
                                              'Link': lnk,
                                              'Source': row['source'],
                                              'Link_HTML': str(link),
                                              'Link_text': link.text,
                                              'Published':str(row['published']),
                                              'Not Shown in Search':str(row['do_not_display_in_search']),
                                              'Status': 'Not in KB',
                                              'type': link_type}, ignore_index=True)
    
                        visited[hash(lnk)] = (status, link_type)
         
                    else:
                        print(f'You shouldn\'t be seeing this... {link}')
                else:
                    if 'procedure_id' in df.columns:
                        df2 = df2.append({'Content Title': row['title'],
                                      'Content_Type': row['content_type'],
                                      'System ID': row['id'],
                                      'Content ID':row['readable_id'],
                                      'Procedure ID':row['procedure_id'],
                                      
                                      'Category': ', '.join(row['category']),
                                      'Updated_At': row['updated_at'],
                                      'Updated_by': row['updated_by'],
                                      'Link': lnk,
                                      'Source':row['source'],
                                      'Link_HTML': str(link),
                                      'Link_text':link.text,
                                      'Published':str(row['published']),
                                      'Not Shown in Search':str(row['do_not_display_in_search']),
                                      'Status': visited[hash(lnk)][0],
                                      'type': visited[hash(lnk)][1]}, ignore_index=True)
                    else:
                            df2 = df2.append({'Content Title': row['title'],
                                              'Content_Type': row['content_type'],
                                              'System ID': row['id'],
                                              'Content ID': row['readable_id'],
                                              #'Procedure ID':row['procedure_id'],
                                             
                                              'Category': ', '.join(row['category']),
                                              'Updated_At': row['updated_at'],
                                              'Updated_by': row['updated_by'],
                                              'Link': lnk,
                                              'Source': row['source'],
                                              'Link_HTML': str(link),
                                              'Link_text': link.text,
                                              'Published':str(row['published']),
                                              'Not Shown in Search':str(row['do_not_display_in_search']),
                                             'Status': visited[hash(lnk)][0],
'type': visited[hash(lnk)][1]}, ignore_index=True)
    return df2


def main_app(csv, report_name, intranet=None):
    start = time.time()
    make = make_df(csv)
    df = make[0]
    df = process_links(df, intranet, make[1])
    df2 = df.loc[df['Status'] != 200]
   
    external_df = df2.loc[df2['type'] == 'external']
    sc_df = df2.loc[(df2['type'] == 'sc') | (df2['type'] == 'sc-doc')]
    intra_df = df2.loc[df2['type'] == 'local']
   
   # writer = pd.ExcelWriter(f'{report_name}.xlsx', engine='xlsxwriter', options ={'remove_timezone': True})
    workbook = excel.Workbook(report_name +'.xlsx', {'strings_to_urls':False}) 
    
    if 'Procedure ID' in external_df.columns:
         
         external_parent=find_parent(csv,external_df)
         format_excel(external_parent, 'External Links', workbook )
       
    else:
        format_excel(external_df, 'External Links', workbook)
    
   
    if 'Procedure ID' in sc_df.columns: 
       sc_parent=find_parent(csv,sc_df)
       format_excel(sc_parent, 'SilverCloud Links', workbook)
    else:
        format_excel(sc_df, 'SilverCloud Links', workbook)    
 #   external_df.to_excel(writer, sheet_name='External_Links', index=False)
  #  sc_df.to_excel(writer, sheet_name='SilverCloud_Links', index=False)
    if len(intra_df) > 0:
         if 'Procedure ID' in intra_df.columns:
             intra_parent = find_parent(csv,intra_df)
             format_excel(intra_parent, 'Internal Links', workbook)
   
   #     intra_df.to_excel(writer, sheet_name='Internal_Links', index=False)
   
    workbook.close()
    print(f'Link report completed in {(time.time() - start) / 60} minutes.')
    return

if __name__ == '__main__':
    if len(argv) > 3:
        main_app(argv[1], argv[2], argv[3])
    elif len(argv) == 3:
        main_app(argv[1], argv[2])
    else:
        file = input('Please enter the all content JSON file name: ')
        name = input('Please enter the name you would like the report to export as. (no extension): ')
        intra = input('Please enter the intranet address (optional, leave blank if none):')
        if intra != '':
            main_app(file, name, intra)
        else:
            main_app(file, name)