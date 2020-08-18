# for writing to excel(xlsx) we will be needing XlsxWriter, please install it first if you don't have it!
import zipfile
import XlsxWriter
# to scrape a table from a webpage
from urllib.parse import urlparse, urlsplit
import requests
import pandas as pd
import os
from xlrd import open_workbook
import csv

# utility functions


def create_dir(dir):
    if not os.path.exists(dir):
        os.makedirs(dir)
        #print("Created Directory : ", dir)
        return dir


# convert the sheetname- remove _ and - , put title case and remove spaces


def modify_name(my_str):
    replaced = my_str.replace("_", " ").replace("-", " ")
    return replaced.title().replace(" ", "")


# get all tables from a url
def get_dataframes(url):
    html = requests.get(url).content
    df_list = pd.read_html(html)
    #print(len(df_list)," Dataframes Returned")
    return df_list

# if df is too small then don't add it


def filter_dfs(dfs_list, min_rows=10):
    new_dfs_list = []
    for each_df in dfs_list:
        if(len(each_df) > min_rows):
            new_dfs_list.append(each_df)
    return new_dfs_list

# to avoid InvalidWorksheetName: Excel worksheet name 'StatesAndUnionTerritoriesOfIndia1' must be <= 31 chars.


def crop_name(name, thres=29):
    if len(name) < thres:
        return name
    else:
        return name[:thres]

# to get first n elements from list only


def crop_list(lst, thres=28):
    if len(lst) < thres:
        return lst
    else:
        return lst[:thres]

# converts urls to dataframes to excel sheets
# get_max= get the maximum number of tables from each url
# min_rows= the minimum number of rows in each table to save it to the excel sheet
# crop_name_thres= some excel sheets can get quite huge sheet names which blows up the code
# so crop the sheet name for the better purpose


def urls_to_excel(urls, excel_path=None, get_max=10, min_rows=0, crop_name_thres=29):
    excel_path = os.path.join(os.getcwd(
    ), "Excel_Multiple_Sheets_Output.xlsx") if excel_path == None else excel_path
    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    i = 0
    num = 3
    for url in urls:
        parsed = urlsplit(url)
        sheet_name = parsed.path.split('/')[-1]
        if sheet_name == '':
            sheet_name = parsed.path.split('/')[-2]
        mod_sheet_name = crop_name(modify_name(
            sheet_name), thres=crop_name_thres)

        dfs_list = get_dataframes(url)
        filtered_dfs_list = filter_dfs(dfs_list, min_rows=min_rows)
        filtered_dfs_list = crop_list(filtered_dfs_list, thres=get_max)
        for each_df in filtered_dfs_list:
            #print("Parsing Excel Sheet ", " : ",str(i).zfill(num)+mod_sheet_name)
            i += 1
            each_df.to_excel(writer, sheet_name=str(
                i).zfill(num)+mod_sheet_name, index=True)
    writer.save()


def excel_to_csv(excel_path, csv_dir):
    create_dir(csv_dir)
    create_dir(os.path.dirname(excel_path))
    wb = open_workbook(excel_path)
    root_dir = csv_dir
    for i in range(wb.nsheets):
        sheet = wb.sheet_by_index(i)
        out_file_name = root_dir+"%s.csv" % (sheet.name.replace(" ", ""))
        #print("Parsing", sheet.name, " to : ", out_file_name)
        with open(out_file_name, "w") as file:
            writer = csv.writer(file, delimiter=",")
            #print (sheet, sheet.name, sheet.ncols, sheet.nrows)
            header = [cell.value for cell in sheet.row(0)]
            writer.writerow(header)
            for row_idx in range(1, sheet.nrows):
                row = [int(cell.value) if isinstance(cell.value, float) else cell.value
                       for cell in sheet.row(row_idx)]
                writer.writerow(row)


def runner():
    urls = ["https://www.macrotrends.net/countries/IND/india/gdp-growth-rate",
            "http://www.inwea.org/wind-energy-in-india/wind-power-potential",
            "https://en.wikipedia.org/wiki/List_of_districts_in_India",
            "https://en.wikipedia.org/wiki/List_of_Indian_people_by_net_worth",
            "https://en.wikipedia.org/wiki/States_and_union_territories_of_India",
            "https://en.wikipedia.org/wiki/List_of_governors-general_of_India",
            "https://en.wikipedia.org/wiki/List_of_Indian_independence_activists",
            "https://en.wikipedia.org/wiki/List_of_Indian_Grammy_Award_winners_and_nominees",
            "https://en.wikipedia.org/wiki/List_of_Indian_Academy_Award_winners_and_nominees",
            "https://en.wikipedia.org/wiki/List_of_highest-grossing_Indian_films"
            ]

    #print(len(urls), "Urls Found")

    urls_to_excel(urls, get_max=1, min_rows=10)

    excel_to_csv(
        excel_path='/content/Excel_Multiple_Sheets_Output.xlsx', csv_dir='/content/')


def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file))


def zipper(dir_path, zip_path):
    zipf = zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED)
    zipdir(dir_path, zipf)
    zipf.close()
