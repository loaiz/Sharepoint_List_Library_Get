# -*- coding: utf-8 -*-
"""
Created on Wed Aug  3 07:29:10 2022

@author: Yeferson Loaiza
"""




from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.fields.lookup_value import FieldLookupValue
from selenium.webdriver.support import expected_conditions as EC
from office365.sharepoint.client_context import ClientContext
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import pandas as pd
import win32com.client as client
import win32com.client
import os
import win32com
import datetime
import os.path
from time import localtime, strftime
import re
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import time
import warnings
import zipfile
import glob
import shutil
import sys, os


def dataframeSP(lista):
    client_id = "xxClientxx"
    client_secret = "xxClientxx"
    site_url = "url_site"
    #client context
    ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))
    
    target_list = ctx.web.lists.get_by_title(lista)
    paged_items = target_list.items.get_all().execute_query()
    columnas=list(pd.DataFrame.from_dict(paged_items[0].properties.items()).iloc[:,0])
    valores=list()
    
    for index, item in enumerate(paged_items):
        data=list(pd.DataFrame.from_dict(item.properties.items()).iloc[:,1])
        valores.append(data)
    resultado=pd.DataFrame(valores,columns=columnas)
    return resultado

def dataframeSpLibrary(relative_url):
    client_id = "xxClientxx"
    client_secret = "xxClientxx"
    site_url = "url_site"
    #client context
    ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))

    libraryRoot = ctx.web.get_folder_by_server_relative_path(relative_url)
    ctx.load(libraryRoot)
    ctx.execute_query()
    #if you want to get the folders within <sub_folder> 
    folders = libraryRoot.folders
    ctx.load(folders)
    ctx.execute_query()

    files = libraryRoot.files
    ctx.load(files)
    ctx.execute_query()
    
    #create a dataframe of the important file properties for me for each file in the folder
    df_files = pd.DataFrame(columns = ['Name', 'ServerRelativeUrl', 'TimeLastModified', 'ModTime'])
    for myfile in files:
        mod_time = time.strptime(myfile.properties['TimeLastModified'], '%Y-%m-%dT%H:%M:%SZ')
        df_dictionary = pd.DataFrame([{'Name': myfile.properties['Name'], 'ServerRelativeUrl': myfile.properties['ServerRelativeUrl'],'TimeLastModified': myfile.properties['TimeLastModified'], 'ModTime': mod_time}])
        df_files = pd.concat([df_files, df_dictionary], ignore_index=True)

    return df_files

#example function with parameters
pa = dataframeSpLibrary('NameLibrary')

x = dataframeSP('NameList')