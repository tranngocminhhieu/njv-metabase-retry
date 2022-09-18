# For building GUI
import sys
import tkinter.messagebox
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import threading

# For getting parameters
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# For saving Excel & CSV
import pandas as pd

# For requesting HTTP
import json
import requests
from tenacity import *

# For printing datatime and sleeping
from datetime import datetime
import time

# For saving app data
import os
from pathlib import Path
import sqlite3

# For fun :))
import random

# Getting the question ID from question URL
def get_question(question_url):
    question = question_url.split('/question/')[-1].split('-')[0]
    return question

# Getting the parameters from question URL
def get_params(cookie, question_url):
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get(question_url)
    driver.add_cookie({'name': 'metabase.SESSION', 'value': cookie})
    driver.get(question_url)
    time.sleep(5)
    for request in driver.requests:
        if request.method == 'POST' and '/api/card/' in request.url:
            payload = request.body
            params = json.dumps(json.loads(payload)['parameters'])
            return params

# Checking valid cookie and quesion
def check_valid_cookie_url(cookie, question):
    try:
        res = requests.post(f'https://metabase.ninjavan.co/api/card/{question}/query/json',
                            headers={'Content-Type': 'application/json', 'X-Metabase-Session': cookie}, timeout=3)
        if res.status_code == 404:
            return 'This question does not exist.'
        elif res.status_code == 401:
            return 'This cookie is incorrect or has expired.'
        else:
            return True
    except:
        return True

# Checking if parameters is exists in local database
def check_params_db(question_url, db_path):
    if not '?' in question_url:
        return '[]'
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute('CREATE TABLE IF NOT EXISTS params (question_url NVACHAR PRIMARY KEY, parameters NVACHAR)')
        cur.execute(f'SELECT parameters FROM params WHERE question_url = \'{question_url}\'')
        sql_result = cur.fetchall()
        if sql_result == []:
            parameters = False
        else:
            parameters = sql_result[0][0]
    return parameters

# Updating parameters to local database for feature using
def update_params_db(question_url, params, db_path):
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute('CREATE TABLE IF NOT EXISTS params (question_url NVACHAR PRIMARY KEY, parameters NVACHAR)')
        try:
            cur.execute(f'INSERT INTO params(question_url, parameters) VALUES (\'{question_url}\', \'{params}\')')
        except:
            cur.execute(f'UPDATE params SET parameters = \'{params}\' WHERE question_url = \'{question_url}\'')

# Printing random emoji for fun
def random_emoji(feeling='happy'):
    list_happy = [*'😀😃😄😁😆😅😂🤣😍🥰😘😙😚😋😛😝😜🤗😈🥳🤩🏆🎉🎊']
    list_sad = [*'😔😟😕🙁😖😫😩🥺😢😭😤😡🤬😨🤔😐😑🙄😧🥱😴😪😮‍💨😵😵‍💫🥴🤢🤧🤒🤕💩💥🔥🧨']
    if feeling=='happy':
        emoji = random.choice(list_happy)
    elif feeling=='sad':
        emoji = random.choice(list_sad)
    else:
        emoji = random.choice(list_happy + list_sad)
    return emoji

# Main app
class Metabase_Retry:
    def __init__(self, root):
        # The version help force user update app
        self.current_version = '1.0'
        try:
            self.latest_version = requests.get('https://pastebin.com/raw/0uJU5URe').text
        except:
            self.latest_version = self.current_version
        self.check_version = self.current_version == self.latest_version

        # Creating app data folder
        self.metabase_retry_path = str(Path.home() / 'Documents' / 'metabase_retry')
        if not os.path.exists(self.metabase_retry_path):
            os.makedirs(self.metabase_retry_path)

        # App name
        root.title("NinjaVan Metabase Retry")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # Area 1: Area for users to interact with the application.
        mainframe = ttk.Frame(root, padding="15 15 15 15")
        mainframe.grid(column=0, row=0, sticky=(W, E, S, N))
        # mainframe.grid_columnconfigure(0, weight=1)
        # mainframe.grid_rowconfigure(0, weight=1)

        # Area 1: Input Question URL
        ttk.Label(mainframe, text="Question URL").grid(column=1, row=1, sticky=W)
        self.question_url = StringVar()
        input_question_url = ttk.Entry(mainframe, textvariable=self.question_url, width=50) # Only set width for one Entry, anothers Entry will responsive with sticky.
        input_question_url.grid(column=2, row=1, sticky=(W,E), columnspan=2)

        # Area 1: Input Cookie
        ttk.Label(mainframe, text="Cookie").grid(column=1, row=2, sticky=W)
        self.cookie = StringVar()
        input_cookie = ttk.Entry(mainframe, textvariable=self.cookie)
        input_cookie.grid(column=2, row=2, sticky=(W,E), columnspan=2)

        # Area 1: Input save_as
        ttk.Label(mainframe, text="Save as (xlsx, csv)").grid(column=1, row=3, sticky=W)
        self.save_as = StringVar()
        input_save_as = ttk.Entry(mainframe, textvariable=self.save_as, width=40)
        input_save_as.grid(column=2, row=3, sticky=(W,E))
        ttk.Button(mainframe, text="Browse", command=lambda: threading.Thread(target=self.save_as_file, daemon=True).start()).grid(column=3, row=3, sticky=E)

        # Area 1: Input Retry times
        ttk.Label(mainframe, text="Retry times (0 = ∞)").grid(column=1, row=4, sticky=W)
        self.retry_times = StringVar()
        input_retry_times = ttk.Entry(mainframe, width=4, textvariable=self.retry_times)
        input_retry_times.grid(column=2, row=4, sticky=(W))

        # Area 1: Button
        button_frame = ttk.Frame(mainframe)
        button_frame.grid(column=2, row=4, sticky=E, columnspan=2)
        # Area 1: Button to run
        ttk.Button(button_frame, text="Run & Download", command=lambda : threading.Thread(target=self.handle_app, daemon=True).start(), default="active").grid(column=3, row=1, sticky=E)
        # Area 1: Button to stop
        # ttk.Button(button_frame, text="Stop").grid(column=2, row=1, sticky=E, padx=10)

        # Area 2: Area to print the process.
        # Area 2: Text box
        self.output = Text(mainframe, height=15, width=70, pady=5, padx=5, wrap=WORD, blockcursor=TRUE, borderwidth=2, relief="sunken")
        self.output.grid(column=1, row=5, sticky=(W, E, S, N), columnspan=3)
        self.output.tag_config('red', foreground='red')
        self.output.tag_config('green', foreground='green')
        self.output.tag_config('yellow', foreground='yellow')

        ttk.Label(mainframe, text="Powered by KAM - Analyst").grid(column=1, row=6, sticky=W, columnspan=3)

        # Area 1: Auto save_as and fill user input.
        # Auto fill Question URL
        try:
            with open(f'{self.metabase_retry_path}/question_url.txt', 'r') as f:
                input_question_url.insert(0, f.read())
        except:
            with open(f'{self.metabase_retry_path}/question_url.txt', 'w') as f:
                pass

        # Auto fill Cookie
        try:
            with open(f'{self.metabase_retry_path}/cookie.txt', 'r') as f:
                input_cookie.insert(0, f.read())
        except:
            with open(f'{self.metabase_retry_path}/cookie.txt', 'w') as f:
                pass

        # Auto fill Save as
        try:
            with open(f'{self.metabase_retry_path}/save_as.txt', 'r') as f:
                input_save_as.insert(0, f.read())
        except:
            with open(f'{self.metabase_retry_path}/save_as.txt', 'w') as f:
                pass

        # Auto Retry times
        try:
            with open(f'{self.metabase_retry_path}/retry_times.txt', 'r') as f:
                input_retry_times.insert(0, f.read())
        except:
            input_retry_times.insert(0, '0')
            with open(f'{self.metabase_retry_path}/retry_times.txt', 'w') as f:
                f.write('0')

        # Adding padding for all child elements
        for child in mainframe.winfo_children():
            child.grid_configure(padx=5, pady=5)

        # Auto focus on a field
        input_question_url.focus()

        # Blind Enter button to do run
        root.bind("<Return>", lambda zzz : threading.Thread(target=self.handle_app, daemon=True).start()) # threading help avoid freezing app daemon=True help stop process after closing window

        # Notifying user to update app
        if self.check_version == False:
            self.output.insert(END, f'Please update the app to the latest version {self.latest_version}. The current version is {self.current_version}.', 'red')

    def save_as_file(self):
        folder_selected = filedialog.asksaveasfilename(typevariable=StringVar, filetypes=[('Excel file', '*.xlsx'), ('CSV file', '*.csv')])
        self.save_as.set(folder_selected)

    # Metabase query function
    @retry(wait=wait_fixed(5))
    def metabase_question_query(self, cookie, question, param='[]', *args):
        self.counter += 1
        self.output.insert(END,f'\n{datetime.now().strftime("%Y-%m-%d %H:%M")}: Start query {question} ({self.counter})')
        self.output.see(END)
        res = requests.post(f'https://metabase.ninjavan.co/api/card/{question}/query/json?parameters={param}',
                            headers={'Content-Type': 'application/json', 'X-Metabase-Session': cookie}, timeout=600)
        # Raise error
        res.raise_for_status()
        if not res.ok:
            raise ConnectionError
        data = res.json()
        try:
            error = data['error']  # Too many queued queries for "admin", Query exceeded the maximum execution time limit of 5.00m
            self.output.insert(END, f'\n{datetime.now().strftime("%Y-%m-%d %H:%M")}: Error - {error} {random_emoji(feeling="sad")}')
            self.output.see(END)
            raise Exception(error)
        except:
            pass
        # Data
        _df = pd.DataFrame(data)
        self.output.insert(END, f'\n{datetime.now().strftime("%Y-%m-%d %H:%M")}: Finish query {question} {random_emoji()}')
        self.output.see(END)
        return _df

    def handle_app(self):
        # Notifying user to update app
        if self.check_version == False:
            self.output.insert(END, f'\nDon\'t be stubborn. Please download the new version, Ninjas! {random_emoji()}', 'red')

        # Check valid infomation
        # Check URL
        check_question_url = False
        if self.question_url.get() == '':
            self.output.insert(END, f'\nPlease fill in the Question URL', 'red')
        elif not 'https://metabase.ninjavan.co/question/' in self.question_url.get():
            self.output.insert(END, '\nThis question URL is invalid. Please paste the URL copied from your browser.', 'red')
        else:
            check_question_url = True
            with open(f'{self.metabase_retry_path}/question_url.txt', 'w') as f:
                f.write(self.question_url.get())

        # Check Cookie
        check_cookie = False
        if self.cookie.get() == '':
            self.output.insert(END, '\nPlease fill in the Cookie', 'red')
        elif len(self.cookie.get().split('-')) != 5:
            self.output.insert(END, '\nThis cookie is invalid. Please fill the metabase.SESSION cookie.\nAdd this extension to Chrome > Go to Metabase > Click the extension > Copy metabase.SESSION.\nhttps://chrome.google.com/webstore/detail/cookie-tab-viewer/fdlghnedhhdgjjfgdpgpaaiddipafhgk', 'red')
        else:
            check_cookie = True
            with open(f'{self.metabase_retry_path}/cookie.txt', 'w') as f:
                f.write(self.cookie.get())

        # Check Save as
        check_save_as = False
        if self.save_as.get() == '':
            self.output.insert(END, '\nPlease fill in the Save as.', 'red')
        elif not '.csv' in self.save_as.get() and not '.xlsx' in self.save_as.get():
            self.output.insert(END, '\nPlease include .xlsx or .csv for file name.', 'red')
        else:
            check_save_as = True
            with open(f'{self.metabase_retry_path}/save_as.txt', 'w') as f:
                f.write(self.save_as.get())

        # Check Retry times
        check_retry_times = False
        if self.retry_times.get() == '':
            self.output.insert(END, '\nPlease fill in the Rety times. A valid number will be greater than or equal to 0. And 0 means infinite.', 'red')
        else:
            try:
                if int(self.retry_times.get()) < 0:
                    self.output.insert(END, '\nPlease enter a positive number.', 'red')
                else:
                    check_retry_times = True
                    with open(f'{self.metabase_retry_path}/retry_times.txt', 'w') as f:
                        f.write(self.retry_times.get())
            except:
                self.output.insert(END, '\nPlease enter a positive number.', 'red')

        # Check valid Cookie and Question online
        question = get_question(self.question_url.get())
        cookie = self.cookie.get()
        check_status = False
        if check_question_url and check_cookie and check_save_as and check_retry_times and self.check_version == True:
            if check_question_url and check_cookie:
                check_status = check_valid_cookie_url(cookie, question)
                if check_status != True:
                    self.output.insert(END, f'\n{check_status}', 'red')

        # Scroll down
        self.output.see(END)

        # Start query retry
        if check_question_url and check_cookie and check_save_as and check_retry_times and check_status==True and self.check_version==True:
            self.output.insert(END, '\nValid information, start processing.', 'green')
            self.output.see(END)

            # Get parameters
            check_params = check_params_db(question_url=self.question_url.get(), db_path=f'{self.metabase_retry_path}/params.db')
            if check_params != False:
                self.output.insert(END, '\nThis Question URL already has parameters stored.', 'green')
                self.output.see(END)
                params = check_params
            else:
                self.output.insert(END, '\nThis Question-URL does not have parameters stored. The application will check with the browser. Once done, the browser will automatically close, and the parameters of this URL will be stored for future use.')
                self.output.see(END)
                params = get_params(cookie=self.cookie.get(), question_url=self.question_url.get())
                update_params_db(question_url=self.question_url.get(), params=params, db_path=f'{self.metabase_retry_path}/params.db')

            # Run query
            self.output.insert(END, '\nStart running the query.')
            self.output.see(END)

            retries = int(self.retry_times.get())
            self.counter = 0 # To print retry times
            start_time = time.time() # To calculate how long it took
            if retries > 0:
                df = self.metabase_question_query.retry_with(wait=wait_fixed(5), stop=stop_after_attempt(retries))(self, cookie=cookie, question=question, param=params)
            else:
                df = self.metabase_question_query(cookie=cookie, question=question, param=params)

            # Save
            try:
                if self.save_as.get().split('.')[-1] == 'xlsx':
                    df.to_excel(self.save_as.get(), index=False)
                    self.output.insert(END, f'\nThe Excel file has been saved as {self.save_as.get()} ! {random_emoji()}', 'green')
                elif self.save_as.get().split('.')[-1] == 'csv':
                    df.to_csv(self.save_as.get(), index=False)
                    self.output.insert(END, f'\nThe CSV file has been saved as {self.save_as.get()}! {random_emoji()}', 'green')
                self.output.see(END)
            except Exception as e:
                self.output.insert(END, f'\n{e}', 'red')
                self.output.see(END)

            # Print how long it took
            end_time = time.time()
            took_time = round(end_time - start_time)
            if took_time < 60:
                self.output.insert(END, f'\nIt took {took_time} seconds.')
            else:
                self.output.insert(END, f'\nIt took {took_time//60} minutes {took_time%60} seconds')
            self.output.see(END)


root = Tk()
Metabase_Retry(root)

# Auto fixed size and center window on the screen
root.eval('tk::PlaceWindow . center')
# Disable user resize window
root.resizable(False, False)

# Auto fixed size and center window on the screen (Old way)
# Window size
# window_width = root.winfo_width()
# window_height = root.winfo_height()
#
# # Get the screen dimension
# screen_width = root.winfo_screenwidth()
# screen_height = root.winfo_screenheight()
#
# # Gind the center point
# center_x = int(screen_width/2 - window_width / 2)
# center_y = int(screen_height/2 - window_height / 2)
#
# # Center screen
# root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

root.mainloop()
