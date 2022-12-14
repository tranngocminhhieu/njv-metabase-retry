# For building GUI
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import tkinter.scrolledtext as scrolledtext
import webbrowser # To open folder macOS and Windows in the same syntax
# import subprocess # To open folder but different syntax for both OS
import threading

# For deep coping user input
import copy

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
import shutil

# For fun :))
import random

# from multiprocessing import Process

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
            params = json.loads(payload)['parameters']
            for p in params:
                try:
                    del p['id']
                except:
                    pass
            params = json.dumps(params)
            return params

# Checking valid cookie and quesion
def check_valid_cookie_url(cookie, question, question_url):
    try:
        res = requests.post(f'https://metabase.ninjavan.co/api/card/{question}/query/json',
                            headers={'Content-Type': 'application/json', 'X-Metabase-Session': cookie}, timeout=3)
        if res.status_code == 404:
            return 'This question does not exist.'
        elif res.status_code == 401:
            return 'This cookie is incorrect or has expired.'
        elif not '?' in question_url:
            try:
                json_data = res.json()
                if 'error' in json_data:
                    return json_data['error']
            except:
                return True
    except:
        return True
    return True

# Printing random emoji for fun
def random_emoji(feeling='happy'):
    list_happy = [*'????????????????????????????????????????????????????????????????????????????????????????????????']
    list_sad = [*'????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????']
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
        self.current_version = 1.1
        try:
            self.latest_version = float(requests.get('https://pastebin.com/raw/0uJU5URe').text)
        except:
            self.latest_version = self.current_version
        self.check_version = self.current_version >= self.latest_version

        # Creating app data folder
        self.metabase_retry_path = str(Path.home() / 'Documents' / 'metabase_retry')
        if not os.path.exists(self.metabase_retry_path):
            os.makedirs(self.metabase_retry_path)

        # App name
        root.title("NinjaVan Metabase Retry")

        # Area 1: Area for users to interact with the application.
        mainframe = ttk.Frame(root, padding="15 15 15 15")
        mainframe.grid(column=0, row=0, sticky=(W, E, S, N))
        mainframe.columnconfigure(2, weight=1) # Column 2 will auto max width

        # Area 1: Input Question URL
        ttk.Label(mainframe, text="Question URL").grid(column=1, row=1, sticky=W)
        self.question_url = StringVar()
        self.input_question_url = ttk.Entry(mainframe, textvariable=self.question_url, width=70) # Only set width for one Entry, anothers Entry will responsive with sticky.
        self.input_question_url.grid(column=2, row=1, sticky=(W,E), columnspan=2)

        # Area 1: Input Cookie
        ttk.Label(mainframe, text="Cookie").grid(column=1, row=2, sticky=W)
        self.cookie = StringVar()
        self.input_cookie = ttk.Entry(mainframe, textvariable=self.cookie)
        self.input_cookie.grid(column=2, row=2, sticky=(W,E), columnspan=2)


        # Area 1: Input save_as
        ttk.Label(mainframe, text="Save as (xlsx, csv)").grid(column=1, row=3, sticky=W)
        self.save_as = StringVar()
        self.input_save_as = ttk.Entry(mainframe, textvariable=self.save_as)
        self.input_save_as.grid(column=2, row=3, sticky=(W,E))
        ttk.Button(mainframe, text="Browse", command=lambda: threading.Thread(target=self.save_as_file, daemon=True).start()).grid(column=3, row=3, sticky=E)

        # Area 1: Input Retry times
        ttk.Label(mainframe, text="Retry times (0 = ???)").grid(column=1, row=4, sticky=W)
        self.retry_times = StringVar()
        self.input_retry_times = ttk.Entry(mainframe, width=4, textvariable=self.retry_times)
        self.input_retry_times.grid(column=2, row=4, sticky=(W))

        # Area 1: Button
        button_frame = ttk.Frame(mainframe)
        button_frame.grid(column=2, row=4, sticky=E, columnspan=2)
        # Area 1: Button to run
        ttk.Button(button_frame, text="Run & Download", command=lambda : threading.Thread(target=self.handle_app, daemon=True).start(), default="active").grid(column=3, row=1, sticky=E)
        # Area 1: Button to open folder
        ttk.Button(button_frame, text="Open folder", command=lambda : threading.Thread(target=self.open_folder, daemon=True).start()).grid(column=2, row=1, sticky=E, padx=10)
        ttk.Button(button_frame, text="Delete input", command=lambda : threading.Thread(target=self.delete_input, daemon=True).start()).grid(column=1, row=1, sticky=E)

        # Area 2: Area to print the process.
        # Area 2: Text box
        # self.output = Text(mainframe, height=20, width=70, pady=5, padx=5, wrap=WORD, blockcursor=TRUE, borderwidth=2, relief="sunken")
        self.output = scrolledtext.ScrolledText(mainframe, height=20, pady=5, padx=5, wrap=WORD, blockcursor=TRUE, borderwidth=2, relief="sunken")
        self.output.grid(column=1, row=5, sticky=(W, E, S, N), columnspan=3)
        self.output.tag_config('red', foreground='red')
        self.output.tag_config('green', foreground='green')
        self.output.tag_config('yellow', foreground='yellow')
        self.output.tag_config('orange', foreground='orange')
        self.output.tag_config('grey', foreground='grey')
        self.output.insert(END, f'Hello Ninjas! {random_emoji()}')

        ttk.Label(mainframe, text="Powered by KAM - Analyst").grid(column=1, row=6, sticky=W, columnspan=3)

        # Area 1: Auto save_as and fill user input.
        # Auto fill Question URL
        try:
            with open(f'{self.metabase_retry_path}/question_url.txt', 'r') as f:
                self.input_question_url.insert(0, f.read())
        except:
            with open(f'{self.metabase_retry_path}/question_url.txt', 'w') as f:
                pass

        # Auto fill Cookie
        try:
            with open(f'{self.metabase_retry_path}/cookie.txt', 'r') as f:
                self.input_cookie.insert(0, f.read())
        except:
            with open(f'{self.metabase_retry_path}/cookie.txt', 'w') as f:
                pass

        # Auto fill Save as
        try:
            with open(f'{self.metabase_retry_path}/save_as.txt', 'r') as f:
                self.input_save_as.insert(0, f.read())
        except:
            with open(f'{self.metabase_retry_path}/save_as.txt', 'w') as f:
                pass

        # Auto Retry times
        try:
            with open(f'{self.metabase_retry_path}/retry_times.txt', 'r') as f:
                self.input_retry_times.insert(0, f.read())
        except:
            self.input_retry_times.insert(0, '0')
            with open(f'{self.metabase_retry_path}/retry_times.txt', 'w') as f:
                f.write('0')

        # Adding padding for all child elements
        for child in mainframe.winfo_children():
            child.grid_configure(padx=5, pady=5)

        # Auto focus on a field
        self.input_question_url.focus()

        # Blind Enter button to do run
        root.bind("<Return>", lambda zzz : threading.Thread(target=self.handle_app, daemon=True).start()) # threading help avoid freezing app daemon=True help stop process after closing window

        # Notifying user to update app
        if not self.check_version:
            self.output.insert(END, f'\nPlease update the app to the latest version {self.latest_version}. The current version is {self.current_version}.', 'red')

        self.counter = {}

    def save_as_file(self):
        folder_selected = filedialog.asksaveasfilename(typevariable=StringVar, filetypes=[('Excel file', '*.xlsx'), ('CSV file', '*.csv')])
        self.save_as.set(folder_selected)

    def open_folder(self):
        folder_path = os.path.split(self.save_as.get())[0]
        
        if os.path.isdir(folder_path):
            # subprocess.Popen(['open', folder_path])
            webbrowser.open('file:///' + folder_path)
        else:
            self.output.insert(END, f'\n{folder_path} does not exist. {random_emoji(feeling="sad")}', 'red')
            self.output.see(END)

    def delete_input(self):
        self.input_question_url.delete(0, END)
        self.input_save_as.delete(0, END)
        self.input_cookie.delete(0, END)
        self.retry_times.set(0)
        try:
            shutil.rmtree(self.metabase_retry_path)
            self.output.insert(END, f'\nThe input data has been deleted. {random_emoji()}', 'green')
            self.output.see(END)
        except:
            pass

    # Metabase query function
    @retry(wait=wait_fixed(5))
    def metabase_question_query(self, cookie, question, short_question_url, params='[]', *args):

        self.counter[short_question_url] += 1
        self.output.insert(END,f'\n{datetime.now().strftime("%Y-%m-%d %H:%M")}: Start query {short_question_url} ({self.counter[short_question_url]})')
        self.output.see(END)
        try:
            res = requests.post(f'https://metabase.ninjavan.co/api/card/{question}/query/json?parameters={params}',
                                headers={'Content-Type': 'application/json', 'X-Metabase-Session': cookie}, timeout=900)
        except Exception as e:
            self.output.insert(END, f'\n{datetime.now().strftime("%Y-%m-%d %H:%M")}: {short_question_url} - Timeout! {random_emoji(feeling="sad")}')
            self.output.insert(END, f'\n{e}', 'grey')
            self.output.see(END)
            time.sleep(10)
            raise e

        # Raise error
        if not res.ok:
            self.output.insert(END, f'\n{datetime.now().strftime("%Y-%m-%d %H:%M")}: {short_question_url} - {res.status_code} {res.reason}, please consider stopping. {random_emoji(feeling="sad")}', 'orange')
            self.output.see(END)
            time.sleep(10)
            raise ConnectionError
        res.raise_for_status()

        try:
            data = res.json()
        except Exception as e:
            self.output.insert(END, f'\n{datetime.now().strftime("%Y-%m-%d %H:%M")}: {short_question_url} - The data is too large or corrupted after downloading, please consider re-selecting the filter with less data and try again. {random_emoji(feeling="sad")}', 'orange')
            self.output.insert(END, f'\n{e}', 'grey')
            self.output.see(END)
            time.sleep(10)
            raise Exception('The data is too large')
        try:
            error = data['error']  # Too many queued queries for "admin", Query exceeded the maximum execution time limit of 5.00m
            self.output.insert(END, f'\n{datetime.now().strftime("%Y-%m-%d %H:%M")}: {short_question_url} - {error} {random_emoji(feeling="sad")}')
            self.output.see(END)
            time.sleep(10)
            raise Exception(error)
        except:
            pass
        # Data
        _df = pd.DataFrame(data)
        self.output.insert(END, f'\n{datetime.now().strftime("%Y-%m-%d %H:%M")}: Finish query {short_question_url} {random_emoji()}', 'green')
        self.output.see(END)
        return _df

    def run_and_download(self, cookie, question, retries, save_as, short_question_url, params='[]', *args):
        _cookie = copy.deepcopy(cookie)
        _question = copy.deepcopy(question)
        _retries = copy.deepcopy(retries)
        _save_as = copy.deepcopy(save_as)
        _params = copy.deepcopy(params)
        _short_question_url = copy.deepcopy(short_question_url)
        self.output.insert(END, f'\nStart running query {_short_question_url} and save as {_save_as}.')
        self.output.see(END)
        if retries > 0:
            df = self.metabase_question_query.retry_with(wait=wait_fixed(5), stop=stop_after_attempt(_retries))(self, cookie=_cookie, question=_question, params=_params, short_question_url=_short_question_url)
        else:
            df = self.metabase_question_query(cookie=_cookie, question=_question, params=_params, short_question_url=_short_question_url)
        # Save
        try:
            if _save_as.split('.')[-1] == 'xlsx':
                df.to_excel(_save_as, index=False)
                self.output.insert(END, f'\nThe Excel file has been saved as {_save_as}! {random_emoji()}', 'green')
            elif _save_as.split('.')[-1] == 'csv':
                df.to_csv(_save_as, index=False, encoding='utf-8-sig')
                self.output.insert(END, f'\nThe CSV file has been saved as {_save_as}! {random_emoji()}', 'green')
            self.output.see(END)
        except Exception as e:
            self.output.insert(END, f'\n{datetime.now().strftime("%Y-%m-%d %H:%M")}: {_short_question_url} - {e}', 'red')
            self.output.see(END)

    def handle_app(self):
        question_url = copy.deepcopy(self.question_url.get())
        cookie = copy.deepcopy(self.cookie.get())
        save_as = copy.deepcopy(self.save_as.get())
        retries = copy.deepcopy(self.retry_times.get())

        # Notifying user to update app
        if not self.check_version:
            self.output.insert(END, f'\nDon\'t be stubborn. Please download the new version, Ninjas! {random_emoji()}', 'red')

        # Notifying user to update app
        if not self.check_version:
            self.output.insert(END, f'\nDon\'t be stubborn. Please download the new version, Ninjas! {random_emoji()}', 'red')

        # Check valid infomation
        # Check URL
        check_question_url = False
        if question_url == '':
            self.output.insert(END, f'\nPlease fill in the Question URL', 'red')
        elif not 'https://metabase.ninjavan.co/question/' in question_url:
            self.output.insert(END, '\nThis question URL is invalid. Please paste the URL copied from your browser.', 'red')
        else:
            check_question_url = True
            if not os.path.exists(self.metabase_retry_path):
                os.makedirs(self.metabase_retry_path)
            with open(f'{self.metabase_retry_path}/question_url.txt', 'w') as f:
                f.write(question_url)

        # Check Cookie
        check_cookie = False
        if cookie == '':
            self.output.insert(END, '\nPlease fill in the Cookie', 'red')
        elif len(cookie.split('-')) != 5:
            self.output.insert(END, '\nThis cookie is invalid. Please fill the metabase.SESSION cookie.\nAdd this extension to Chrome > Go to Metabase > Click the extension > Copy metabase.SESSION.\nhttps://chrome.google.com/webstore/detail/cookie-tab-viewer/fdlghnedhhdgjjfgdpgpaaiddipafhgk', 'red')
        else:
            check_cookie = True
            if not os.path.exists(self.metabase_retry_path):
                os.makedirs(self.metabase_retry_path)
            with open(f'{self.metabase_retry_path}/cookie.txt', 'w') as f:
                f.write(cookie)

        # Check Save as
        check_save_as = False
        folder_path = os.path.split(save_as)[0]
        if save_as == '':
            self.output.insert(END, '\nPlease fill in the Save as.', 'red')
        elif not os.path.isdir(folder_path):
            self.output.insert(END, f'\n{folder_path} does not exist.', 'red')
        elif not '.csv' in save_as and not '.xlsx' in save_as:
            self.output.insert(END, '\nPlease include .xlsx or .csv in file name.', 'red')
        else:
            check_save_as = True
            if not os.path.exists(self.metabase_retry_path):
                os.makedirs(self.metabase_retry_path)
            with open(f'{self.metabase_retry_path}/save_as.txt', 'w') as f:
                f.write(save_as)

        # Check Retry times
        check_retry_times = False
        if retries == '':
            self.output.insert(END, '\nPlease fill in the Rety times. A valid number will be greater than or equal to 0. And 0 means infinite.', 'red')
        else:
            try:
                if int(retries) < 0:
                    self.output.insert(END, f'\nPlease enter a positive number.', 'red')
                else:
                    check_retry_times = True
                    if not os.path.exists(self.metabase_retry_path):
                        os.makedirs(self.metabase_retry_path)
                    with open(f'{self.metabase_retry_path}/retry_times.txt', 'w') as f:
                        f.write(retries)
            except:
                self.output.insert(END, f'\nPlease enter a positive number.', 'red')

        # Check valid Cookie and Question online
        question = get_question(question_url)
        check_status = False
        if check_question_url and check_cookie and check_save_as and check_retry_times and self.check_version:
            self.output.insert(END, f'\nVerifying your information on the server.')
            self.output.see(END)
            check_status = check_valid_cookie_url(cookie, question, question_url)
            if check_status != True:
                self.output.insert(END, f'\n{check_status}', 'red')

        # Scroll down
        self.output.see(END)

        # Start query retry
        if check_question_url and check_cookie and check_save_as and check_retry_times and check_status==True and self.check_version:
            self.output.insert(END, '\nYour information is valid, start processing.', 'green')
            self.output.see(END)

            # Get parameters
            if not '?' in question_url:
                params = '[]'
            else:
                self.output.insert(END, '\nGetting parameters from payload, Chrome window will automatically close once done.')
                self.output.see(END)
                params = get_params(cookie=cookie, question_url=question_url)

            # Run query
            retries = int(retries)

            params_url = '?' + question_url.split('?')[-1] if '?' in question_url else ''
            short_question_url = str(question) + params_url

            self.counter[short_question_url] = 0 # To print retry times

            start_time = time.time() # To calculate how long it took
            self.run_and_download(cookie=cookie, question=question, retries=retries, save_as=save_as, params=params, short_question_url=short_question_url)

            # Print how long it took
            end_time = time.time()
            took_time = round(end_time - start_time)
            if took_time < 60:
                self.output.insert(END, f'\nIt took {took_time} {"second" if took_time <= 1 else "seconds"}.')
            elif took_time < 3600:
                self.output.insert(END, f'\nIt took {took_time//60} {"minute" if took_time//60 <= 1 else "minutes"} {took_time%60} {"second" if took_time%60 <= 1 else "seconds"}')
            else:
                self.output.insert(END, f'\nIt took {took_time//3600} {"hour" if took_time//3600 <= 1 else "hours"} {(took_time%3600) // 60} {"minute" if (took_time%3600) // 60 <= 1 else "minutes"} {(took_time%3600) % 60} {"second" if (took_time%3600) % 60 <= 1 else "seconds"}')
            self.output.see(END)

root = Tk()
Metabase_Retry(root)

# Auto fixed size and center window on the screen
root.eval('tk::PlaceWindow . center')
# Disable user resize window
root.resizable(False, False)

root.mainloop()