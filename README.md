## How to convert py to app
To convert a Python script to a macOS application, run this shell on a macOS device. And vice versa, to convert Python script to a Windows application, run this shell on a Windows device.

### macOS
```
pip install pyinstaller

pyinstaller --clean --add-data 'ca.crt:seleniumwire' --add-data 'ca.key:seleniumwire' --onefile --windowed --icon="icon.icns" gui.py
```

### Windows
Install pyinstaller from GitHub instead of pip to avoid Windows Defender false positive flagging (warning virus).
read: https://python.plainenglish.io/pyinstaller-exe-false-positive-trojan-virus-resolved-b33842bd3184

Download pyinstaller on https://github.com/pyinstaller/pyinstaller/releases

```
# cd to pyinstaller folder

# activate your env and run this shell to install pyinstaller
python.exe setup.py install

# cd to source
pyinstaller --clean --add-data "ca.crt;seleniumwire" --add-data "ca.key;seleniumwire" --onefile --windowed --icon="icon.ico" gui.py
```
`ca.crt` and `ca.key` copy from `site-packages/seleniumwire`.

## How to force users to download the latest version
We can change the version in https://pastebin.com/raw/0uJU5URe, the application will check the version when opening automatically. If the version is not the latest version, the application can not run anymore.

## Changelog
### 1.1
- User can delete input (Documents/metabase_retry) with a button.
- User can scroll through the text box.
- User can quickly add multiple queries without waiting for the text box's response.
- Fix the wrong counter when running multiple queries.
- Remove save parameters to the local host feature.
- Check valid parameters online if '?' is not the in the question URL.
- Print query error in more detail.


### 1.0
1. Checking the user input information.
- The app will check valid question URL.
- The app will check valid cookie.
- The app will check valid status of the question URL and cookie.
- The app will check valid directory.
- The app will check valid number of retry times.

2. Storing the application data for future use.
- The app will create a directory in Documents/metabase_retry.
- The app will save the user input if it is valid.
- The app will save the parameters of each question URL.

3. Getting data.
- The app will get the parameters of the question URL if it does not exist.
- The app will get JSON data with the retry function.
- The app can run multiple queries at the same time.

4. Saving file.
- The app can save the query result to Excel and CSV with a directory of users.
- Downloads is the default directory if user input does not contain any directory.

5. Interacting with users.
- The app will print an error, and the processing to the Text box. Random emojis will be included for cases.
- User can click Button or Enter to run.
- User can open a folder with a button.