## How to convert py to app
```
pip install pyinstaller
pyinstaller --onefile -w gui.py
```
## How to force users to download the latest version
We can change the version in https://pastebin.com/raw/0uJU5URe, the application will check the version when opening automatically. If the version is not the latest version, the application can not run anymore.

## Changelog
### 1.0
1. Checking the user input infomations.
- The app will check valid question URL.
- The app will check valid cookie.
- The app will check valid status of question URL and cookie.
- The app will check valid directory.
- The app will check valid number of retry times.

2. Storing the application data for future use.
- The app will create a directory in Documents/metabase_retry.
- THe app will save the user input if it is vaid.
- The app will save the parameters of each question URL.

3. Getting data.
- The app will get parameters of the question URL if it is not exists.
- The app will get json data with retry function.

4. Saving file.
- The app can save the query result to Excel and CSV with directory of user.

5. Interacting with user.
- The app will print error, process to the Text box. Random emoji will be included for some case.
- User can click Button or Enter to run.