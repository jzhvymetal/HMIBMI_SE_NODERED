:: ***GLOBAL JS NODES LOCATION:=%APPDATA%\npm
:: ***NODERED_PROGRAM LOCATION:=C:\%HOMEPATH%\.node-red

:: ***REMOVE INSTALL
:: Uninstall Node.js from windows Add/Remove Programs
:: rmdir /S /Q %APPDATA%\npm & rmdir /S /Q C:\%HOMEPATH%\.node-red && rmdir /S /Q c:\NR_SERVICE

:: This program needs offline versions of following in the same directory:
::	 1:  se-node-red-palette_manager.zip
::	 2:  se-node-red-modbus.zip
:: It is required to unzip the file from link with password set in e-mail
::	The process of distribution of offline build has changed as per the below steps:
::	1. Open the registration form from the below URL
::       https://ecostruxure-data-expert-essential.se.app:3000/se-nodes
::	2. Enter the User Name(Mandatory), between 3- 32 characters.
::	3. Enter the User Email(Mandatory) with valid Customer email address with standard email format.
::	4. Select the required node(Mandatory) from the Nodes dropdown and click on Submit button.
::	5. Requested node will be zipped with password protection and download link will be send to 
::	   User Email address.(Download link will be valid only for 1 week.)
::	6. Password will be send to the provide User Email.
::	7. User can unzip the node file with the password he received in the mail.
::	8. Unzip the se-node-red-palette_manager.zip package with password.

:: *****NO PROXY       :  SYS_PROXY=DIRECT 
:: *****France         :  SYS_PROXY=http://gateway.schneider.zscaler.net:80/
:: *****Germany        :  SYS_PROXY=http://force-proxy-eur.pac.schneider-electric.com:80
:: *****North America  :  SYS_PROXY=http://gateway.schneider.zscaler.net:9480/
:: *****	 	          SYS_PROXY=http://gateway.zscaler.ne:80/
:: *****APAC           :  SYS_PROXY=http://gateway.schneider.zscaler.net:9480/
:: *****ZSCALER        :  SYS_PROXY=http://127.0.0.1:9000
SET SYS_PROXY=DIRECT
::*****NODEJS_VER=-1 will install latest version. 
::*****NODEJS_VER=10  will install latest major version. NodeJS 10 was tested with SE_MODBUS
::*****NODEJS_VER=10.15.0  will install exact version. 
::*****SE DOC REQUEST Node.js V14 or greater....Tested and worked with 20.x 
SET NODEJS_VER=-1
SET NR_SERVICE_NAME=Node-RED
SET NR_SERVICE_DIR=C:\NR_SERVICE
:: *****NODERED_VER=-1 will install latest version
:: *****NODERED_VER=0.20.* will install 20 latest version. 
:: *****SE DOC REQUEST 2.1.4 but still works with the latest 3.x
SET NODERED_VER=-1
SET ZIP_se-node-red-palette_manager=se-node-red-palette_manager.zip
SET ZIP_se-node-red-modbus=se-node-red-modbus.zip
::Used if newer version of NodeJS is used this will be renamed to NodeJs Installed
SET ZIP_latest_ver=windows-v16

::****************START OF SCRIPT DO NOT MODIFY**********************************************
@ECHO OFF
SETLOCAL EnableDelayedExpansion
CLS

CALL :INSTALL_PYTHON_MSI
CALL :REFREASH_PATH
CALL :INSTALL_NODEJS_MSI
CALL :REFREASH_PATH
CALL :INSTALL_NODE_RED
CALL :REFREASH_PATH
CALL :CREATE_START_NODERED_CMD
CALL :INSTALL_NSS_NODERED_SERVICE
CALL :NODERED_RUN_WAIT
CALL :NODERED_KILL_WINDOW
CALL :INSTALL_se-node-red-palette_manager
CALL :INSTALL_se-node-red-modbus
CALL :NODERED_RUN_WAIT
START "" http://127.0.0.1:1880
exit

GOTO :EXIT_FINAL

:NODERED_KILL_WINDOW
CALL taskkill /F /FI "WINDOWTITLE eq node-red" /T
GOTO :EOF

:NODERED_RUN_WAIT
START cmd /c %NR_SERVICE_DIR%\%NR_SERVICE_NAME%_START.cmd
:LOOP_FOR_PORT
	netstat -o -n -a | find "LISTENING" | find "1880"
	if %ERRORLEVEL% equ 0 goto PORT_FOUND
	ping /n 1 /w 2000 localhost >nul
    goto LOOP_FOR_PORT
:PORT_FOUND
GOTO :EOF

:INSTALL_se-node-red-palette_manager
IF EXIST "%temp%\%ZIP_se-node-red-palette_manager%" rmdir /S /Q "%temp%\%ZIP_se-node-red-palette_manager%"
ECHO UNZIPPING se-node-red-palette_manager
CALL :UnZipFile "%temp%\%ZIP_se-node-red-palette_manager%" "%~dp0%ZIP_se-node-red-palette_manager%"
cd /D "%temp%\%ZIP_se-node-red-palette_manager%"
ECHO INSTALLING se-node-red-palette_manager
CALL "%temp%\%ZIP_se-node-red-palette_manager%\se-node-red-palette_manager_offline_install.bat"
cd /D %~dp0
IF EXIST "%temp%\%ZIP_se-node-red-palette_manager%" rmdir /S /Q "%temp%\%ZIP_se-node-red-palette_manager%"
GOTO :EOF

:INSTALL_se-node-red-modbus
IF EXIST "%temp%\%ZIP_se-node-red-modbus%" rmdir /S /Q "%temp%\%ZIP_se-node-red-modbus%"
ECHO UNZIPPING se-node-red-modbus
CALL :UnZipFile "%temp%\%ZIP_se-node-red-modbus%" "%~dp0%ZIP_se-node-red-modbus%"
::GET NODEJS VER	
CALL "node" -v > %temp%\node_ver.tmp
SET /p SUB_INSTALLED_VER=<%temp%\node_ver.tmp
SET "Up2Sub=%SUB_INSTALLED_VER:v=%"
SET "SUB_INSTALLED_VER=%Up2Sub:.="&:"%"
ECHO %SUB_INSTALLED_VER%
IF NOT EXIST "%temp%\%ZIP_se-node-red-modbus%\files\windows-v%SUB_INSTALLED_VER%" (
	rename "%temp%\%ZIP_se-node-red-modbus%\files\%ZIP_latest_ver%" "windows-v%SUB_INSTALLED_VER%"
	)	
cd /D "%temp%\%ZIP_se-node-red-modbus%"
ECHO INSTALLING se-node-red-modbus
CALL npm install modbus-serial@8.0.5
CALL "%temp%\%ZIP_se-node-red-modbus%\se-node-red-modbus_offline_install.bat"
cd /D %~dp0
IF EXIST "%temp%\%ZIP_se-node-red-modbus%" rmdir /S /Q "%temp%\%ZIP_se-node-red-modbus%"
SET Up2Sub=
SET SUB_INSTALLED_VER=
GOTO :EOF

:INSTALL_NODE_RED
SET SUB_VER=
::CLEAR CACHE SO IF WILL NOT GIVE FALUSE POSTIVE
CALL npm config delete registry
CALL npm cache clean --force  >nul 2>&1 
IF NOT "!SYS_PROXY!"=="DIRECT" (
	CALL npm config set proxy !SYS_PROXY!
)
ECHO TESTING NPM INTERNET CONNECTION
CALL npm --fetch-retry-maxtimeout 11000 --fetch-retries 1 ping 2<&1 || GOTO :EXIT_NPM_INTERNET
ECHO NPM INTERNET CONNECTION SUCCESSFUL
	
IF NOT %NODERED_VER%==-1 SET SUB_VER=@!NODERED_VER!
ECHO INSTALLING node-red
call npm install -g --unsafe-perm node-red%SUB_VER%
SET SUB_VER=
GOTO :EOF

:INSTALL_PYTHON_MSI
CALL :GET_LATEST_PYTHON_VER

IF "!PROCESSOR_ARCHITECTURE!"=="x86" SET PYTHON_MSI_FILENAME=python-%PYTHON_VER%.exe
IF NOT "!PROCESSOR_ARCHITECTURE!"=="x86" SET PYTHON_MSI_FILENAME=python-%PYTHON_VER%-amd64.exe
ECHO DOWNLOADING python
CALL :WGET "https://www.python.org/ftp/python/%PYTHON_VER%/%PYTHON_MSI_FILENAME%" "%PYTHON_MSI_FILENAME%" 
ECHO INSTALLLING python
CALL %PYTHON_MSI_FILENAME% /passive InstallAllUsers=1 PrependPath=1 Include_pip=1  Include_launcher=1 AssociateFiles=1
ECHO %PYTHON_MSI_FILENAME%
IF EXIST %PYTHON_MSI_FILENAME% DEL /Q %PYTHON_MSI_FILENAME%
SET NODEJS_PYTHON_FILENAME=
GOTO :EOF

:GET_LATEST_PYTHON_VER
::LATEST VERSION
SET PYTHON_VER_FILE=PYTHON_VER.TXT
SET PYTHON_VER=-1
IF EXIST %PYTHON_VER_FILE% DEL /Q %PYTHON_VER_FILE%
IF %PYTHON_VER%==-1 CALL :WGET "https://www.python.org/doc/versions/" "%PYTHON_VER_FILE%"
IF NOT EXIST "%PYTHON_VER_FILE%" GOTO :EXIT_PYTHON_VER_FILE
for /f "delims=" %%a in ('findstr "<li>" %PYTHON_VER_FILE%') DO SET PYTHON_LINE="%%a" &	GOTO :PYTHONLINE_END_LOOP
:PYTHONLINE_END_LOOP
::Remove chars to make it easier to parse
set "PYTHON_LINE="%PYTHON_LINE:<=*%""
set "PYTHON_LINE="%PYTHON_LINE:>=*%""
set "PYTHON_LINE="%PYTHON_LINE:/=$%""
SET "Up2Sub=%PYTHON_LINE:*$release$=%"
SET "PYTHON_VER=%Up2Sub:$="&:"%"
SET PYTHON_LINE=
SET Up2Sub=
::REMOVE UNUSED FILES
IF EXIST %PYTHON_VER_FILE% DEL /Q %PYTHON_VER_FILE%
SET PYTHON_VER_FILE=
Echo %PYTHON_VER%
GOTO :EOF

:INSTALL_NODEJS_MSI
CALL :GET_LATEST_NODEJS_VER
IF "!PROCESSOR_ARCHITECTURE!"=="x86" SET NODEJS_MSI_FILENAME=node-v%NODEJS_VER%-x86.msi
IF NOT "!PROCESSOR_ARCHITECTURE!"=="x86" SET NODEJS_MSI_FILENAME=node-v%NODEJS_VER%-x64.msi
ECHO DOWNLOADING nodejs
CALL :WGET "https://nodejs.org/dist/v%NODEJS_VER%/%NODEJS_MSI_FILENAME%" "%NODEJS_MSI_FILENAME%" 
ECHO INSTALLING nodejs
MsiExec.exe /i "%NODEJS_MSI_FILENAME%" /passive
ECHO %NODEJS_MSI_FILENAME%
IF EXIST %NODEJS_MSI_FILENAME% DEL /Q %NODEJS_MSI_FILENAME%
SET NODEJS_MSI_FILENAME=
GOTO :EOF

:GET_LATEST_NODEJS_VER
::LATEST VERSION
SET NODEJS_VER_FILE=NODEJS_VER.TXT
SET NODELINE_FILE=NODELINE.TXT
IF EXIST %NODEJS_VER_FILE% DEL /Q %NODEJS_VER_FILE%
IF EXIST %NODELINE_FILE% DEL /Q %NODELINE_FILE%
IF %NODEJS_VER%==-1 CALL :WGET "http://nodejs.org/dist/latest/SHASUMS256.txt" "%NODEJS_VER_FILE%"

IF NOT %NODEJS_VER%==-1 (
	echo.%NODEJS_VER%|findstr /C:"." >nul 2>&1
	if not errorlevel 1 (
	   :: EXACT VERSION
	   GOTO :EOF
	) else (
	   :: LATEST MAIN VERSION
	   CALL :WGET "http://nodejs.org/dist/latest-v%NODEJS_VER%.x/SHASUMS256.txt" "%NODEJS_VER_FILE%"
	)
)
IF NOT EXIST %NODEJS_VER_FILE% GOTO :EXIT_NODEJS_VER_FILE
findstr /R "node-v.*.pkg" %NODEJS_VER_FILE% > %NODELINE_FILE%
SET /p NODELINE=<%NODELINE_FILE%
SET "Up2Sub=%NODELINE:*node-v=%"
SET "NODEJS_VER=%Up2Sub:.pkg="&:"%"
::REMOVE UNUSED FILES
IF EXIST %NODEJS_VER_FILE% DEL /Q %NODEJS_VER_FILE%
IF EXIST %NODELINE_FILE% DEL /Q %NODELINE_FILE%
::CLEAR ALL UNUSED VARS
SET NODELINE=
SET Up2Sub=
SET NODEJS_VER_FILE=
GOTO :EOF

:UnZipFile <ExtractTo> <newzipfile>
SET UnZipEXE="7za.exe"
IF NOT EXIST "%UnZipEXE%" (
	CALL :WGET "https://www.7-zip.org/a/7zr.exe" "test.tmp"
	rename "test.tmp" "7zr.exe"
	CALL :WGET "https://www.7-zip.org/a/7z2301-extra.7z" "7z2301-extra.7z" -y
	CALL 7zr x D:\_NodeRedTemp\7z2301-extra.7z 7za.exe -y
	DEL /Q 7z2301-extra.7z
	DEL /Q 7zr.exe
	)
IF NOT EXIST "%UnZipEXE%" (CALL :UnZipFileVBA %1 %2)
IF EXIST "%UnZipEXE%" ("%UnZipEXE%" x %2 -o%1)
SET UnZipEXE=
GOTO :EOF


:UnZipFileVBA <ExtractTo> <newzipfile>
SET UNZIP_VBS="%VBS_SAFE_PATH%\unzip.vbs"
IF EXIST %UNZIP_VBS% DEL /f /q %UNZIP_VBS%
>%UNZIP_VBS%  ECHO Set fso = CreateObject("Scripting.FileSystemObject")
>>%UNZIP_VBS% ECHO If NOT fso.FolderExists(%1) Then
>>%UNZIP_VBS% ECHO fso.CreateFolder(%1)
>>%UNZIP_VBS% ECHO End If
>>%UNZIP_VBS% ECHO set objShell = CreateObject("Shell.Application")
>>%UNZIP_VBS% ECHO set FilesInZip=objShell.NameSpace(%2).items
>>%UNZIP_VBS% ECHO objShell.NameSpace(%1).CopyHere FilesInZip , 1044
>>%UNZIP_VBS% ECHO Set fso = Nothing
>>%UNZIP_VBS% ECHO Set objShell = Nothing
cscript //nologo %UNZIP_VBS%
IF EXIST %UNZIP_VBS% DEL /f /q %UNZIP_VBS%
SET UNZIP_VBS=
GOTO :EOF

:WGET <URL> <NEWFILE>
::Define SYS_PROXY if does not exist
SET WGET_VBS="%VBS_SAFE_PATH%\wget.vbs"
IF EXIST %WGET_VBS% DEL /f /q %WGET_VBS%
>%WGET_VBS%  ECHO Url = %1
>>%WGET_VBS% ECHO dim xHttp: Set xHttp = createobject("MSXML2.ServerXMLHTTP.6.0")
>>%WGET_VBS% ECHO dim bStrm: Set bStrm = createobject("Adodb.Stream")
>>%WGET_VBS% ECHO xHttp.Open "GET", Url, False
IF NOT "%SYS_PROXY%"=="DIRECT" (
>>%WGET_VBS% ECHO	xHttp.setProxy 2, "%SYS_PROXY%"
)
>>%WGET_VBS% ECHO xHttp.Send
>>%WGET_VBS% ECHO with bStrm
>>%WGET_VBS% ECHO     .type = 1 '//binary
>>%WGET_VBS% ECHO     .open
>>%WGET_VBS% ECHO     .write xHttp.responseBody
>>%WGET_VBS% ECHO     .savetofile %2, 2 '//overwrite
>>%WGET_VBS% ECHO end with
cscript //nologo %WGET_VBS%
IF EXIST %WGET_VBS% DEL /f /q %WGET_VBS%
SET WGET_VBS=
GOTO :EOF

:REFREASH_PATH
SET REFRESH_VBS="REFRESH_PATH.vbs"
IF EXIST %REFRESH_VBS% DEL /f /q %REFRESH_VBS%
IF EXIST REFRESH_PATH.INI DEL /f /q REFRESH_PATH.INI
>%REFRESH_VBS% ECHO Set oShell = WScript.CreateObject("WScript.Shell")
>>%REFRESH_VBS% ECHO filename = oShell.ExpandEnvironmentStrings("REFRESH_PATH.INI")
>>%REFRESH_VBS% ECHO Set objFileSystem = CreateObject("Scripting.fileSystemObject")
>>%REFRESH_VBS% ECHO Set oFile = objFileSystem.CreateTextFile(filename, TRUE)
>>%REFRESH_VBS% ECHO:
>>%REFRESH_VBS% ECHO set oEnv=oShell.Environment("System")
>>%REFRESH_VBS% ECHO path = oEnv("PATH")
>>%REFRESH_VBS% ECHO set oEnv=oShell.Environment("User")
>>%REFRESH_VBS% ECHO path = path ^& ";" ^& oEnv("PATH")
>>%REFRESH_VBS% ECHO oFile.WriteLine("PATH=" ^& path)
>>%REFRESH_VBS% ECHO oFile.Close
cscript //nologo %REFRESH_VBS%
for /F "delims== tokens=1,2" %%A  IN (REFRESH_PATH.INI) DO SET %%A=%%B
IF EXIST %REFRESH_VBS% DEL /f /q %REFRESH_VBS%
IF EXIST REFRESH_PATH.INI DEL /f /q REFRESH_PATH.INI
SET REFRESH_VBS=
GOTO :EOF


:CREATE_START_NODERED_CMD
IF NOT EXIST %NR_SERVICE_DIR% mkdir %NR_SERVICE_DIR%
SET NODERED_CMD="%NR_SERVICE_DIR%\%NR_SERVICE_NAME%_START.cmd"
IF EXIST %NODERED_CMD% DEL /f /q %NODERED_CMD%
>%NODERED_CMD%  ECHO call node-red
GOTO :EOF 

:INSTALL_NSS_NODERED_SERVICE
SET NSSM_URL=https://nssm.cc/release/nssm-2.24.zip
SET NSSM_ZIP_DIR=nssm-2.24
IF NOT EXIST %NR_SERVICE_DIR% mkdir %NR_SERVICE_DIR%
:: DOWNLOAD NSSM
ECHO DOWNLOADING NSSM PLEASE WAIT
CALL :WGET "%NSSM_URL%" "%temp%\nssm.zip"

IF NOT EXIST "%temp%\nssm.zip" GOTO :EXIT_BAD_NSSM_ZIP
:: REMOVE OLD UNZIP FOLDER IF EXIST
IF EXIST %NR_SERVICE_DIR%\%NSSM_ZIP_DIR% rmdir /S /Q %NR_SERVICE_DIR%\%NSSM_ZIP_DIR%
IF EXIST %NR_SERVICE_DIR%\NSSM rmdir /S /Q %NR_SERVICE_DIR%\NSSM
:: UNZIP NSSM
ECHO UNZIPPING NSS PLEASE WAIT 
CALL :UnZipFile "%NR_SERVICE_DIR%" "%temp%\nssm.zip"
DEL /Q "%temp%\nssm.zip"
IF NOT EXIST "%NR_SERVICE_DIR%\%NSSM_ZIP_DIR%" GOTO :EXIT_BAD_NSSM_ZIP
RENAME "%NR_SERVICE_DIR%\%NSSM_ZIP_DIR%" "NSSM"

IF "!PROCESSOR_ARCHITECTURE!"=="x86" (
	SET "NSSM_CMD=NSSM\win32\nssm.exe"
)
IF NOT "!PROCESSOR_ARCHITECTURE!"=="x86" (
	SET "NSSM_CMD=NSSM\win64\nssm.exe"
)

::Create Batch for INSTALL service
SET NODERED_CMD="%NR_SERVICE_DIR%\%NR_SERVICE_NAME%_INSTALL_SERVICE.cmd"
IF EXIST %NODERED_CMD% DEL /f /q %NODERED_CMD%
>%NODERED_CMD%  ECHO %%~dp0\%NSSM_CMD% install %NR_SERVICE_NAME% "%NR_SERVICE_DIR%\%NR_SERVICE_NAME%_START.cmd"
>>%NODERED_CMD% ECHO %%~dp0\%NSSM_CMD% set %NR_SERVICE_NAME% AppDirectory "%NR_SERVICE_DIR%"
>>%NODERED_CMD% ECHO %%~dp0\%NSSM_CMD% set %NR_SERVICE_NAME% Description "A wiring tool for the Internet of Things"
>>%NODERED_CMD% ECHO net start "%NR_SERVICE_NAME%"

::Create Batch for REMOVE service
SET NODERED_CMD="%NR_SERVICE_DIR%\%NR_SERVICE_NAME%_REMOVE_SERVICE.cmd"
IF EXIST %NODERED_CMD% DEL /f /q %NODERED_CMD%
>%NODERED_CMD% ECHO net stop "%NR_SERVICE_NAME%"
>>%NODERED_CMD%  ECHO %%~dp0\%NSSM_CMD% remove %NR_SERVICE_NAME%

SET NODERED_CMD=
SET NSSM_URL=
SET NSSM_ZIP_DIR=
SET NSSM_PRINT=`-- %NR_SERVICE_NAME% Windows Service
GOTO :EOF 

:EXIT_NO_INTERNET
@ECHO OFF
SET INTERNET_RESULT=
DEL /Q %temp%\SYS_PROXY.TMP
CLS
IF "%SYS_PROXY%"=="DIRECT" ECHO Internet connection failure with Connection Direct access (no proxy server)
IF NOT "%SYS_PROXY%"=="DIRECT" ECHO Internet connection failure with Proxy Server: %SYS_PROXY%
PAUSE
GOTO :EXIT_FINAL

:EXIT_BAD_NSSM_ZIP
@ECHO OFF
SET NSSM_URL=
SET NSSM_ZIP_DIR=
CLS
ECHO "NSSM download bad and does not exist"
PAUSE
GOTO :EXIT_FINAL

:EXIT_NODEJS_VER_FILE
@ECHO OFF
CLS
ECHO "NODEJS VER FILE download bad and does not exist"
PAUSE
GOTO :EXIT_FINAL

:EXIT_PYTHON_VER_FILE
@ECHO OFF
CLS
ECHO "PYTHON VER FILE download bad and does not exist"
PAUSE
GOTO :EXIT_FINAL

:EXIT_FINAL

:EOF