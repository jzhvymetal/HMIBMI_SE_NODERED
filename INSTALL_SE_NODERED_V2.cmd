:: ***NODE INSTALLS*** %APPDATA%\npm
:: ***NODERED_PRG  *** C:\%HOMEPATH%\.node-red

::REMOVE INSTALL
:: rmdir /S /Q %APPDATA%\npm & rmdir /S /Q C:\%HOMEPATH%\.node-red && rmdir /S /Q c:\NR_SERVICE

SET SYS_PROXY=DIRECT
SET NODEJS_VER=-1
SET NR_SERVICE_NAME=Node-RED
SET NR_SERVICE_DIR=C:\NR_SERVICE
::SE DOC REQUEST 2.1.4 but still works with the latest 3.x
SET NODERED_VER=-1

@ECHO OFF
SETLOCAL EnableDelayedExpansion
CLS

CALL :INSTALL_NODEJS_MSI
CALL refresh.cmd
CALL :INSTALL_NODE_RED
CALL :CREATE_START_NODERED_CMD
CALL :INSTALL_NSS_NODERED_SERVICE
CALL :NODERED_RUN_WAIT
CALL :NODERED_KILL_WINDOW
CALL :INSTALL_se-node-red-palette_manager
CALL :NODERED_RUN_WAIT
start microsoft-edge:http://127.0.0.1:1880
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
SET FILENAME=se-node-red-palette_manager
IF EXIST "%temp%\%FILENAME%" rmdir /S /Q "%temp%\%FILENAME%"
CALL :UnZipFile "%temp%\%FILENAME%" "%~dp0%FILENAME%.zip"
CALL "%temp%\%FILENAME%\se-node-red-palette_manager_offline_install.bat"
IF EXIST "%temp%\%FILENAME%" rmdir /S /Q "%temp%\%FILENAME%"
GOTO :EOF

:INSTALL_NODE_RED
SET SUB_VER=
IF NOT %NODERED_VER%==-1 SET SUB_VER=@!NODERED_VER!
call npm install -g --unsafe-perm node-red%SUB_VER%
SET SUB_VER=
GOTO :EOF


:INSTALL_NODEJS_MSI
CALL :GET_LATEST_NODEJS_VER

IF "!PROCESSOR_ARCHITECTURE!"=="x86" SET NODEJS_MSI_FILENAME=node-v%NODEJS_VER%-x86.msi
IF NOT "!PROCESSOR_ARCHITECTURE!"=="x86" SET NODEJS_MSI_FILENAME=node-v%NODEJS_VER%-x64.msi
CALL :WGET "https://nodejs.org/dist/v%NODEJS_VER%/%NODEJS_MSI_FILENAME%" "%NODEJS_MSI_FILENAME%" 
MsiExec.exe /i "%NODEJS_MSI_FILENAME%" /passive
ECHO %NODEJS_MSI_FILENAME%
SET NODEJS_MSI_FILENAME=

GOTO :EOF


:GET_LATEST_NODEJS_VER
::LATEST VERSION
IF %NODEJS_VER%==-1 CALL :WGET "http://nodejs.org/dist/latest/SHASUMS256.txt" "%temp%\SHASUMS256.tmp"

IF NOT %NODEJS_VER%==-1 (
	echo.%NODEJS_VER%|findstr /C:"." >nul 2>&1
	if not errorlevel 1 (
	   :: EXACT VERSION
	   GOTO :EOF
	) else (
	   :: LATEST MAIN VERSION
	   CALL :WGET "http://nodejs.org/dist/latest-v%NODEJS_VER%.x/SHASUMS256.txt" "%temp%\SHASUMS256.tmp"
	)
)
IF NOT EXIST %temp%\SHASUMS256.tmp GOTO :EXIT_NODEJS_VER_FILE
findstr /R "node-v.*.pkg" %temp%\SHASUMS256.tmp > %temp%\NODELINE.tmp
SET /p NODELINE=<%temp%\NODELINE.tmp
SET "Up2Sub=%NODELINE:*node-v=%"
SET "NODEJS_VER=%Up2Sub:.pkg="&:"%"
::CLEAR ALL UNUSED VARS
SET NODELINE=
SET Up2Sub=
::REMOVE UNUSED FILES
DEL /Q %temp%\NODELINE.tmp
DEL /Q %temp%\SHASUMS256.tmp
GOTO :EOF

:UnZipFile <ExtractTo> <newzipfile>
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

:EXIT_BAD_NSSM_ZIP
@ECHO OFF
SET NSSM_URL=
SET NSSM_ZIP_DIR=
CLS
ECHO "NSSM download bad and does not exist"
PAUSE
GOTO :EXIT_FINAL

:EXIT_FINAL

:EOF