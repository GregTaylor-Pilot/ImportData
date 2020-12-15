Imports System.IO
Imports System.Xml
Imports System.Configuration
Imports System.Data
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Text.RegularExpressions
Imports System.Security.Permissions
Imports Microsoft.Win32


Module mdlMain

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Private IsLogging As Boolean = False
    Private StoreNumber, OfficeDB As String
    Private iStoreNumber As Integer
    Private WinVersionInfo As OperatingSystem = Environment.OSVersion
    Private AppPath As String = My.Application.Info.DirectoryPath
    Private ScriptLog As String = Path.Combine(AppPath, My.Application.Info.AssemblyName & ".log")
    Private ErrorLog As String = Path.Combine(AppPath, My.Application.Info.AssemblyName & "_err.log")
    Private mMailSender As String = "DoNotReply@pilottravelcenters.com"
    Private mRelayServer As String
    Private DoExternalEvent As Boolean = True
    Private DoRPOBuild As Boolean = True
    Private ShowProgress As Boolean = False
    Private HasDatFiles As Boolean = False
    Private ImportMustRun As Boolean = False
    Private LogDebugMessages As Boolean = False
    Private DoMaskPrompts As Boolean = True
    Private ASNDataPath, ASNFileSpec, ShipFile, ShipBak As String
    Private mFirstAlert, mSecondAlert, mExternalEventTimer, mLogLevel, mPriceBatchDays, mLogFileMaxSize, mDelayAfterImport As Integer
    Private mRPOBuildFilesMinHour, mRPOBuildFilesMaxHour As Integer
    Private mPendingPriceBatchesCheckTimes As New List(Of Integer)
    Private mSendPendingPriceBatchEmails As Boolean = True
    Private mPendingPostSalesContinueIfTrue As Boolean = False
    Private mPendingPostSalesForceSetToTrue As Boolean = False
    Private mAutoPriceSrvOnEachRun As Boolean = False
    Private mIncludeNullPriceModHdrDates As Boolean = False
    Private mAutoPriceSrvOnPriceModHdrRecordsReturned As Boolean = False
    Private mKillExternalEventRunnerAfterMaxRetries As Boolean = False
    Private mIsR10Site As Boolean = False
    Private mDoAutoPostSales As Boolean = True
    Private mDoAutoPriceSrv As Boolean = True
    Private mDoFixItWithoutDatFiles As Boolean = True
    Private mPriceModHdrCriteria As String = "=P"
    Private mPriceModHdrDateField As String
    Private mPendingPriceBatchEmailDelayInHours As Integer
    Private mPostSalesSleep, mPostSalesRetry, mDbUpgraderDelay, mDbUpgraderRetry, mExternalEventDelay, mExternalEventRetry As Integer
    Private mRcvDatFilesAgeWarningInDays, mImportExeMaxRetryAttempts As Integer
    Private FirebirdDriverVersion As String = "1.0"
    Private MyVersion As String = Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString
    Private dbDriver As String = ""
    Private dbConnectionString As String = "DRIVER={0};UID={1};PWD={2};DBNAME={3};"
    'Private MyProtection As SecureIt = New SecureIt("PilotFlyingJ2018", MyEncryptionType:=SecureIt.MyEncryptionTypes.SHA1, Iterations:=2)
    Private MyProtection As ASM.MyFunctions = New ASM.MyFunctions("PilotFlyingJ2018", MyEncryptionType:=ASM.MyFunctions.MyEncryptionTypes.SHA1, Iterations:=2)
    Private SectionVal As String = "Production"
    Private Header As String
    Private ImportExe As String = "C:\Office\exe\Import.exe"
    Private Salt As String = "P!l0+fLy!Nj01"

    Private mLogger As LogWriter
    Private mMailRecipient, mOFBMailRecipient, mPricebatchMailRecipient, mRcvDatAgeWarningMailRecipient As String
    Private mMyUser, mMyPass As String

    Public myProcessId As Integer = Process.GetCurrentProcess().Id

    'Const SEPARATOR As String = " --------------------------------------------------"

    'Const FIX_CROSS_SITE_PROCEDURE As String = "execute procedure fix_cross_site;"
    'Const PILOT_MASKING_PROCEDURE As String = "execute procedure pilot_masking;"

    'Const FIXIT_PROCEDURE As String = "execute procedure fixit;"
    Const RCV_FILE_PATH As String = "C:\Office\Rcv"
    Const IMPORT_LOCKED_UP As String = "Import.exe is locked up at this location.  Please stop maintenance (Pos Transmitter), kill import.exe and importdata.exe, then restart maintenance and re-launch importdata.exe."
    Const IMPORT_X_TRIES As String = "Import.exe has run {0} times and still has DAT files to be processed in the RCV directory. Manual intervention may be required. Please check the site."

    Public Sub Main()

        Dim HasShippers As Boolean = False
        Dim HasPriceChangeFiles As Boolean = False
        Dim HasCategoryFile As Boolean

        Dim myConfig As String = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile
        Dim myConfigKey As String = myConfig & ".key"
        Dim isValidHash As Boolean = False

        'StoreNumber = GetStoreNumber()
        StoreNumber = GetStoreNumberFromMachineName(Environment.MachineName)

        Try
            FirebirdDriverVersion = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\Firebird/InterBase(r) driver", "DriverODBCVer", "1.0").ToString()
            dbDriver = "Firebird/InterBase(r) driver"
        Catch ex As Exception
            Try
                FirebirdDriverVersion = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\InterBase ODBC driver", "DriverODBCVer", "1.0").ToString()
                dbDriver = "InterBase ODBC driver"
            Catch ex1 As Exception
                FirebirdDriverVersion = "1.0"
            End Try
        End Try

        If IsLogging = False Then
            mLogger = New LogWriter(ScriptLog, funcNum:=-1)
            IsLogging = True
            Log(String.Format("Started ImportData version {0}", MyVersion), 1, 1)
        End If

        If Not Integer.TryParse(StoreNumber, iStoreNumber) Then
            iStoreNumber = 0
        End If
        Log("Loading configuration", 1, 1)
        GetConfigOptions()

        InitiateOpenLogMessages()

        Try
            If File.Exists(myConfigKey) Then
                Log(String.Format("{0} key file was found. Loading current hash.", myConfigKey), 5, 1)
                Dim currentHash As String = MyProtection.GetValue(myConfigKey, "/importdata/@hash")
                Log(String.Format("Current hash is {0}", currentHash), 10, 1)
                isValidHash = MyProtection.VerifySha512FileHash(myConfig, 32, currentHash, Salt)
                Log(String.Format("IsHasValidHash =  {0}", isValidHash.ToString), 10, 1)
            Else
                Log("Error - Key file is missing. Unable to continue with ImportData routines. Exiting.", 1, 9001)
                mMailSender = "DoNotReply@pilottravelcenters.com"
                DbgMsg("Key file is missing")
                EmailError(mOFBMailRecipient, "Missing key file", String.Format("The {0} key file is missing. Please generate a key file in order to launch ImportData",
                                                                            myConfigKey))
                If IsLogging Then
                    mLogger.Close()
                End If
                Exit Sub
            End If
        Catch ex As Exception
            Log("Error - Unable to check config hash (check to verify the ASM.dll is present and current version.", 1, 9001)
            Log("Unable to continue with ImportData routines. Exiting.", 1, 9001)

            mLogger.Close()
            Exit Sub
        End Try

        If Not isValidHash Then
            mMailSender = "DoNotReply@pilottravelcenters.com"
            DbgMsg("Key file did not contain the correct hash")
            Log(String.Format("The {0} key file did not contain the correct hash",
                                                                        myConfigKey), 1, 9001)
            Log("Unable to continue with ImportData routines. Exiting.", 1, 9001)

            EmailError(mOFBMailRecipient, "Incorrect hash in key file", String.Format("The {0} key file did not contain the correct hash",
                                                                        myConfigKey))
            mLogger.Close()
            Exit Sub
        End If


        Dim Progress As New ProgressDialog()

        'MaskPrompts()




        'Dim CP As New ConnectionProtection(Application.ExecutablePath)

        'CP.EncryptFileAppData()
        'CP.EncryptFileConnectionString()
        'CP = Nothing

        'Dim Conn As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("EncryptConnStringSection.My.MySettings.MyConn")

        'Dim Args As New ArrayList(Environment.GetCommandLineArgs())
        'Dim CurrentCommand As String

        'Args.RemoveAt(0)
        'If Args.Count > 0 Then
        '    For Each arg As String In Args
        '        CurrentCommand = arg.ToString.ToLower
        '        If (((CurrentCommand.Chars(0) = "-"c) OrElse (CurrentCommand.Chars(0) = "/"c)) AndAlso (CurrentCommand.Length > 1)) Then
        '            CurrentCommand = CurrentCommand.Substring(1)
        '        End If
        '        Select Case CurrentCommand
        '            Case "init" 'Expected to use Init flag to encrypt the config file when deployed.
        '                Application.Exit()
        '                End

        '            Case Else


        '        End Select
        '    Next
        'End If
        'MessageBox.Show(Conn.ConnectionString)



        'Dim myVar As String = InputBox("Enter the value", "Encrypt", "")

        'Dim encryptedVar As String = MyProtection.AESEncrypt(myVar, "ImportData", "PFJ")
        'MessageBox.Show(encryptedVar)
        'My.Computer.Clipboard.Clear()
        'My.Computer.Clipboard.SetText(encryptedVar)

        'Dim decryptedVar As String = MyProtection.AESDecrypt(encryptedVar, "ImportData", "PFJ")
        'MessageBox.Show(decryptedVar)

        'Application.Exit()

        'MessageBox.Show(SecureIt.Decrypt(mMyUser, "ImportData"))
        'MessageBox.Show(SecureIt.Decrypt(mMyPass, "ImportData"))



        Header &= "Machine: " & Environment.MachineName & vbCrLf
        Header &= "ImportData Version: " & MyVersion & vbCrLf
        Header &= "Database Driver Version: " & FirebirdDriverVersion & vbCrLf
        Header &= "Windows Version Info: " & getOSInfo() & vbCrLf
        Header &= "Mode: " & SectionVal & vbCrLf
        Header &= "Message: "


        'MessageBox.Show("Firebird Driver version: " & FirebirdDriverVersion)

        If CDbl(FirebirdDriverVersion) < 3.51 Then
            DbgMsg("Driver is not current version")
            Log("Firebird driver is not the current version. Must be 3.51 or greater. Exiting.", 1, 9001)
            EmailError(mOFBMailRecipient, "Driver Issue", String.Format("Firebird/Interbase ODBC driver version {0} found at this location (3.51 or greater required)", FirebirdDriverVersion))
            mLogger.Close()
            IsLogging = False
            Exit Sub
        Else
            Log("Firebird driver version is >= minimum driver version of 3.51. Processing will continue.", 1, 1)
        End If


        If Not ImportDataRunning() Then
            DbgMsg("ImportData was not running")
            Log("ImportData was not preivously running. Continuing.", 5, 1)
            If Not MaintSrvRunning(False) Then
                Log("MaintSrv was not running, continuing ImportData main routines", 5, 1)
                DbgMsg("MaintSrv was not running")
                ProgressDialog.AddProgress("ImportData started")

                DbgMsg("Reading command line switches")
                Log("Calling ReadCommandLineSwitches()", 5, 1)
                ReadCommandLineSwitches()
                Try
                    Log("Creating LaunchTime.txt", 5, 1)
                    Dim LaunchWriter As New StreamWriter(New FileStream(Path.Combine(AppPath, "LaunchTime.txt"), FileMode.Create, FileAccess.Write))
                    LaunchWriter.WriteLine(Date.Now.Ticks)
                    LaunchWriter.Close()
                Catch ex As Exception
                    Log("Error creating LaunchTime.txt", 1, 9001)
                End Try

                DbgMsg("Setting paths")
                Log("Calling SetPaths()", 5, 1)
                SetPaths()
                DbgMsg("Rolling large logs")
                Log("Calling RollLargeLog()", 5, 1)
                RollLargeLog()

                DbgMsg("Checking for *.dat files")
                Log("Checking if HasDatFiles", 5, 1)
                HasDatFiles = Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length > 0
                Log(String.Format("{0} *.dat file(s) found in {1}", Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length, RCV_FILE_PATH), 3, 1)
                Log(String.Format("Setting HasDatFiles to {0}", HasDatFiles), 3, 1)

                DbgMsg("Checking for cr*.dat files")
                Log("Checking if HasPriceChangeFiles", 5, 1)
                HasPriceChangeFiles = Directory.GetFiles(RCV_FILE_PATH, "CR*.dat").Length > 0
                Log(String.Format("{0} cr*.dat file(s) found in {1}", Directory.GetFiles(RCV_FILE_PATH, "cr*.dat").Length, RCV_FILE_PATH), 3, 1)
                Log(String.Format("Setting HasPriceChangeFiles to {0}", HasPriceChangeFiles), 3, 1)

                DbgMsg("Checking for caty*.dat files")
                Log("Checking if HasCategoryFiles", 5, 1)
                HasCategoryFile = Directory.GetFiles(RCV_FILE_PATH, "caty*.dat").Length > 0 OrElse Directory.GetFiles(RCV_FILE_PATH, "category.dat").Length > 0
                Log(String.Format("{0} caty*.dat file(s) found in {1}", Directory.GetFiles(RCV_FILE_PATH, "caty*.dat").Length, RCV_FILE_PATH), 3, 1)
                Log(String.Format("{0} category.dat file(s) found in {1}", Directory.GetFiles(RCV_FILE_PATH, "category.dat").Length, RCV_FILE_PATH), 3, 1)
                Log(String.Format("Setting HasCategoryFile to {0}", HasCategoryFile), 3, 1)

                DbgMsg("Checking for *.* files")
                Log("Checking if ImportMustRun", 5, 1)
                ImportMustRun = Directory.GetFiles(RCV_FILE_PATH, "*.*").Length > 0 OrElse Directory.GetFiles(Path.Combine(RCV_FILE_PATH, "RcvTbl"), "*.*").Length > 0
                Log(String.Format("{0} *.* file(s) found in {1}", Directory.GetFiles(RCV_FILE_PATH, "*.*").Length, RCV_FILE_PATH), 3, 1)
                Log(String.Format("{0} *.* file(s) found in {1}", Directory.GetFiles(Path.Combine(RCV_FILE_PATH, "RcvTbl"), "*.*").Length,
                                          Path.Combine(RCV_FILE_PATH, "RcvTbl")), 3, 1)
                Log(String.Format("Setting ImportMustRun to {0}", ImportMustRun), 3, 1)

                DbgMsg("Setting database path")
                Log("Calling SetDatabasePath()", 5, 1)
                If Not SetDatabasePath() Then
                    Log("Calling SetDatabasePath()", 5, 9001)
                    DbgMsg("Database path not set. Exiting.")
                    ProgressDialog.AddProgress("ImportData ended")
                    Log("ImportData End. Script ended", 1, 9001)
                    mLogger.Close()
                    IsLogging = False
                    ProgressDialog.AddProgress("... done")
                    ProgressDialog.Done()
                    Exit Sub
                Else
                    Log("Database path was set. Database connection string will be set", 5, 1)
                    Dim dbEncryptedConnectionString As String = String.Format(dbConnectionString, dbDriver, mMyUser,
                                                       mMyPass, OfficeDB) 'before dbconnection string is unencrypted
                    dbConnectionString = String.Format(dbConnectionString, dbDriver, MyProtection.AESDecrypt(mMyUser, "ImportData", "PFJ"),
                                                       MyProtection.AESDecrypt(mMyPass, "ImportData", "PFJ"), OfficeDB)
                    Log(String.Format("Database connection string (encrypted user/pass): {0}", dbEncryptedConnectionString), 10, 1)
                End If

                'MessageBox.Show(dbConnectionString)
                'Log("    *** DB Connection String = " & dbConnectionString, 5)

                DbgMsg("Checking External event runner")
                If DoExternalEvent Then
                    Log("DoExternalEvent was set to true. Checking to ensure ExternalEventRunner is running", 5, 1)
                    If Not File.Exists("C:\office\exe\externaleventrunner.exe") Then
                        DbgMsg("External event runner not found")
                        ProgressDialog.AddProgress("ExternalEventRunner.exe not found")
                        ProgressDialog.AddProgress("... exiting")
                        ProgressDialog.Done()
                        Log("ExternalEventRunner.exe was not found", 2, 9001)
                        mLogger.Close()
                        IsLogging = False
                        Exit Sub
                    Else
                        Log("Externaleventrunner.exe was found. Continuing script", 5, 1)
                    End If
                Else
                    Log("DoExternalEvent was set to false", 5, 1)
                End If

                Try
                    DbgMsg("Checking for shippers")
                    Log("Checking for shippers", 10, 1)
                    HasShippers = CheckForShippers()
                    Log(String.Format("Checking for shippers has returned {0}", HasShippers), 10, 1)

                    DbgMsg("Checking for Pending Post Sales")
                    Log("Checking for pending post sales", 10, 1)
                    If PendingPostSales() Then
                        Log("Pending post sales returned true", 1, 9001)
                        If mPendingPostSalesContinueIfTrue = False Then
                            ProgressDialog.AddProgress("ImportData ended")
                            Log("ImportData End. Script ended", 1, 9001)
                            mLogger.Close()
                            IsLogging = False
                            ProgressDialog.AddProgress("... done")
                            ProgressDialog.Done()
                            Exit Sub
                        Else
                            Log("Pending post sales continue if true was set. Continuing with the script.", 5, 1)
                        End If
                    Else
                        Log("Checking for pending post sales has returned false", 10, 1)
                    End If

                    DbgMsg("Checking if Import is still running")
                    Log("Checking if Import.exe is running", 10, 1)
                    If ProcessRunning("Import") Then
                        Log("Import.exe appears to be running. Initiating recheck function ImportStillRunning()", 5, 1)
                        If ImportStillRunning() Then
                            Log("Import is still running. Exiting", 1, 9001)
                            DbgMsg("Import is still running. Exiting.")
                            ProgressDialog.AddProgress("ImportData ended")
                            Log("ImportData End. Script ended", 1, 1)
                            mLogger.Close()
                            IsLogging = False
                            ProgressDialog.AddProgress("... done")
                            ProgressDialog.Done()
                            Exit Sub
                        Else
                            Log("Import.exe appears to have stopped running. Launching Import", 5, 1)
                            DbgMsg("Launching import")
                            LaunchImport()
                        End If
                    Else
                        Log("Import.exe was not running. Launching Import", 5, 1)
                        DbgMsg("Launching import")
                        LaunchImport()
                    End If

                    'Call Update Item Status and Fixit
                    If HasDatFiles Then
                        Log("Performing prerequesite fuctions required for when dat files are present.", 3, 1)
                        DbgMsg("SRV has dat files. Running UpdateItemStatus")
                        UpdateItemStatus()
                        DbgMsg("Running AutoPostSales")
                        AutoPostSales()
                        DbgMsg("Running AutoPriceSrv")
                        AutoPriceSrv()
                        Log("Prerequisite functions for dat files completed.", 3, 1)
                    Else
                        If mDoFixItWithoutDatFiles Then
                            FixIt()
                        End If
                        If mAutoPriceSrvOnEachRun Then
                            Log("No dat files exist. UpdateItemStatus() and AutoPostSales() will not be called", 5, 1)
                            Log("AutoPriceSrvOnEachRun is set to true. Running AutoPriceSrv()", 5, 1)
                            AutoPriceSrv()
                        Else
                            Log("No dat files exist. UpdateItemStatus(), AutoPostSales() and AutoPriceSrv() will not be called", 5, 1)
                        End If
                    End If

                    DbgMsg("Processing remaining DAT files")
                    Log("Processing remaining dat files", 5, 1)
                    ProcessDatFiles()

                    'Check if there are still any pending pricebatches
                    DbgMsg("Checking for pending price batches")
                    Log("Checking for pending price batches", 5, 1)
                    Dim pbResults As Integer = 0
                    pbResults = CheckForPendingPricebatches(mSendPendingPriceBatchEmails)

                    If mAutoPriceSrvOnPriceModHdrRecordsReturned Then
                        If pbResults > 0 Then
                            Log(String.Format("AutoPriceSrvOnPriceModHdr is set to true and there were {0} pending price batches returned.", pbResults), 5, 1)
                            AutoPriceSrv()
                        End If
                    End If

                    'Execute fix_cross_site procedure if we had a category file
                    If HasCategoryFile Then
                        DbgMsg("Call Cross Site Fix")
                        Log("HasCategoryFiles is true. CallCrossSiteFix() will be called", 5, 1)
                        CallCrossSiteFix()
                    Else
                        Log("HasCategoryFiles is false. CallCrossSiteFix() will not be called", 5, 1)
                    End If

                    DbgMsg("Process ASN files")
                    Log("Processing ASN files", 5, 1)
                    ProcessASNFiles()

                    'checks if store has a shipper file produced by HOST
                    If HasShippers Then
                        DbgMsg("Process shippers")
                        Log("HasShippers is true. ProcessShippers() will be called", 5, 1)
                        ProcessShippers()
                    Else
                        Log("HasShippers is false. ProcessShippers() will not be called", 5, 1)
                    End If

                    DbgMsg("Call pilot masking")
                    Log("Calling CallPilotMasking()", 5, 1)
                    CallPilotMasking()

                    'Notify home office of price changes
                    If HasPriceChangeFiles Then
                        DbgMsg("Has price change files.")
                        Log("Price change files detected.... sending notification to price change service", 2, 1)
                        Try
                            Dim ws As New MessengerWebService.MessengerWebService()
                            Dim Result As Boolean = ws.PriceBookDownloadCompletedAlert(iStoreNumber)
                            Log("Web service result: " & Result, 3, 1)
                        Catch ex As Exception
                            Log("Error instantiating MessengerWebService(): " & ex.Message, 1, 1)
                        End Try
                    Else
                        Log("HasPriceChangeFiles is false. Notification will not be sent to price change service.", 5, 1)
                    End If

                    DbgMsg("Prepare RPO Data")
                    Log("Calling RPOPrepareData()", 5, 1)
                    RPOPrepareData()

                    DbgMsg("Checking for failure log")
                    Log("Calling CheckForFailureLog()", 5, 1)
                    CheckForFailureLog()

                    If Not MaintSrvRunning(True) Then
                        DbgMsg("MaintSrv not running, masking prompts")
                        Log("MaintSrv is not running. Calling MaskPrompts()", 5, 1)
                        MaskPrompts()
                    Else
                        Log("MaintSrv is running. MaskPrompts() will not be called", 5, 1)
                    End If

                    DbgMsg("Finishing processing")
                    ProgressDialog.AddProgress("ImportData ended")
                    Log("ImportData End. Script ended", 1, 1)
                    mLogger.Close()
                    IsLogging = False
                    ProgressDialog.AddProgress("... done")
                    ProgressDialog.Done()

                Catch ex As Exception
                    DbgMsg("Main routine exception")
                    EmailError(mMailRecipient, "Main Routine Exception: ", ex.Message)
                    Log(String.Format("Main routine encountered an exception. Exception: {0}", ex.Message), 1, 9001)
                    Log("E-Mail sent (Main Routine)", 1, 9001)

                    Log("ImportData ended", 1, 1)
                    mLogger.Close()
                    IsLogging = False
                End Try
            Else
                DbgMsg("MaintSrv was running, unable to launch ImportData")
                If IsLogging = False Then
                    mLogger = New LogWriter(ScriptLog, funcNum:=-1)
                    IsLogging = True
                    Log("ImportData is started", 1, 1)
                    InitiateOpenLogMessages()
                End If

                Log("Script ImportData launch attempted, but MaintSrv appears to be running", 1, 9001)
                Log("(No task finished line detected at end of MaintSrv log)", 1, 9001)
                Log("ImportData will now exit", 1, 1)
                mLogger.Close()
                IsLogging = False
            End If
        End If
    End Sub

    Private Sub InitiateOpenLogMessages()
        Log(String.Format("Store = {0}", StoreNumber), 1, 101)
        Log(String.Format("Version = {0}", MyVersion), 1, 101)
        Log(String.Format("Driver = {0}; version = {1}", dbDriver, FirebirdDriverVersion), 10, 101)
        Log(String.Format("Mode = {0}", SectionVal), 5, 101)
        Log(String.Format("OS = {0}", getOSInfo), 5, 101)
    End Sub

    Private Sub Log(ByVal Message As String, ByVal LogLevel As Integer, ByVal fNum As Integer)
        Try
            If Not mLogger Is Nothing Then
                mLogger.Log(Message, LogLevel, fNum)
            Else
                Try
                    mLogger = New LogWriter(ScriptLog, funcNum:=-1)
                    IsLogging = True
                    mLogger.Log(Message, LogLevel, fNum)
                Catch ex1 As Exception
                    LogError(Message, fNum, ex1.Message)
                End Try

            End If

        Catch ex As Exception
            LogError(Message, fNum, ex.Message)
        End Try
    End Sub

    Private Sub LogError(ByVal Message As String, ByVal fNum As Integer, ByVal ErrMsg As String)
        If fNum < 1000 Then fNum = fNum + 9000
        Try
            Dim mLoggerErr As LogWriter = New LogWriter(ErrorLog, funcNum:=-1)
            mLoggerErr.Log(String.Format("Unable to log message: ({0}){1}", fNum, Message), 1, fNum)
            mLoggerErr.Log(String.Format("Error: {0}", ErrMsg), 1, fNum)
            mLoggerErr.Close()
        Catch ex As Exception

        End Try

    End Sub

    Private Function ValidateEmailByDefault(ByVal EmailAddr As String, ByVal defaultEmail As String) As String
        Dim regex As Regex = New Regex("^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$")
        Dim isValid As Boolean = regex.IsMatch(EmailAddr)
        If Not isValid Then
            Return defaultEmail
        Else
            Return EmailAddr
        End If
    End Function

    Private Function ValidateEmailBool(ByVal EmailAddr As String) As Boolean
        Dim regex As Regex = New Regex("^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$")
        Dim isValid As Boolean = regex.IsMatch(EmailAddr)
        Return isValid
    End Function



    Private Sub DbgMsg(ByVal Text As String, Optional ByVal Title As String = "Debug",
                                Optional ByVal Button As MessageBoxButtons = MessageBoxButtons.OK,
                                Optional ByVal Icon As MessageBoxIcon = MessageBoxIcon.None)
        If LogDebugMessages Then
            MessageBox.Show(Text, Title, Button, Icon)
        End If
    End Sub

    Private Function getOSInfo() As String
        Dim os As OperatingSystem = Environment.OSVersion
        Dim vs As Version = os.Version
        Dim operatingSystem As String = ""

        If os.Platform = PlatformID.Win32Windows Then

            Select Case vs.Minor
                Case 0
                    operatingSystem = "95"
                Case 10

                    If vs.Revision.ToString() = "2222A" Then
                        operatingSystem = "98SE"
                    Else
                        operatingSystem = "98"
                    End If

                Case 90
                    operatingSystem = "Me"
                Case Else
            End Select
        ElseIf os.Platform = PlatformID.Win32NT Then

            Select Case vs.Major
                Case 3
                    operatingSystem = "NT 3.51"
                Case 4
                    operatingSystem = "NT 4.0"
                Case 5

                    If vs.Minor = 0 Then
                        operatingSystem = "2000"
                    Else
                        operatingSystem = "XP"
                    End If

                Case 6

                    If vs.Minor = 0 Then
                        operatingSystem = "Vista"
                    ElseIf vs.Minor = 1 Then
                        operatingSystem = "7"
                    ElseIf vs.Minor = 2 Then
                        operatingSystem = "8"
                    Else
                        operatingSystem = "8.1"
                    End If

                Case 10
                    operatingSystem = "10"
                Case Else
            End Select
        End If

        If operatingSystem <> "" Then
            operatingSystem = "Windows " & operatingSystem

            If os.ServicePack <> "" Then
                operatingSystem += " " & os.ServicePack
            End If
        End If

        Return operatingSystem
    End Function

    Private Sub EmailError(ByVal recipients As String, ByVal Subject As String, ByVal body As String)
        Log("Entering EmailError(recipients,Subject,body) Sub", 10, 100)
        Log(String.Format("Recipients: {0}", recipients), 10, 100)
        Log(String.Format("Subject: {0}", Subject), 10, 100)
        Log(String.Format("Body: {0}", body), 15, 100)
        Try
            Dim SendMail As New Net.Mail.SmtpClient(mRelayServer)
            Dim mailMessage As New Net.Mail.MailMessage()
            If mMailSender = "" Then
                mMailSender = "DoNotReply@pilottravelcenters.com"
            End If
            If mMailSender.IndexOf("@") < 0 Then
                mMailSender = mMailSender.Trim & "@pilottravelcenters.com"
            End If
            Log(String.Format("Sender: {0}", mMailSender), 10, 100)
            mailMessage.From = New Net.Mail.MailAddress(mMailSender)
            For Each recipient As String In recipients.Split(";"c)
                If recipient.IndexOf("@") < 0 Then
                    recipient = recipient.Trim() & "@pilottravelcenters.com"
                End If
                If ValidateEmailBool(recipient) Then
                    mailMessage.To.Add(recipient.Trim())
                Else
                    Log(String.Format("{0} is not a valid email recipient. Recipient will not be added to list", recipient.Trim), 1, 9100)
                End If
            Next
            mailMessage.Subject = String.Format("Store{0} - {1}", StoreNumber, Subject)
            mailMessage.Body = Header & vbCrLf & body
            SendMail.EnableSsl = True
            Try
                Log("Sending email", 5, 100)
                SendMail.Send(mailMessage)
                Log("Email sent", 5, 100)
            Catch ex1 As Net.Mail.SmtpException
                If ex1.Message.IndexOf("does not support secure connections") > 0 Then
                    Try
                        Log("Email relay server does not support secure connections. Setting SSL to false", 1, 9100)
                        SendMail.EnableSsl = False
                        Log("Senging email", 5, 100)
                        SendMail.Send(mailMessage)
                        Log("Email sent", 5, 100)
                    Catch ex2 As Exception
                        Log(String.Format("EmailError sub encountered an error. Exception: {0}", ex2.Message), 1, 9100)
                        Log("Exiting EmailError(recipients,Subject,body) Sub", 10, 100)
                        Exit Sub
                    End Try

                Else
                    Log(String.Format("EmailError sub encountered an error. Exception: {0}", ex1.Message), 1, 9100)
                    Log("Exiting EmailError(recipients,Subject,body) Sub", 10, 100)
                    Exit Sub
                End If
            End Try
        Catch ex As Exception
            Log(String.Format("EmailError sub encountered an error. Exception: {0}", ex.Message), 1, 9100)
            Log("Exiting EmailError(recipients,Subject,body) Sub", 10, 100)
            Exit Sub
        End Try
        Log("Exiting EmailError(recipients,Subject,body) Sub", 10, 100)
    End Sub

    Private Function ReadLineSafe(ByVal reader As TextReader, ByVal maxLength As Integer) As String
        Dim sb As Text.StringBuilder = New Text.StringBuilder()

        While True
            Dim ch As Integer = reader.Read()
            If ch = -1 Then Exit While

            If ch = CInt(vbCr) OrElse ch = CInt(vbLf) Then
                If ch = CInt(vbCr) AndAlso reader.Peek() = CInt(vbLf) Then reader.Read()
                Return sb.ToString()
            End If

            sb.Append(ChrW(ch))
            If sb.Length > maxLength Then Throw New InvalidOperationException("Line is too long")
        End While

        If sb.Length > 0 Then Return sb.ToString()
        Return Nothing
    End Function

    Private Function ReadFile(ByRef FileName As String) As String
        If File.Exists(FileName) Then
            Dim NextLine As String
            Dim Text As New Text.StringBuilder(500)
            Try
                Dim Reader As TextReader = New StreamReader(New FileStream(FileName, FileMode.Open, FileAccess.Read))
                While True
                    NextLine = ReadLineSafe(Reader, 2048)
                    If NextLine = Nothing Then
                        Exit While
                    End If
                    NextLine = NextLine & vbCrLf
                    Text.Append(NextLine)
                End While
                Reader.Close()
                Reader = Nothing

                'Dim ReaderS As New StreamReader(New FileStream(FileName, FileMode.Open, FileAccess.Read))
                'While Not ReaderS.EndOfStream
                '    NextLine = ReaderS.ReadLine()
                '    NextLine = NextLine & vbCrLf
                '    Text.Append(NextLine)
                'End While
                'ReaderS.Close()
            Catch iox As IOException

            Catch ex As Exception

            End Try

            Return Text.ToString()
        Else
            Return ""
        End If

    End Function

    Function CountCharOccurencesInStr(ByRef sStringToSearch As String, ByRef sCharacter As Char) As Integer
        Return sStringToSearch.Split(sCharacter).Length - 1
    End Function

    Private Sub RollLargeLog()
        If IsLogging = False Then
            mLogger = New LogWriter(ScriptLog, funcNum:=-1)
            IsLogging = True
        End If
        Log("Entering RollLargeLog() Sub", 10, 2)
        Dim errMsg As String = ""
        If File.Exists(ScriptLog) Then
            Dim length As Long
            Log(String.Format("{0} exists. Checking size.", ScriptLog), 10, 2)
            Try
                Dim info As New FileInfo(ScriptLog)
                length = info.Length
            Catch ex As Exception
                Log(String.Format("Error encountered with FileInfo for log. Exception: ", ex.Message), 5, 9002)
                length = 0
            End Try

            Log(String.Format("{0} is {1} bytes", ScriptLog, length.ToString), 10, 2)
            If length > (mLogFileMaxSize * 1048576) Then
                Log(String.Format("{0} exceeds {1} MB. Attempting to roll log.", ScriptLog, mLogFileMaxSize), 3, 2)
                Try
                    If File.Exists(ScriptLog & "_bak") Then
                        Log(String.Format("Backup exists. Attempting to delete {0}", ScriptLog & "_bak"), 5, 2)
                        File.Delete(ScriptLog & "_bak")
                        Log(String.Format("{0} backup deleted successfully", ScriptLog & "_bak"), 5, 2)
                    End If
                Catch ex As Exception

                End Try
                Log(String.Format("Attempting to move {0} to {1}", ScriptLog, ScriptLog & "_bak"), 5, 2)
                Log("Closing log file before backing up", 1, 2)
                mLogger.Close()
                IsLogging = False

                Try
                    File.Move(ScriptLog, ScriptLog & "_bak")
                    Threading.Thread.Sleep(5000)

                    If IsLogging = False Then
                        mLogger = New LogWriter(ScriptLog, funcNum:=-1)
                        IsLogging = True
                    End If
                    InitiateOpenLogMessages()
                    Log("Already entered RollLargeLog() Sub", 10, 2)
                    Log(String.Format("Rolling log file completed succesfully! {0} was moved to {1}", ScriptLog, ScriptLog & "_bak"), 1, 2)
                Catch ex As Exception
                    errMsg = ex.Message()
                    Log(String.Format("Error encountered rolling log file. Exception: ", errMsg), 5, 9002)
                Finally
                    InitiateOpenLogMessages()
                End Try
            Else
                Log(String.Format("{0} does not exceed {1} MB. File will not be rolled.", ScriptLog, mLogFileMaxSize), 10, 2)
            End If

        Else
            Log(String.Format("{0} does not exist.", ScriptLog), 3, 2)
        End If
        Log("Exiting RollLargeLog() Sub", 10, 2)
    End Sub

    Public Function ProcessRunning(ByRef pName As String) As Boolean
        Log("Entering ProcessRunning(pName) Function (As Boolean)", 10, 3)
        Dim Processes() As Process = Process.GetProcessesByName(pName)
        Log(String.Format("Exiting ProcessRunning(pName) Function ({0})", Processes.Length > 0), 10, 3)
        Return Processes.Length > 0
    End Function

    Private Sub MaskPrompts()
        Log("Entering MaskPrompts() Sub", 10, 4)
        Dim configPath As String = Path.Combine(AppPath, "ImportData.exe.config")
        Dim config As XmlDocument = New XmlDocument()
        Dim template As String = "select pmntcode_id, pmsubcode_id from pmnt where pmnt_name like '{0}';"
        Dim selectTemplate As String = "select * from tillpmnt_prompt where sernum_tillpmnt in " &
        "(select sernum from tillpmnt where pmcode = {0} And pmsubcode = {1}) " &
        "And prompt_name = '{2}' and sernum_tillpmnt > {3};"
        Dim maskSQL As String = "update tillpmnt_prompt set prompt_response = {0} where sernum_tillpmnt in " &
        "(select sernum from tillpmnt where pmcode = {1} and pmsubcode = {2}) " &
        "and prompt_name = '{3}' and sernum_tillpmnt > {4};"
        Dim maxSeq As String = "select max(sernum) from tillpmnt where pmcode = {0} and pmsubcode = {1};"
        Dim querySql As String = Path.Combine(AppPath, "CodeQuery.sql")
        Dim queryFile As String = Path.Combine(AppPath, "QueryResults.txt")
        Dim pmntcode_id, pmsubcode_id As String
        Dim result, maskText As String
        Dim cardname As String
        Dim prompt As String
        Dim minSequence, maxSequence As String
        Dim partialMaskLength As Integer
        Dim writer As StreamWriter

        'mLogger = New LogWriter(ScriptLog)
        'Log("Script ImportData masking started", 1, 4)

        Try
            config.XmlResolver = Nothing
            Log(String.Format("Loading prompts from {0}", configPath), 5, 4)
            config.Load(configPath)
            Log("Config loaded. Parsing nodes", 11, 4)
            For Each node As XmlNode In config.SelectNodes("/configuration/userSettings/ImportData.My.MySettings/setting")
                cardname = node.Attributes("name").InnerText
                Log(String.Format("Loading node for {0}", cardname), 11, 4)
                prompt = node.ChildNodes(0).InnerText
                Log(String.Format("Prompt value = {0}", prompt), 11, 4)
                If prompt.IndexOf(","c) > 0 Then
                    Log("Comma was found in prompt. Parsing values", 11, 4)
                    If IsNumeric(prompt.Substring(prompt.IndexOf(","c) + 1)) Then
                        Try
                            If Not Integer.TryParse(prompt.Substring(prompt.IndexOf(","c) + 1), partialMaskLength) Then
                                partialMaskLength = 10
                                Log(String.Format("Unable to parse partial mask length. Setting to {0}", partialMaskLength), 11, 4)
                            End If
                        Catch ex As Exception
                            partialMaskLength = 10
                            Log(String.Format("Unable to parse partial mask length. Setting to {0}. Exception: {1}",
                                                      partialMaskLength, ex.Message), 11, 9004)
                        End Try
                    Else
                        partialMaskLength = 10
                    End If
                    Log(String.Format("Partial mask length = {0}", partialMaskLength), 11, 4)
                    prompt = prompt.Substring(0, prompt.IndexOf(","c)).Replace("'", "")
                    Log(String.Format("Prompt = {0}", prompt), 11, 4)
                    maskText = String.Format("substr(prompt_response, 1, {0}) || substr('XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX', 1, strlen(prompt_response) - {0})",
                                             partialMaskLength)
                    Log(String.Format("Masked Text = {0}", maskText), 11, 4)
                Else
                    Log("Comma was not found in prompt. Response will be masked with generic mask", 11, 4)
                    partialMaskLength = 0
                    Log(String.Format("Partial mask length = {0}", partialMaskLength), 11, 4)
                    maskText = "'MASKED'"
                    Log(String.Format("Masked Text = {0}", maskText), 11, 4)
                End If

                Try
                    Log("Creating masking query", 5, 4)
                    writer = New StreamWriter(New FileStream(querySql, FileMode.Create, FileAccess.Write))
                    writer.WriteLine(String.Format(template, cardname.Replace("'", "")))
                    writer.Close()
                Catch ex As Exception
                    Log(String.Format("Error creating masking query. Exception: {0}", ex.Message), 11, 9004)
                End Try

                ' a blank file will cause the masking to be skipped
                result = InterbaseQuery(String.Format(template, cardname))

                minSequence = "0"
                Dim iMinSeq, iMaxSeq As Integer

                If result <> "" Then
                    Log("Result was not null. Attempting to parse pmcode and pmsubcode from result.", 11, 4)
                    If result.IndexOf(",") > -1 AndAlso result.IndexOf(",") < 10 Then
                        pmntcode_id = result.Split(","c)(0)
                        pmsubcode_id = result.Split(","c)(1)
                    Else
                        pmntcode_id = result.Substring(0, 5).Trim()
                        pmsubcode_id = result.Substring(6).Trim()
                    End If
                    Log(String.Format("pmntcode_id = {0}, pmsubcode_id = {1}", pmntcode_id, pmsubcode_id), 11, 4)
                    If IsNumeric(pmntcode_id) Then

                        Log(String.Format("Getting max(sernum) from tillpmnt where pmcode = {0} and pmsubcode = {1}", pmntcode_id, pmsubcode_id), 11, 3)

                        minSequence = GetRegistryValue("HKLM", "Software\PilotInfo\ImportData", String.Format("{0}_{1}", cardname.Trim(), prompt), "0")
                        maxSequence = InterbaseQuery(String.Format(maxSeq, pmntcode_id, pmsubcode_id))

                        If Not Integer.TryParse(minSequence, iMinSeq) Then
                            Log("Failed to parse integer from min sequence, setting min sequence to 0.", 11, 9004)
                            iMinSeq = 0
                        End If

                        If Not Integer.TryParse(maxSequence, iMaxSeq) Then
                            Log("Failed to parse integer from max sequence, setting max sequence to 0.", 11, 9004)
                            iMaxSeq = 0
                        End If

                        Log(String.Format("Min Sequence = {0}", iMaxSeq.ToString), 11, 4)
                        Log(String.Format("Max Sequence = {0}", iMaxSeq.ToString), 11, 4)


                        If iMinSeq < iMaxSeq Then 'If maxSequence <> "" AndAlso maxSequence <> "<null>" Then
                            Log(String.Format("{0} (min) < {1} (max) evaluates to True, masking will occur", iMinSeq.ToString, iMaxSeq.ToString), 11, 4)
                            Log(String.Format("Masking {0} prompt for {1}", prompt, cardname), 11, 4)
                            result = InterbaseQuery(String.Format(maskSQL, maskText, pmntcode_id, pmsubcode_id, prompt, iMinSeq.ToString))

                            Log("Done masking prompt.", 11, 4)
                            Log(String.Format("Updating prompt registry value with max sequence = {0}", iMaxSeq.ToString), 11, 4)
                            RegistryFunction("HKLM", "Software\PilotInfo\ImportData", String.Format("{0}_{1}", cardname.Trim(), prompt), False, iMaxSeq.ToString)
                        Else
                            If iMinSeq > iMaxSeq Then
                                Log(String.Format("{0} (min) > {1} (max) evaluates to True, updating registry", iMinSeq.ToString, iMaxSeq.ToString), 11, 4)
                                iMinSeq = iMaxSeq
                                'Log(String.Format("Masking {0} prompt for {1}", prompt, cardname), 10, 4)
                                'result = InterbaseQuery(String.Format(maskSQL, maskText, pmntcode_id, pmsubcode_id, prompt, iMinSeq.ToString), -1)

                                'Log("Done masking prompt.", 10, 4)
                                Log(String.Format("Updating prompt registry value with max sequence = {0}", iMaxSeq.ToString), 11, 4)
                                RegistryFunction("HKLM", "Software\PilotInfo\ImportData", String.Format("{0}_{1}", cardname.Trim(), prompt), False, iMaxSeq.ToString)
                            Else
                                Log(String.Format("{0} (min) < {1} (max) evaluates to False, masking will not occur", iMinSeq.ToString, iMaxSeq.ToString), 11, 4)
                            End If

                        End If
                    Else
                        Log(String.Format("pmcode was not numeric: {0}", pmntcode_id), 11, 4)
                    End If
                Else
                    Log("Result was null. Masking will not occur", 11, 4)
                End If

            Next
        Catch ex As Exception
            Log("Error occurred with masking. Exception: " & ex.Message, 1, 9004)
        End Try
        Log("ImportData masking done", 3, 4)
        'mLogger.Close()
        Log("Exiting MaskPrompts() Sub", 10, 4)
    End Sub

    Private Sub AutoPostSales()
        Log("Entering AutoPostSales() Sub", 10, 5)
        If DoExternalEvent Then
            Log("DoExternalEvent is True. Continuing checks.", 5, 6)
            If mIsR10Site Then
                Log("Site is configured as R10 site. ImportData should not run AutoPostSales for R10 locations. Setting DoAutoPostSales to False.", 5, 6)
                mDoAutoPostSales = False
            End If
            If mDoAutoPostSales Then
                ProgressDialog.AddProgress("Calling Auto Post Sales")
                Log("DoAutoPostSales is True. External event runner will be called to execute autopostsales", 5, 6)
                CallExternalEventRunner("autopostsales")
            Else
                Log("DoAutoPostSales is False. External event runner will NOT be called to execute autopostsales", 5, 6)
            End If

        Else
            Log("DoExternalEvent is false. External event runner will not be called to execute autopostsales", 5, 6)
        End If
        Log("Exiting AutoPostSales() Sub", 10, 5)
    End Sub

    Private Sub AutoPriceSrv()
        Log("Entering AutoPriceSrv() Sub", 10, 6)
        If mDoAutoPriceSrv Then
            ProgressDialog.AddProgress("Calling Auto Price Srv")
            Log("DoAutoPriceSrv is True. External event runner will be called to execute autopricesrv", 5, 6)
            CallExternalEventRunner("autopricesrv")
        Else
            Log("DoAutoPriceSrv is False. External event runner will NOT be called to execute autopricesrv", 5, 6)
        End If

        Log("Exiting AutoPriceSrv() Sub", 10, 6)
    End Sub

    Private Sub CallCrossSiteFix()
        Log("Entering CallCrossSiteFix() Sub", 10, 7)
        ProgressDialog.AddProgress("Calling Fix_cross_site procedure")
        Log("Call fix_cross_site procedure started", 2, 7)
        'InterbaseQuery(FIX_CROSS_SITE_PROCEDURE)
        InterbaseExecuteProcedure("fix_cross_site")
        Log("Call fix_cross_site procedure ended", 3, 7)
        Log("Exiting CallCrossSiteFix() Sub", 10, 7)
    End Sub

    Private Sub CallExternalEventRunner(ByVal eventName As String)
        Log("Entering CallExternalEventRunner(eventName) Sub", 10, 8)
        Dim counter As Integer = 0
        Dim returnVal As Integer

        If Not ProcessRunning("Import") Then
            Log("Import not running - good", 5, 8)
            If ProcessRunning("ExternalEventRunner") Then
                Log("*** ExternalEventRunner currently running", 2, 8)
                While ProcessRunning("ExternalEventRunner") AndAlso counter < mExternalEventRetry
                    Threading.Thread.Sleep(mExternalEventDelay * 1000)
                    Log(String.Format("Waited {0} seconds", mExternalEventDelay), 10, 8)
                    counter += 1
                End While
                If ProcessRunning("ExternalEventRunner") Then
                    Log(String.Format("Waited {0} seconds for ExternalEventRunner to complete.... ", mExternalEventRetry * mExternalEventDelay), 3, 8)
                    If mKillExternalEventRunnerAfterMaxRetries Then
                        Log("Warning - ExternalEventRunnerKillOnMaxRetries is set to True and max retries was reached.", 5, 8)
                        KillProcess("ExternalEventRunner.exe")
                    Else
                        Log("ExternalEventRunnerKillOnMaxRetries is set to False. Exiting...", 5, 8)
                        Log("Exiting CallExternalEventRunner(eventName) Sub", 10, 8)
                        Return
                    End If

                End If
            End If

            Log(String.Format("Call ExternalEventRunner.exe /{0}", eventName), 1, 8)
            'mLogger.Flush()

            Log(String.Format("External event runner timer set to {0} minutes", mExternalEventTimer), 5, 6)
            returnVal = Shell("C:\office\exe\externaleventrunner.exe /" & eventName, AppWinStyle.Hide, True, mExternalEventTimer * 60000)

            If returnVal <> 0 Then
                Log(String.Format("ExternalEventRunner failed to complete in {0} minutes.... terminating process", mExternalEventTimer), 1, 8)
                Log(String.Format("(ProcessId: {0})", returnVal), 2, 8)
                Dim aProcess As Process
                aProcess = Process.GetProcessById(returnVal)
                aProcess.Kill()
            Else
                Log("Call ExternalEventRunner.exe completed successfully", 2, 8)
            End If

        Else
            Log("Skipping call to ExternalEventRunner due to Import running", 2, 8)
        End If
        Log("Exiting CallExternalEventRunner(eventName) Sub", 10, 8)
    End Sub

    Private Sub CallPilotMasking()
        Log("Entering CallPilotMasking() Sub", 10, 9)
        If DoMaskPrompts Then
            ProgressDialog.AddProgress("Calling Pilot_masking procedure")
            Log("Call Pilot_masking procedure started", 2, 9)
            'InterbaseQuery(PILOT_MASKING_PROCEDURE)
            InterbaseExecuteProcedure("pilot_masking")
            Log("Call Pilot_masking procedure ended", 3, 9)
        Else
            Log("Call Pilot_masking procedure is not enabled. Masking will not occur.", 2, 9)
        End If
        Log("Exiting CallPilotMasking() Sub", 10, 9)
    End Sub

    Private Sub CheckForFailureLog()
        Log("Entering CheckForFailureLog() Sub", 10, 10)
        Dim FileName As String
        Dim dt As DateTime
        Dim counter As Integer = -3

        Do
            dt = DateAdd(DateInterval.Minute, counter, Now)
            FileName = Path.Combine("C:\Office\Log", String.Format("ImpFailure{0}.log", dt.ToString("yyyyMMddHHmm")))
            Log(String.Format("Import failure log set to {0}", FileName), 10, 10)
            If File.Exists(FileName) Then
                Log(String.Format("{0} failure log was found. Sending email.", FileName), 1, 9010)
                Dim Message As New Text.StringBuilder(100)
                Message.Append(FileName & " was found at store " & StoreNumber & "." & Chr(13))
                Message.Append(ReadFile(FileName))

                EmailError(mMailRecipient, "Import Failure Log", Message.ToString())
                Log("Exiting CheckForFailure() Sub", 10, 10)
                Exit Sub
            Else
                Log("Import failure log was not found - good", 10, 10)
            End If

            counter += 1
        Loop While counter <= 1
        Log("Exiting CheckForFailure() Sub", 10, 10)
    End Sub

    Private Function CheckForPendingPricebatches(ByVal SendEmail As Boolean) As Integer

        Log("Entering CheckForPendingPricebatches() Sub", 10, 11)
        Dim CurrDate As String = Date.Now.ToString("MM/dd/yyy")
        Dim addlquery As String = ""
        'Dim PENDING_PRICE_BATCHES_QUERY As String = "select count(*) from pricemodhdr where status = 'P' and (moddate > Cast('today' as date) - {0}) and (moddate <= Cast('today' as date));"
        Dim PENDING_PRICE_BATCHES_QUERY As String
        If mPriceBatchDays <> 0 Then
            Log(String.Format("Price batch days is set to {0}. Querying for pending price batches older than", mPriceBatchDays), 10, 11)
            Dim OldDate As String = DateAdd(DateInterval.Day, mPriceBatchDays * -1, Date.Now).ToString("MM/dd/yyyy")
            Log(String.Format("Old Date = {0}", OldDate), 10, 11)
            Log(String.Format("Current Date = {0}", CurrDate), 10, 11)
            Log(String.Format("Price batch days is set to {0}. Querying for date range less than oldest date", mPriceBatchDays), 10, 11)
            If mIncludeNullPriceModHdrDates Then
                addlquery = " or moddate IS NULL"
            End If
            PENDING_PRICE_BATCHES_QUERY = String.Format("select count(*) from pricemodhdr where status {0} and (moddate < '{1}'{2});", mPriceModHdrCriteria, OldDate, addlquery)
        Else
            Log(String.Format("Price batch days is set to {0}. Querying for all days.", mPriceBatchDays), 10, 11)
            Log(String.Format("Current Date = {0}", CurrDate), 10, 11)
            If mIncludeNullPriceModHdrDates Then
                addlquery = " or moddate IS NULL"
            End If
            PENDING_PRICE_BATCHES_QUERY = String.Format("select count(*) from pricemodhdr where status {0} and (moddate <= '{1}'{2});", mPriceModHdrCriteria, CurrDate, addlquery)
        End If
        Log(String.Format("Query: {0}", PENDING_PRICE_BATCHES_QUERY), 10, 11)

        Log(String.Format("Pending price batch check times contains {0} items in the list", mPendingPriceBatchesCheckTimes.Count), 10, 11)

        For i As Integer = 0 To mPendingPriceBatchesCheckTimes.Count - 1
            Log(String.Format("Item {0} in hour list: {1} hr", i + 1, mPendingPriceBatchesCheckTimes.Item(i)), 10, 11)
        Next



        If mPendingPriceBatchesCheckTimes.Contains(Date.Now.Hour) Then
            Log(String.Format("Hour list for checking  pending price batch contains {0}. Checking for pending price batches", Date.Now.Hour), 10, 11)
            Dim results As String
            Dim iResults As Integer

            ProgressDialog.AddProgress("Checking for pending pricebatches")
            Log("Call pending pricebatches sql", 2, 11)
            results = InterbaseQuery(PENDING_PRICE_BATCHES_QUERY) 'InterbaseQuery(String.Format(PENDING_PRICE_BATCHES_QUERY, mPriceBatchDays))

            Log("Call pending pricebatches sql ended", 5, 11)

            results = results.Trim(" "c)

            If Not Integer.TryParse(results, iResults) Then
                Log("Failed to parse results to integer. Setting results to 0.", 1, 9011)
                iResults = 0
            End If

            Log(String.Format("This location has {0} pending pricebatch(es) older than {1} days or null moddate.", iResults, mPriceBatchDays), 5, 11)
            If iResults <> 0 And SendEmail Then

                Dim dtPendingPriceBatchEmailDateTime As DateTime
                Dim PendingPriceBatchEmailDateTime As String = GetRegistryValue("HKLM", "Software\PilotInfo\ImportData\Email", "LastEmailDateTime", "")

                If Not DateTime.TryParse(PendingPriceBatchEmailDateTime, dtPendingPriceBatchEmailDateTime) Then
                    Log(String.Format("Unable to parse date from HKLM\Software\PilotInfo\ImportData\Email[LastEmailDateTime]. Current value is [{0}]", PendingPriceBatchEmailDateTime), 5, 11)
                    dtPendingPriceBatchEmailDateTime = DateTime.Now().AddHours(-72)
                    Log(String.Format("Updating the value to [{0}]", dtPendingPriceBatchEmailDateTime.ToString), 5, 11)
                    RegistryFunction("HKLM", "Software\PilotInfo\ImportData\Email", "LastEmailDateTime", False, dtPendingPriceBatchEmailDateTime.ToString)
                Else
                    Log(String.Format("Successfully parsed date from HKLM\Software\PilotInfo\ImportData\Email[LastEmailDateTime]. Current value is [{0}]", PendingPriceBatchEmailDateTime), 5, 11)
                End If

                Dim startTime As DateTime = dtPendingPriceBatchEmailDateTime
                Dim endTime As DateTime = DateTime.Now
                Dim duration As TimeSpan = endTime - startTime        'Subtract start time from end time
                Log(String.Format("Duration ({0}) = End Time ({1}) - Start Time ({2})", duration.TotalHours.ToString, endTime.ToString, startTime.ToString), 5, 11)
                Dim HoursSinceLastRun As Integer = CInt(duration.TotalHours)
                Log(String.Format("Hours since last run: {0}", HoursSinceLastRun), 5, 11)

                If HoursSinceLastRun >= mPendingPriceBatchEmailDelayInHours Then
                    Log(String.Format("Last run {0} hours >= {1} hours delay between emails. Email will be sent.", HoursSinceLastRun, mPendingPriceBatchEmailDelayInHours), 5, 11)
                    Log(String.Format("Results > 0. Pending price batch warning email will be sent", iResults, mPriceBatchDays), 5, 9011)
                    EmailError(mPricebatchMailRecipient, "Pending Price Batch Warning",
                            String.Format("This location has {0} pending pricebatch(es) older than {1} days.",
                                            iResults, mPriceBatchDays))
                    RegistryFunction("HKLM", "Software\PilotInfo\ImportData\Email", "LastEmailDateTime", False, DateTime.Now().ToString)
                Else
                    Log(String.Format("Last run {0} hours < {1} hours delay between emails. Email will not be sent.", HoursSinceLastRun, mPendingPriceBatchEmailDelayInHours), 5, 11)
                End If

                Log("Exiting CheckForPendingPricebatches() Sub", 10, 11)
                Return iResults

            Else
                Log(String.Format("Results = 0. No email will be sent", iResults, mPriceBatchDays), 5, 11)
            End If
        Else
            Log("Not scheduled to check for pending pricebatches this hour", 5, 11)
        End If
        Log("Exiting CheckForPendingPricebatches() Sub", 10, 11)
        Return 0
    End Function

    Private Function CheckForShippers() As Boolean
        Log("Entering CheckForShippers() Function (As Boolean)", 10, 12)
        ProgressDialog.AddProgress("Checking for shippers")
        Log("Checking for shipper file", 3, 12)
        If File.Exists(RCV_FILE_PATH & "\SHeader.dat") Then
            Log(String.Format("Shipper file found as {0}", RCV_FILE_PATH & "\SHeader.dat"), 4, 12)
            Log("Exiting CheckForShippers() Function (True)", 10, 12)
            Return True
        Else
            Log("No shipper file found", 4, 12)
            Log("Exiting CheckForShippers() Function (False)", 10, 12)
            Return False
        End If
    End Function

    Private Function GetConfigOption(ByVal optionName As String, ByVal defaultValue As String, Optional ByVal Section As String = "") As String

        Dim result As String
        If Section <> "" Then
            Try
                Dim nvc As NameValueCollection = TryCast(ConfigurationManager.GetSection(Section), NameValueCollection)
                result = nvc(optionName).Trim()
                If result = "0" Then
                    result = defaultValue
                End If
            Catch ex As Exception
                result = defaultValue
            End Try

            'Dim config As System.Configuration.Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
            'Dim appSettingSection As NameValueConfigurationCollection = config.GetSection(Section)
            'Try
            '    result = appSettingSection.Settings(optionName).Value.Trim()
            '    If result = "0" Then
            '        result = defaultValue
            '    End If
            'Catch ex As Exception
            '    result = defaultValue
            'End Try
        Else

            Try
                If ConfigurationManager.AppSettings(optionName) Is Nothing Then
                    result = defaultValue
                ElseIf ConfigurationManager.AppSettings(optionName).Trim() = "0" Then
                    result = defaultValue
                Else
                    result = ConfigurationManager.AppSettings(optionName)
                End If
            Catch ex As Exception
                result = defaultValue
            End Try

        End If

        Return result
    End Function

    Private Function ValidatePriceModHdrCriteria(ByVal criteria As String) As String
        Dim resp As String() = criteria.Split(";"c)
        Select Case resp(0).Trim.ToUpper
            Case "=P", "= P", "='P'", "= 'P'"
                Return "= 'P'"
            Case "=A", "= A", "='A'", "= 'A'"
                Return "= 'A'"
            Case "!C", "! C", "NOT C", "!'C'", "! 'C'", "NOT 'C'", "NOT'C'"
                Return "<> 'C'"
            Case Else
                Return "= 'P'"
        End Select
        Return ""
    End Function

    Private Sub GetConfigOptions()
        Log("Entering GetConfigOptions()", 10, 102)
        Dim StoreInt As Integer
        Dim ForceQAMode As Boolean = False
        Dim ForceProductionMode As Boolean = False
        If Not Integer.TryParse(StoreNumber, StoreInt) Then
            Log(String.Format("Failed to parse Store number into integer (Store number was passed as {0} which was unable to be parsed to Integer)", StoreNumber), 1, 9102)
            StoreInt = 0
        Else
            Log(String.Format("Parsed Store # to Integer successfully. Store Int = {0}", StoreInt.ToString), 1, 102)
        End If

        If (StoreInt >= 9000 AndAlso StoreInt <= 9999) Then
            Log("Store number >= 9000 AndAlso <= 9999. Store number is assumed to be a lab.", 10, 102)
            SectionVal = "QA"
        Else
            SectionVal = "Production"
        End If
        Log(String.Format("Configuration will be loaded from the {0} section of the configuration file.", SectionVal), 1, 102)

        If Not Integer.TryParse(GetConfigOption("LogLevel", "5", SectionVal), mLogLevel) Then
            Log("Failed to parse LogLevel integer value. Setting to default.", 1, 9102)
            mLogLevel = 5
        Else
            If mLogLevel < 0 OrElse mLogLevel > 20 Then
                Log("LogLevel was set < 0 or > 20. Setting to default.", 1, 9102)
                mLogLevel = 5
            End If
        End If
        Log(String.Format("LogLevel = {0}", mLogLevel.ToString), 6, 102)
        mLogger.LogLevel = mLogLevel

        Dim ModeChanges As Boolean = False

        If Not Boolean.TryParse(GetConfigOption("ForceQAMode", "False", "General"), ForceQAMode) Then
            Log("Failed to parse ForceQAMode boolean value. Setting to default.", 1, 9102)
            ForceQAMode = False
        End If
        Log(String.Format("ForceQAMode = {0}", ForceQAMode.ToString), 6, 102)

        If ForceQAMode = True Then
            If SectionVal <> "QA" Then
                SectionVal = "QA"
                Log(String.Format("Section value changed. Updating Section value to {0}", SectionVal), 6, 102)
                ModeChanges = True
            End If
        End If

        If Not Boolean.TryParse(GetConfigOption("ForceProductionMode", "False", "General"), ForceQAMode) Then
            Log("Failed to parse ForceProductionMode boolean value. Setting to default.", 1, 9102)
            ForceProductionMode = False
        End If
        Log(String.Format("ForceProductionMode = {0}", ForceProductionMode.ToString), 6, 102)

        If ForceProductionMode = True Then
            If SectionVal <> "Production" Then
                SectionVal = "Production"
                Log(String.Format("Section value changed. Updating Section value to {0}", SectionVal), 6, 102)
                ModeChanges = True
            End If
        End If

        If ModeChanges Then
            Log("Mode changed. Reloading LogLevel.", 1, 102)
            If Not Integer.TryParse(GetConfigOption("LogLevel", "5", SectionVal), mLogLevel) Then
                Log("Failed to parse LogLevel integer value. Setting to default.", 1, 9102)
                mLogLevel = 5
            Else
                If mLogLevel < 0 OrElse mLogLevel > 20 Then
                    Log("LogLevel was set < 0 or > 20. Setting to default.", 1, 9102)
                    mLogLevel = 5
                End If
            End If
            Log(String.Format("LogLevel = {0}", mLogLevel.ToString), 6, 102)
            mLogger.LogLevel = mLogLevel
        End If


        'Load Boolean Settings
        'LogDebugMessages = CBool(GetConfigOption("Debug", "False"))
        If Not Boolean.TryParse(GetConfigOption("Debug", "False", SectionVal), LogDebugMessages) Then
            Log("Failed to parse Debug boolean value. Setting to default.", 1, 9102)
            LogDebugMessages = False
        End If
        Log(String.Format("Debug = {0}", LogDebugMessages.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("PendingPriceBatchSendEmailAlerts", "True", SectionVal), mSendPendingPriceBatchEmails) Then
            Log("Failed to parse PendingPriceBatchSendEmailAlerts boolean value. Setting to default.", 1, 9102)
            mSendPendingPriceBatchEmails = True
        End If
        Log(String.Format("PendingPriceBatchEmails = {0}", mSendPendingPriceBatchEmails.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("MaskPrompts", "True", SectionVal), DoMaskPrompts) Then
            Log("Failed to parse MaskPropts boolean value. Setting to default.", 1, 9102)
            DoMaskPrompts = True
        End If
        Log(String.Format("MaskPrompts = {0}", DoMaskPrompts.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("PendingPostSalesContinueIfTrue", "True", SectionVal), mPendingPostSalesContinueIfTrue) Then
            Log("Failed to parse PendingPostSalesContinueIfTrue boolean value. Setting to default.", 1, 9102)
            mPendingPostSalesContinueIfTrue = True
        End If
        Log(String.Format("PendingPostSalesContinueIfTrue = {0}", mPendingPostSalesContinueIfTrue.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("PendingPostSalesForceSetToTrue", "False", SectionVal), mPendingPostSalesForceSetToTrue) Then
            Log("Failed to parse PendingPostSalesForceSetToTrue boolean value. Setting to default.", 1, 9102)
            mPendingPostSalesForceSetToTrue = False
        End If
        Log(String.Format("PendingPostSalesForceSetToTrue = {0}", mPendingPostSalesForceSetToTrue.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("IncludeNullPriceModHdrDates", "False", SectionVal), mIncludeNullPriceModHdrDates) Then
            Log("Failed to parse IncludeNullPriceModHdrDates boolean value. Setting to default.", 1, 9102)
            mIncludeNullPriceModHdrDates = False
        End If
        Log(String.Format("IncludeNullPriceModHdrDates = {0}", mIncludeNullPriceModHdrDates.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("AutoPriceSrvOnPriceModHdrRecordsReturned", "False", SectionVal), mAutoPriceSrvOnPriceModHdrRecordsReturned) Then
            Log("Failed to parse AutoPriceSrvOnPriceModHdrRecordsReturned boolean value. Setting to default.", 1, 9102)
            mAutoPriceSrvOnPriceModHdrRecordsReturned = False
        End If
        Log(String.Format("AutoPriceSrvOnPriceModHdrRecordsReturned = {0}", mAutoPriceSrvOnPriceModHdrRecordsReturned.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("ExternalEventRunnerKillOnMaxRetries", "False", SectionVal), mKillExternalEventRunnerAfterMaxRetries) Then
            Log("Failed to parse ExternalEventRunnerKillOnMaxRetries boolean value. Setting to default.", 1, 9102)
            mKillExternalEventRunnerAfterMaxRetries = False
        End If
        Log(String.Format("ExternalEventRunnerKillOnMaxRetries = {0}", mKillExternalEventRunnerAfterMaxRetries.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("IsR10Site", "False", SectionVal), mIsR10Site) Then
            Log("Failed to parse IsR10Site boolean value. Setting to default.", 1, 9102)
            mIsR10Site = False
        End If
        Log(String.Format("IsR10Site = {0}", mIsR10Site.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("DoFixItWithoutDatFiles", "True", SectionVal), mDoFixItWithoutDatFiles) Then
            Log("Failed to parse DoFixItWithoutDatFiles boolean value. Setting to default.", 1, 9102)
            mDoFixItWithoutDatFiles = True
        End If
        Log(String.Format("DoFixItWithoutDatFiles = {0}", mDoFixItWithoutDatFiles.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("DoAutoPostSales", "True", SectionVal), mDoAutoPostSales) Then
            Log("Failed to parse DoAutoPostSales boolean value. Setting to default.", 1, 9102)
            mDoAutoPostSales = True
        End If
        Log(String.Format("DoAutoPostSales = {0}", mDoAutoPostSales.ToString), 6, 102)

        If Not Boolean.TryParse(GetConfigOption("DoAutoPriceSrv", "True", SectionVal), mDoAutoPriceSrv) Then
            Log("Failed to parse DoAutoPriceSrv boolean value. Setting to default.", 1, 9102)
            mDoAutoPriceSrv = True
        End If
        Log(String.Format("DoAutoPriceSrv = {0}", mDoAutoPriceSrv.ToString), 6, 102)

        mMailSender = GetConfigOption("MailSender", "DoNotReply@pilottravelcenters.com", SectionVal)
        mMailSender = ValidateEmailByDefault(mMailSender, "DoNotReply@pilottravelcenters.com")
        Log(String.Format("MailSender = {0}", mMailSender), 6, 102)

        Dim ListHour As Integer
        'Load String Settings
        mMyUser = GetConfigOption("EncryptedUser", "kGImGjO2wSn3WSY2kOjo/w==", SectionVal)
        mMyPass = GetConfigOption("EncryptedPass", "17P1PL0A7Z5jfcH9HIEqtg==", SectionVal)
        mMailRecipient = GetConfigOption("EmailRecipient", "TicketTracker.HelpDesk@pilottravelcenters.com", SectionVal)
        Log(String.Format("EmailRecipient = {0}", mMailRecipient), 6, 102)

        mOFBMailRecipient = GetConfigOption("OFBMissingEmailRecipient", "Dustin.Harmon@pilottravelcenters.com;
                                            Nathan.Bowers@pilottravelcenters.com;Travis.Russell@pilottravelcenters.com;
                                            Jennifer.McClurkan@pilottravelcenters.com;Dean.Day@pilottravelcenters.com", SectionVal)
        Log(String.Format("OFBMissingEmailRecipient = {0}", mOFBMailRecipient), 6, 102)


        mPricebatchMailRecipient = GetConfigOption("PendingPricebatchEmailRecipient", "PriceBatchAlert@pilottravelcenters.com;
                                                    TicketTracker.HelpDesk@pilottravelcenters.com", SectionVal)
        Log(String.Format("PendingPricebatchEmailRecipient = {0}", mPricebatchMailRecipient), 6, 102)

        mRcvDatAgeWarningMailRecipient = GetConfigOption("RcvDatFileAgeWarningEmailRecipient", "PriceBatchAlert@pilottravelcenters.com;
                                                    TicketTracker.HelpDesk@pilottravelcenters.com", SectionVal)
        Log(String.Format("RcvDatFileAgeWarningEmailRecipient = {0}", mRcvDatAgeWarningMailRecipient), 6, 102)

        mRelayServer = GetConfigOption("RelayServer", "pilotrelay")
        If mRelayServer = "" Then
            Log("RelayServer is null. Setting to default.", 1, 9102)
            mRelayServer = "pilotrelay"
        End If
        Log(String.Format("RelayServer = {0}", mRelayServer), 6, 102)

        ASNDataPath = GetConfigOption("ASNDataPath", "C:\Pilot\data\", SectionVal)
        Log(String.Format("ASNDataPath = {0}", ASNDataPath), 6, 102)

        mPriceModHdrCriteria = GetConfigOption("PriceModHdrCriteria", "='P';Options are ='P', ='A', !'C', NOT 'C'", SectionVal)
        If mPriceModHdrCriteria.Trim = "" Then
            mPriceModHdrCriteria = "= 'P'"
        End If
        mPriceModHdrCriteria = ValidatePriceModHdrCriteria(mPriceModHdrCriteria)
        Log(String.Format("PriceModHdrCriteria = {0}", mPriceModHdrCriteria), 6, 102)

        'mPriceModHdrDateField = GetConfigOption("PriceModHdrCriteria", "='P';Options are ='P', ='A', !'C', NOT 'C'", SectionVal)
        'If mPriceModHdrDateField.Trim = "" Then
        '    mPriceModHdrDateField = "= 'P'"
        'End If
        'mPriceModHdrCriteria = ValidatePriceModHdrCriteria(mPriceModHdrCriteria)
        'Log(String.Format("PriceModHdrCriteria = {0}", mPriceModHdrCriteria), 6, 102)

        'Load Integer Settings
        If Not Integer.TryParse(GetConfigOption("LogFileMaxSizeInMB", "5", SectionVal), mLogFileMaxSize) Then
            Log("Failed to parse LogFileMaxSizeInMB integer value. Setting to default.", 1, 9102)
            mLogFileMaxSize = 5
        Else
            If mLogFileMaxSize < 1 OrElse mLogFileMaxSize > 20 Then
                Log("LogFileMaxSizeInMB is < 1 or > 20. Setting to default.", 1, 9102)
                mLogFileMaxSize = 5
            End If
        End If
        Log(String.Format("LogFileMaxSizeInMB = {0}", mLogFileMaxSize.ToString), 6, 102)

        'Load Integer Settings
        For Each hour As String In GetConfigOption("PendingPricebatchEmailHours", "2,8,12", SectionVal).Split(","c)
            If Integer.TryParse(hour, ListHour) Then
                If ListHour > 0 AndAlso ListHour < 24 Then
                    mPendingPriceBatchesCheckTimes.Add(ListHour)
                Else
                    Log(String.Format("Hour = {0} integer value from PendingPriceBatchEmailHours is < 0 or > 24. Value will not be added to list.", hour), 1, 9102)
                End If
            Else
                Log(String.Format("Failed to parse hour = {0} integer value from PendingPriceBatchEmailHours. Value will not be added to list.", hour), 1, 9102)
            End If
        Next

        If Not Integer.TryParse(GetConfigOption("PendingPriceBatchDelayBetweenEmailsInHours", "24", SectionVal), mPendingPriceBatchEmailDelayInHours) Then
            Log("Failed to parse PendingPriceBatchDelayBetweenEmailsInHours integer value. Setting to default.", 1, 9102)
            mPendingPriceBatchEmailDelayInHours = 24
        Else
            If mPendingPriceBatchEmailDelayInHours < 0 OrElse mPendingPriceBatchEmailDelayInHours > 72 Then
                Log("PendingPriceBatchDelayBetweenEmailsInHours is < 0 or > 72. Setting to default.", 1, 9102)
                mPendingPriceBatchEmailDelayInHours = 24
            End If
        End If
        Log(String.Format("PendingPriceBatchDelayBetweenEmailsInHours = {0}", mPendingPriceBatchEmailDelayInHours.ToString), 6, 102)


        If Not Integer.TryParse(GetConfigOption("PendingPriceBatchDays", "3", SectionVal), mPriceBatchDays) Then
            Log("Failed to parse PendingPriceBatchDays integer value. Setting to default.", 1, 9102)
            mPriceBatchDays = 3
        Else
            If mPriceBatchDays < 0 OrElse mPriceBatchDays > 31 Then
                Log("PendingPriceBatchDays is < 0 or > 31. Setting to default.", 1, 9102)
                mPriceBatchDays = 3
            End If
        End If
        Log(String.Format("PendingPriceBatchDays = {0}", mPriceBatchDays.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("ImportExeDelayAfterExitInSeconds", "10", SectionVal), mDelayAfterImport) Then
            Log("Failed to parse ImportExeDelayAfterExitInSeconds integer value. Setting to default.", 1, 9102)
            mDelayAfterImport = 10
        Else
            If mDelayAfterImport < 1 OrElse mDelayAfterImport > 300 Then
                Log("ImportExeDelayAfterExitInSeconds is < 1 or > 300. Setting to default.", 1, 9102)
                mDelayAfterImport = 10
            End If
        End If
        Log(String.Format("ImportExeDelayAfterExitInSeconds = {0}", mDelayAfterImport.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("FirstAlert", "3", SectionVal), mFirstAlert) Then
            Log("Failed to parse FirstAlert integer value. Setting to default.", 1, 9102)
            mFirstAlert = 3
        Else
            If mFirstAlert < 0 OrElse mFirstAlert > 23 Then
                Log("FirstAlert is < 0 or > 23. Setting to default.", 1, 9102)
                mFirstAlert = 3
            End If
        End If
        Log(String.Format("FirstAlert = {0}", mFirstAlert.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("NextAlert", "6", SectionVal), mSecondAlert) Then
            Log("Failed to parse NexAlert integer value. Setting to default.", 1, 9102)
            mSecondAlert = 6
        End If

        If mSecondAlert < 0 OrElse mSecondAlert > 23 Then
            Log("NextAlert is < 0 or > 23. Setting to default.", 1, 9102)
            mSecondAlert = 6
        End If
        If mSecondAlert <= mFirstAlert Then
            Log("NextAlert is < FirstAlert. Setting to FirstAlert + 1", 1, 9102)
            mSecondAlert = mFirstAlert + 1
        End If
        Log(String.Format("NextAlert = {0}", mSecondAlert.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("RPOBuildFilesMinHour", "0", SectionVal), mRPOBuildFilesMinHour) Then
            Log("Failed to parse RPOBuildFilesMinHour integer value. Setting to default.", 1, 9102)
            mRPOBuildFilesMinHour = 0
        Else
            If mRPOBuildFilesMinHour < 0 OrElse mRPOBuildFilesMinHour > 23 Then
                Log("RPOBuildFilesMinHour is < 0 or > 23. Setting to default.", 1, 9102)
                mRPOBuildFilesMinHour = 0
            End If
        End If
        Log(String.Format("RPOBuildFilesMinHour = {0}", mRPOBuildFilesMinHour.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("RPOBuildFilesMaxHour", "3", SectionVal), mRPOBuildFilesMaxHour) Then
            Log("Failed to parse RPOBuildFilesMaxHour integer value. Setting to default.", 1, 9102)
            mRPOBuildFilesMaxHour = 3
        End If

        If mRPOBuildFilesMaxHour < 0 OrElse mSecondAlert > 23 Then
            Log("RPOBuildFilesMaxHour is < 0 or > 23. Setting to default.", 1, 9102)
            mRPOBuildFilesMaxHour = 3
        End If
        If mRPOBuildFilesMaxHour <= mRPOBuildFilesMinHour Then
            Log("RPOBuildFilesMaxHour is < RPOBuildFilesMinHour. Setting to RPOBuildFilesMinHour + 1", 1, 9102)
            mRPOBuildFilesMaxHour = mRPOBuildFilesMinHour + 1
        End If
        Log(String.Format("RPOBuildFilesMaxHour = {0}", mRPOBuildFilesMaxHour.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("ExternalEventLaunchProcessTimerInMinutes", "15", SectionVal), mExternalEventTimer) Then
            Log("Failed to parse ExternalEventLaunchProcessTimerInMinutes integer value. Setting to default.", 1, 9102)
            mExternalEventTimer = 15
        Else
            If mExternalEventTimer < 1 Or mExternalEventTimer > 60 Then
                Log("ExternalEventLaunchProcessTimerInMinutes is < 1 or > 60. Setting to default.", 1, 9102)
                mExternalEventTimer = 15
            End If
        End If
        Log(String.Format("ExternalEventLaunchProcessTimerInMinutes = {0}", mExternalEventTimer.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("PostSalesSleepTimeInSeconds", "180", SectionVal), mPostSalesSleep) Then
            Log("Failed to parse PostSalesSleepTimeInSeconds integer value. Setting to default.", 1, 9102)
            mPostSalesSleep = 180
        Else
            If mPostSalesSleep < 30 OrElse mPostSalesSleep > 600 Then
                Log("PostSalesSleepTimeInSeconds is < 30 or > 600. Setting to default.", 1, 9102)
                mPostSalesSleep = 180
            End If
        End If
        Log(String.Format("PostSalesSleepTimeInSeconds = {0}", mPostSalesSleep.ToString), 6, 102)


        If Not Integer.TryParse(GetConfigOption("PostSalesMaxRetryAttempts", "3", SectionVal), mPostSalesRetry) Then
            Log("Failed to parse PostSalesMaxRetryAttempts integer value. Setting to default.", 1, 9102)
            mPostSalesRetry = 3
        Else
            If mPostSalesRetry < 0 OrElse mPostSalesRetry > 10 Then
                Log("PostSalesMaxRetryAttempts is < 0 or > 10. Setting to default.", 1, 9102)
                mPostSalesRetry = 3
            End If
        End If
        Log(String.Format("PostSalesMaxRetryAttempts = {0}", mPostSalesRetry.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("DbUpgraderDelayInSeconds", "60", SectionVal), mDbUpgraderDelay) Then
            Log("Failed to parse DbUpgraderDelayInSeconds integer value. Setting to default.", 1, 9102)
            mDbUpgraderDelay = 60
        Else
            If mDbUpgraderDelay < 1 OrElse mDbUpgraderDelay > 600 Then
                Log("DbUpgraderDelayInSeconds is < 1 or > 600. Setting to default.", 1, 9102)
                mDbUpgraderDelay = 60
            End If
        End If
        Log(String.Format("DbUpgraderDelayInSeconds = {0}", mDbUpgraderDelay.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("DbUpgraderRunningMaxRetryAttempts", "60", SectionVal), mDbUpgraderRetry) Then
            Log("Failed to parse DbUpgraderRunningMaxRetryAttempts integer value. Setting to default.", 1, 9102)
            mDbUpgraderRetry = 60
        Else
            If mDbUpgraderRetry < 1 OrElse mDbUpgraderRetry > 60 Then
                Log("DbUpgraderRunningMaxRetryAttempts is < 1 or > 60. Setting to default.", 1, 9102)
                mDbUpgraderRetry = 60
            End If
        End If
        Log(String.Format("DbUpgraderRunningMaxRetryAttempts = {0}", mDbUpgraderRetry.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("ExternalEventRunnerDelayInSeconds", "60", SectionVal), mExternalEventDelay) Then
            Log("Failed to parse ExternalEventRunnerDelayInSeconds integer value. Setting to default.", 1, 9102)
            mExternalEventDelay = 60
        Else
            If mExternalEventDelay < 1 OrElse mExternalEventDelay > 600 Then
                Log("ExternalEventRunnerDelayInSeconds is < 1 or > 600. Setting to default.", 1, 9102)
                mExternalEventDelay = 60
            End If
        End If
        Log(String.Format("ExternalEventRunnerDelayInSeconds = {0}", mExternalEventDelay.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("ExternalEventRunnerMaxRetryAttempts", "5", SectionVal), mExternalEventRetry) Then
            Log("Failed to parse ExternalEventRunnerMaxRetryAttempts integer value. Setting to default.", 1, 9102)
            mExternalEventRetry = 5
        Else
            If mExternalEventRetry < 1 OrElse mExternalEventRetry > 60 Then
                Log("ExternalEventRunnerMaxRetryAttempts is < 1 or > 60. Setting to default.", 1, 9102)
                mExternalEventRetry = 5
            End If
        End If
        Log(String.Format("ExternalEventRunnerMaxRetryAttempts = {0}", mExternalEventRetry.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("ImportExeMaxRetryAttempts", "5", SectionVal), mImportExeMaxRetryAttempts) Then
            Log("Failed to parse ImportExeMaxRetryAttempts integer value. Setting to default.", 1, 9102)
            mImportExeMaxRetryAttempts = 5
        Else
            If mImportExeMaxRetryAttempts < 1 OrElse mImportExeMaxRetryAttempts > 10 Then
                Log("ImportExeMaxRetryAttempts is < 1 or > 10. Setting to default.", 1, 9102)
                mImportExeMaxRetryAttempts = 5
            End If
        End If
        Log(String.Format("ImportExeMaxRetryAttempts = {0}", mImportExeMaxRetryAttempts.ToString), 6, 102)

        If Not Integer.TryParse(GetConfigOption("RcvDatFileAgeWarningInDays", "3", SectionVal), mRcvDatFilesAgeWarningInDays) Then
            Log("Failed to parse RcvDatFileAgeWarningInDays integer value. Setting to default.", 1, 9102)
            mRcvDatFilesAgeWarningInDays = 3
        Else
            If mRcvDatFilesAgeWarningInDays < 1 OrElse mRcvDatFilesAgeWarningInDays > 60 Then
                Log("RcvDatFileAgeWarningInDays is < 1 or > 60. Setting to default.", 1, 9102)
                mRcvDatFilesAgeWarningInDays = 3
            End If
        End If
        Log(String.Format("RcvDatFileAgeWarningInDays = {0}", mRcvDatFilesAgeWarningInDays.ToString), 6, 102)

        Log("Exiting GetConfigOptions()", 10, 102)
    End Sub

    Public Function GetStoreNumberFromMachineName(ByVal machineName As String) As String

        ' Assume the store number is the first numeric characters of the machine name

        Dim StoreNumber As String = ""

        For i As Integer = 0 To Len(machineName) - 1
            If IsNumeric(machineName.Substring(i, 1)) Then
                StoreNumber &= machineName.Substring(i, 1)
            Else
                Exit For
            End If
        Next

        If StoreNumber.Length = 0 Then StoreNumber = "000"

        Return StoreNumber.PadLeft(3, "0"c)

    End Function

    Private Function GetStoreNumber() As String
        Dim length As Integer
        Dim machine As String = Environment.MachineName

        For Counter As Integer = 5 To 0 Step -1
            If IsNumeric(machine.Substring(0, Counter)) Then
                length = Counter
                Exit For
            End If
        Next

        Return machine.Substring(0, length)
    End Function

    Private Function ImportDataRunning() As Boolean
        If Process.GetProcessesByName(My.Application.Info.AssemblyName).Length > 1 Then
            DbgMsg("ImportData was running.")
            'Windows.Forms.MessageBox.Show("ImportData process already running" & ControlChars.CrLf & "Click OK to close", "ImportData already running", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Dim mLoggerErr As New LogWriter(ErrorLog, funcNum:=-1)
            mLoggerErr.Log(" Script ImportData start attempted, but failed due to process already running", 1, 13)
            mLoggerErr.Log(String.Format("{0} processes running with the process name {1}",
                                      Process.GetProcessesByName(My.Application.Info.AssemblyName).Length.ToString, My.Application.Info.AssemblyName.ToString), 1, 13)
            Dim Line As String
            Try
                Dim LaunchReader As New StreamReader(New FileStream(Path.Combine(AppPath, "LaunchTime.txt"), FileMode.Open, FileAccess.Read))
                Line = LaunchReader.ReadLine()
                LaunchReader.Close()
            Catch ex As Exception
                Line = Date.Now.Ticks.ToString
            End Try


            Dim LDateTime As DateTime = New DateTime(CLng(Line))
            mLoggerErr.Log(String.Format("   Launch time in ticks per last update to LaunchTime.txt: {0}", Line.ToString), 5, 13)
            mLoggerErr.Log(String.Format("   Launch time in datetime per last update to LaunchTime.txt: {0}", LDateTime.ToString), 5, 13)
            Header &= vbCrLf & String.Format("   Launch time in ticks per last update to LaunchTime.txt: {0}", Line.ToString) & vbCrLf
            Header &= vbCrLf & String.Format("   Launch time in datetime per last update to LaunchTime.txt: {0}", LDateTime.ToString) & vbCrLf
            Header &= String.Format("   Current time in ticks: {0}", Date.Now.Ticks.ToString) & vbCrLf
            Header &= String.Format("   Current time in datetime: {0}", Date.Now.ToString) & vbCrLf
            Dim Elapsed As New TimeSpan(Date.Now.Ticks - CLng(Line))
            Dim minutes As Integer = (Elapsed.Days * 24 + Elapsed.Hours) * 60 + Elapsed.Minutes
            Dim CalcMins As String = String.Format("  Minutes as integer {0} = ({1} * 24 + {2}) * 60 + {3}",
                                          minutes, Elapsed.Days, Elapsed.Hours, Elapsed.Minutes)
            Header &= CalcMins & vbCrLf
            mLoggerErr.Log(CalcMins, 5, 13)
            Dim CalcMinsCondensed As String = String.Format("  {0} = ({1} + {2}) * 60 + {3}",
                                          minutes, Elapsed.Days * 24, Elapsed.Hours, Elapsed.Minutes)
            mLoggerErr.Log(CalcMinsCondensed, 5, 13)
            Header &= CalcMinsCondensed & vbCrLf

            CalcMinsCondensed = String.Format("  {0} = {1} * 60 + {2}",
                                          minutes, Elapsed.Days * 24 + Elapsed.Hours, Elapsed.Minutes)
            mLoggerErr.Log(CalcMinsCondensed, 5, 13)
            Header &= CalcMinsCondensed & vbCrLf

            CalcMinsCondensed = String.Format("  {0} = {1} + {2}",
                                          minutes, (Elapsed.Days * 24 + Elapsed.Hours) * 60, Elapsed.Minutes)
            mLoggerErr.Log(CalcMinsCondensed, 5, 13)
            Header &= CalcMinsCondensed & vbCrLf

            If (minutes > (mFirstAlert * 60) AndAlso minutes < ((mFirstAlert + 1) * 60)) OrElse (minutes > mSecondAlert * 60) Then
                CalcMinsCondensed = String.Format("  If ({0} < ({1} * 60) AndAlso {2} < (({3}+1) * 60)) OrElse ({4} > {5} * 60) Then Trigger email alert.",
                                          minutes, mFirstAlert, minutes, mFirstAlert, minutes, mSecondAlert)
                mLoggerErr.Log(CalcMinsCondensed, 5, 13)
                Header &= CalcMinsCondensed & vbCrLf

                CalcMinsCondensed = String.Format("  If ({0} < {1} AndAlso {2} < {3}) OrElse {4} > {5} Then Trigger email alert.",
                                          minutes, mFirstAlert * 60, minutes, (mFirstAlert + 1) * 60, minutes, mSecondAlert * 60)
                mLoggerErr.Log(CalcMinsCondensed, 5, 13)
                Header &= CalcMinsCondensed & vbCrLf

                mLoggerErr.Log(String.Format("  ImportData has been running for {0} minutes. Sending alert e-mail", minutes), 1, 13)
                EmailError(mMailRecipient, "ImportData already running", String.Format("ImportData has been running for at least {0} minutes.
                                            First alert is scheduled {1}. Second alert is scheduled {2}.", minutes, mFirstAlert, mSecondAlert))
            End If
            mLoggerErr.Close()
            Return True
        End If
        Return False
    End Function

    Private Function ImportStillRunning() As Boolean
        Log("Entering ImportData() Function (As Boolean)", 10, 14)
        ProgressDialog.AddProgress("Retalix Import is running")
        Log("Retalix Import is running", 3, 14)
        Dim Processes() As Process = Process.GetProcessesByName("Import")
        Dim Times() As Int32 = {5, 10, 10}
        Dim counter As Integer = 0

        While counter < 3 And Not Processes(0).HasExited
            ProgressDialog.AddProgress(" Waiting " & Times(counter) & " minutes for Import to finish")
            Log("Waiting " & Times(counter) & " minutes for Import to finish", 3, 14)
            mLogger.Flush()
            Sleep(Times(counter) * 60000)
            counter += 1
        End While

        If counter = 3 And Not Processes(0).HasExited Then
            DbgMsg("Import.exe appears to be locked up.")
            EmailError(mMailRecipient, "Import.exe Locked Up", IMPORT_LOCKED_UP)
            Log("Sent e-mail (Import.exe is still running)", 1, 14)
            ProgressDialog.AddProgress(" Import.exe appears locked up")
            ProgressDialog.AddProgress("... exiting")
            ProgressDialog.Done()
            Log("Exiting ImportData() Function (True)", 10, 14)
            Return True
        End If
        Log("Exiting ImportData() Function (False)", 10, 14)
        Return False
    End Function

    Private Function BuildCommand(ByVal query As String) As Odbc.OdbcCommand
        Dim command As New Odbc.OdbcCommand
        command.CommandText = query
        Return command
    End Function

    Private Function InterbaseQuery(ByVal query As String) As String
        Log("Entering InterbaseQuery(query) Function (As String)", 10, 15)
        Dim connection As New Odbc.OdbcConnection()
        Dim connectionString As New Odbc.OdbcConnectionStringBuilder(dbConnectionString)
        connection = New Odbc.OdbcConnection(connectionString.ToString)

        Dim adapter As New Odbc.OdbcDataAdapter(query, connectionString.ToString)
        Dim resultTable As New DataTable
        Dim HasException As Boolean = False
        Log(String.Format("Query = {0}", query), 10, 15)

        Try
            Log("Opening database object", 10, 15)
            connection.Open()
            Log("Filling result table", 10, 15)
            adapter.Fill(resultTable)
            Log("Loaded results into result table successfully", 10, 15)
            'connection.Close()
            'connection.Dispose()
        Catch ex As Exception
            Log(String.Format("Error processing query. Exception: {0}", ex.Message), 1, 9015)
            HasException = True
            'MessageBox.Show(ex.Message)
        Finally
            Log("Closing and disposing database object", 10, 15)
            adapter.Dispose()
            connection.Close()
            connection.Dispose()
        End Try

        If HasException Then
            Log("Returning null due to exception processing query", 10, 9015)
            Log("Exiting InterbaseQuery(query) Function (null)", 10, 15)
            Return ""
        End If

        If resultTable.Rows.Count > 0 Then
            Log(String.Format("{0} rows were returned in results", resultTable.Rows.Count), 10, 15)
            Dim count As Integer = resultTable.Rows(0).ItemArray.Count
            Log(String.Format("Item count for first row: {0}", count.ToString), 10, 15)
            Dim result As String = ""
            For counter As Integer = 0 To count - 1
                If counter = 0 Then
                    result = resultTable.Rows(0).Item(counter).ToString().Trim()
                Else
                    result = result & "," & resultTable.Rows(0).Item(counter).ToString().Trim()
                End If
            Next
            Log("Query Result = " & result, 2, 15)
            Log(String.Format("Exiting InterbaseQuery(query) Function ({0})", result), 10, 15)
            Return result
        Else
            Log("No rows were returned from query", 10, 15)
        End If
        Log("Exiting InterbaseQuery(query) Function (null)", 10, 15)
        Return ""
    End Function

    Private Function InterbaseExecuteProcedure(ByVal procedure As String) As String
        Log("Entering InterbaseExecuteProcedure(procedure) Function (As String)", 10, 16)
        procedure = "execute procedure " & procedure & ";"

        Dim connection As New Odbc.OdbcConnection()
        Dim connectionString As New Odbc.OdbcConnectionStringBuilder(dbConnectionString)
        connection = New Odbc.OdbcConnection(connectionString.ToString)

        Dim adapter As New Odbc.OdbcDataAdapter(procedure, connectionString.ToString)
        Dim resultTable As New DataTable
        Dim HasException As Boolean = False
        Log(String.Format("Procedure = {0}", procedure), 10, 16)

        Try
            Log("Opening database object", 10, 16)
            connection.Open()
            Log("Filling result table", 10, 16)
            adapter.Fill(resultTable)
            Log("Loaded results into result table successfully", 10, 16)
            'connection.Close()
            'connection.Dispose()
        Catch ex As Exception
            Log("Error processing query. Exception: " & ex.Message, 1, 9016)
            HasException = True
            'MessageBox.Show(ex.Message)
        Finally
            Log("Closing and disposing database object", 10, 16)
            adapter.Dispose()
            connection.Close()
            connection.Dispose()
        End Try

        If HasException Then
            Log("Returning null due to exception processing query", 10, 9016)
            Log("Exiting InterbaseExecuteProcedure(procedure) Function (null)", 10, 16)
            Return ""
        End If

        If resultTable.Rows.Count > 0 Then
            Log(String.Format("{0} rows were returned in results", resultTable.Rows.Count), 10, 16)
            Dim count As Integer = resultTable.Rows(0).ItemArray.Count
            Log(String.Format("Item count for first row: {0}", count.ToString), 10, 16)
            Dim result As String = ""
            For counter As Integer = 0 To count - 1
                If counter = 0 Then
                    result = resultTable.Rows(0).Item(counter).ToString().Trim()
                Else
                    result = result & "," & resultTable.Rows(0).Item(counter).ToString().Trim()
                End If
            Next
            Log("Result = " & result, 2, 16)
            Log(String.Format("Exiting InterbaseExecuteProcedure(procedure) Function ({0})", result.Substring(1)), 10, 16)
            Return result
        Else
            Log("No rows were returned from query", 10, 16)
        End If

        Log("Exiting InterbaseExecuteProcedure(procedure) Function (Null Results)", 10, 16)
        Return ""
    End Function

    Private Function InterbaseMultilineQuery(ByVal query As String) As String
        Log("Entering InterbaseMultilineQuery(query,tablename) Function (As String)", 10, 17)
        Dim connection As New Odbc.OdbcConnection()
        Dim connectionString As New Odbc.OdbcConnectionStringBuilder(dbConnectionString)
        Dim queries() As String = query.Split(";"c)
        connection = New Odbc.OdbcConnection(connectionString.ToString)

        Dim adapter As Odbc.OdbcDataAdapter
        Dim resultTable As New DataTable

        Log(String.Format("Multiline query - {0} lines in query", queries.Length()), 3, 17)

        Dim hasExceptionsCounter As Integer = 0
        Dim goodResultsCounter As Integer = 0

        For Each line As String In queries
            If line.Trim() <> "" Then
                Log(String.Format("Current Query = {0}", line), 10, 17)
                Log("Setting new OdbcDataAdapter for current query", 10, 17)
                adapter = New Odbc.OdbcDataAdapter(line, connection)
                Try
                    Log("Opening database object", 10, 17)
                    connection.Open()
                    Log("Filling result table", 10, 17)
                    adapter.Fill(resultTable)
                    Log("Loaded results into result table successfully", 10, 17)
                    goodResultsCounter += 1
                Catch ex As Exception
                    hasExceptionsCounter += 1
                    Log(String.Format("Error processing query. Exception: {0}", ex.Message), 1, 9017)
                Finally
                    Log("Closing database object", 10, 17)
                    adapter.Dispose()
                    connection.Close()
                End Try

            Else
                Log("Value is null for query. Will not be processes", 10, 17)
            End If
        Next
        Log("Disposing database object", 10, 17)
        connection.Dispose()


        If goodResultsCounter = 0 AndAlso hasExceptionsCounter > 0 Then
            Log(String.Format("No good results were returned ({0}), Exceptions were returned {1}. Returning null",
                              goodResultsCounter.ToString, hasExceptionsCounter.ToString), 1, 9017)
            Log("Exiting InterbaseMultilineQuery(query,tablename) Function (null)", 10, 17)
            Return ""
        End If

        If resultTable.Rows.Count > 0 Then
            Log(String.Format("{0} rows were returned in results", resultTable.Rows.Count), 10, 17)
            Dim count As Integer = resultTable.Rows(0).ItemArray.Count
            Log(String.Format("Item count for first row: {0}", count.ToString), 10, 17)
            Dim result As String = ""
            For counter As Integer = 0 To count - 1
                If counter = 0 Then
                    result = resultTable.Rows(0).Item(counter).ToString().Trim()
                Else
                    result = result & "," & resultTable.Rows(0).Item(counter).ToString().Trim()
                End If
            Next
            Log(String.Format("Exiting InterbaseMultilineQuery(query,tablename) Function ({0})", result), 10, 17)
            Return result
        Else
            Log("No rows were returned from query", 10, 17)
        End If
        Log("Exiting InterbaseMultilineQuery(query,tablename) Function (Null Results)", 10, 17)
        Return ""
    End Function

    Private Sub LaunchImport()
        Log("Entering LaunchImport() Sub", 10, 18)
        Dim id As Integer
        If Not ImportMustRun Then
            Log(String.Format("No files in {0} or {1}\RcvTbl - Import will not run.", RCV_FILE_PATH, RCV_FILE_PATH), 3, 18)
            Log("Exiting LaunchImport() Sub (No files to import)", 10, 18)
            Exit Sub
        End If

        'Make sure DbUpgrader isn't running (it shouldn't be)
        Dim retryAttempts As Integer = 0
        While ProcessRunning("DbUpgrader")
            If retryAttempts = mDbUpgraderRetry Then
                Log(String.Format("Error - DbUpgrader did not exit before max number of retry attempts. Import will not run.", mDbUpgraderDelay), 2, 9018)
                Log("Exiting LaunchImport() Sub (DBUpgrader did not exit)", 10, 18)
                Exit Sub
            End If
            If retryAttempts = 0 And File.Exists("ProcessFunctions.dll") Then
                Dim pf As ProcessFunctions.Functions = New ProcessFunctions.Functions
                Dim parentProcess As String = pf.GetParentByName("DbUpgrader")
                Log(String.Format("DbUpgrader parent process owner is {0}", parentProcess), 2, 18)
                pf = Nothing
            End If
            Log(String.Format("DbUpgrader currently running; waiting {0} second (Retry #{1})", mDbUpgraderDelay, retryAttempts), 2, 18)
            Sleep(mDbUpgraderDelay * 1000)
            retryAttempts += 1
        End While

        'Call the retalix import program to load HOST download files
        If Not File.Exists(ImportExe) Then
            Log("Error locating C:\Office\exe\Import.exe", 10, 9018)
            Log("Exiting LaunchImport() Sub (Import does not exist)", 10, 9018)
            EmailError(mOFBMailRecipient, "Import.exe missing", String.Format("{0} was not found. Please check site.", ImportExe))
            Exit Sub
        End If
        Checkfiles(RCV_FILE_PATH, mRcvDatFilesAgeWarningInDays)
        ProgressDialog.AddProgress("Calling Import.exe")
        Log("Call Retalix import started", 1, 18)
        mLogger.Flush()
        If ShowProgress Then
            Try
                id = Shell(ImportExe, AppWinStyle.MaximizedFocus, True)
            Catch ex As Exception
                Log(String.Format("Error launching {0}. Exception: ", ImportExe, ex.Message()), 3, 9018)
            End Try

        Else
            Try
                id = Shell(String.Format("{0} /M", ImportExe), AppWinStyle.Hide, True)
            Catch ex As Exception
                Log(String.Format("Error launching {0} /M. Exception: ", ImportExe, ex.Message()), 3, 9018)
            End Try

        End If
        ProgressDialog.AddProgress("Import.exe finished")
        Log("Call Retalix import finished", 1, 18)

        'Delay because import seems to take a second or two to actually quit
        Sleep(5000)
        Log("Exiting LaunchImport() Sub", 10, 18)
    End Sub

    Private Sub Checkfiles(ByVal path As String, ByVal Age As Integer)
        Log("Entering CheckFiles(path,age) Sub", 10, 19)
        Dim OldestDate As DateTime = GetOldestFileDateTime(path) 'GetOldestFileDateTime(path)
        Log(String.Format("Oldest date of file in {0} is {1}", path, OldestDate.ToString), 10, 19)
        Dim CurrDate As DateTime = DateTime.Now()
        Dim days As Long = DateDiff(DateInterval.Day, CurrDate, OldestDate)
        Log(String.Format("Difference in days for current date and oldest date = {0}", days), 10, 19)
        If Math.Abs(days) > Age Then
            Log(String.Format("Files found in {0} folder older than {0} days. Sending e-mail", path, Age), 1, 19)
            EmailError(mRcvDatAgeWarningMailRecipient, "File age warning",
                       String.Format("Files were found in {0} folder older than {1} days (Oldest was {2} days old). Please check the store for issues. Note: Files may now be in the most recent processed subfolder.", path, Age, Math.Abs(days)))
            Log("E-Mail sent (Old files)", 2, 9019)
        Else
            Log(String.Format("Date difference does not exceed {0} days. Nothing to do. ({1} > {2} = {3})",
                                      Age, Math.Abs(days), Age, Math.Abs(days) > Age), 10, 19)
        End If

        Log("Exiting CheckFiles(path,age)", 10, 19)
    End Sub

    Private Function MaintSrvRunning(ByVal LogIt As Boolean) As Boolean
        If LogIt Then Log("Entering MaintSrvRunning() Function (As Boolean)", 10, 20)
        Dim MaintSrvLog As String = String.Format("C:\Office\Log\MaintSrv-{0}.log", Date.Today.ToString("yyyyMMdd"))
        Dim result As Boolean = True
        If File.Exists(MaintSrvLog) Then
            If LogIt Then Log(String.Format("{0} was found. Attempting to parse", MaintSrvLog), 10, 9020)
            DbgMsg("MaintSrv log exists. Attempting to parse.")
            Try
                Dim reader As StreamReader = New StreamReader(MaintSrvLog)
                Dim line As String

                While Not reader.EndOfStream
                    line = reader.ReadLine().ToLower().Trim()
                    If line.Length > 10 Then
                        If line.IndexOf("task:") > 0 Then
                            If line.IndexOf("finished") > 0 Then
                                result = False
                            Else
                                result = True
                            End If
                        End If
                    End If
                End While
                reader.Close()
            Catch ex As Exception
                If LogIt Then Log(String.Format("Error parsing {0}. Exception: {1}", MaintSrvLog, ex.Message), 1, 9020)
            End Try
        Else
            If LogIt Then Log("Exiting MaintSrvRunning() Function (False)", 10, 20)
            result = False
        End If
        If LogIt Then Log(String.Format("Exiting MaintSrvRunning() Function ({0})", result), 10, 20)
        Return result
    End Function

    Private Function PendingPostSales() As Boolean
        Dim FlagTries As Integer = 1
        Dim hasPendingSales As Boolean = True
        Dim results As String = ""
        Dim query As String = "select count(Postsales) from daybatch where PostSales = 'F';"
        Dim iResult As Integer

        Log("Entering PendingPostSales() Function (As Boolean)", 10, 21)
        Log(String.Format("Running query: {0}", query), 10, 21)
        While hasPendingSales

            'Execute FlagCount.sql and view results. If other than 0, try mPostRetries times then send error msg
            ProgressDialog.AddProgress("Searching database for pending post sales, try #" & FlagTries)
            Log("Searching Database for pending post sales, Try #" & FlagTries, 1, 21)
            results = InterbaseQuery(query)
            If Not Integer.TryParse(results, iResult) Then
                iResult = 0
            End If


            If iResult = 0 Then
                ProgressDialog.AddProgress(String.Format("Database returned: {0}", iResult.ToString))
                Log(String.Format("Database returned: {0}", iResult.ToString), 2, 21)
                hasPendingSales = False
            Else
                ProgressDialog.AddProgress(String.Format("Database returned: {0}", iResult.ToString))
                ProgressDialog.AddProgress(String.Format("Sleeping for {0} seconds", mPostSalesSleep))
                Log(String.Format("Database returned: {0}", iResult.ToString), 3, 21)
                Log(String.Format("Sleeping for {0} seconds", mPostSalesSleep), 3, 21)
                Sleep(mPostSalesSleep * 1000)
            End If

            FlagTries += 1
            If FlagTries > mPostSalesRetry Then
                Log(String.Format("Pending post sales has remained {0} > 0 for greater than {1} tries. Exiting While Loop for hasPendingSales",
                                  iResult.ToString, mPostSalesRetry), 1, 9021)
                Exit While
            End If

        End While

        If hasPendingSales Then
            ProgressDialog.AddProgress("... exiting")
            ProgressDialog.Done()
            Log(String.Format("Database has returned {0} pending post sales after {1} attempts", iResult.ToString, mPostSalesRetry), 2, 9021)
            If mPendingPostSalesForceSetToTrue Then
                Try
                    Log("The set pending post sales forced to T was set. Attempting to set set.", 10, 21)
                    query = "Update daybatch set postsales = 'T'  where postsales = 'F';"
                    Log(String.Format("Query: {0}", query), 10, 21)
                    results = InterbaseQuery(query)
                    Log("Update pending post sales to 'T' succeeded.", 10, 21)
                    Log("Exiting PendingPostSales() Function (False)", 10, 21)
                    Return False

                Catch ex As Exception
                    Log(String.Format("Error forcing post sales to T: {0}", ex.Message), 1, 9021)
                    Log("Sending email (hasPendingSales)", 1, 9021)
                    EmailError(mMailRecipient, "Has Pending Sales", String.Format("Pending post sales records has remained > 0 with {0} records for {1} retry attempts, with a total delay of {2} seconds. Please check the site.",
                                                                                  iResult.ToString, mPostSalesRetry, mPostSalesRetry * mPostSalesSleep))
                    Log("E-Mail sent (hasPendingSales)", 1, 9021)
                    Log("Exiting PendingPostSales() Function (True)", 10, 9021)
                    Return True
                End Try
            Else
                Log("Sending email (hasPendingSales)", 1, 9021)
                EmailError(mMailRecipient, "Has Pending Sales", String.Format("Pending post sales records has remained > 0 with {0} records for {1} retry attempts, with a total delay of {2} seconds. Please check the site.",
                                                                              iResult.ToString, mPostSalesRetry, mPostSalesRetry * mPostSalesSleep))
                Log("E-Mail sent (hasPendingSales)", 1, 9021)
                Log("Exiting PendingPostSales() Function (True)", 10, 9021)
                Return True
            End If

        End If
        Log("No pending post sales records were found", 5, 21)
        Log("Exiting PendingPostSales() Function (False)", 10, 21)
        Return False
    End Function

    Private Sub ProcessASNFiles()
        Log("Entering ProcessASNFiles() Sub", 10, 22)
        Dim ASNFileArray(), ASNBackupArray(), ASNFile, line, rec As String
        Dim query As New Text.StringBuilder()
        Dim CommaCount As Integer = 0
        Dim OldNumOfCommas As Integer = 19
        Dim NewNumOfCommas As Integer = 20

        'Create ASN file array
        If Directory.Exists(ASNDataPath) Then
            Try
                ASNFileArray = Directory.GetFiles(ASNDataPath, ASNFileSpec)
            Catch ex As Exception
                Log(String.Format("Error retreiving file information from {0}. Exception: {1}", ASNDataPath, ex.Message), 3, 9022)
                ASNFileArray = Nothing
            End Try

        Else
            Try
                Directory.CreateDirectory(ASNDataPath)
            Catch ex As Exception
                Log(String.Format("Error creating {0}. Exception: ", ASNDataPath, ex.Message), 3, 9022)
            End Try

            ASNFileArray = Nothing
        End If

        If Not ASNFileArray Is Nothing AndAlso ASNFileArray.Length > 0 Then
            ProgressDialog.AddProgress("Processing ASN files")
            Log(String.Format("Found {0} ASN file(s)", ASNFileArray.Length), 3, 22)
            Log("Processing ASN files", 2, 22)

            ASNBackupArray = Directory.GetFiles(ASNDataPath, ASNFileSpec & ".bak")
            Try
                For Each ASNFile In ASNBackupArray
                    File.Delete(ASNFile)
                Next
                Log(String.Format("Deleted {0} backup files", ASNBackupArray.Length), 3, 22)
            Catch ex As Exception
                Log(String.Format("Error deleting ASN backup files. Exception: ", ex.Message), 3, 9022)
            End Try


            query.Append("delete from pilot_invoicestmp;")
            query.Append("commit;")

            'Process ASN files
            For Each ASNFile In ASNFileArray
                Log(String.Format("Found ASN input file {0}", ASNFile), 2, 22)
                Dim Counter As Integer = 0

                Dim Reader As StreamReader = New StreamReader(New FileStream(ASNFile, FileMode.Open, FileAccess.Read))

                While Not Reader.EndOfStream
                    rec = Reader.ReadLine()
                    Counter += 1
                    If CommaCount = 0 Then
                        CommaCount = CountCharOccurencesInStr(rec, ","c)
                    End If

                    'check if file is old type invoice or new type invoice (old type has OldNumOfCommas)
                    If CommaCount = OldNumOfCommas Then
                        line = "insert into pilot_invoicestmp (" & "ExtRef, LinkedExtRef, ExtCode, LocationId, DateRef, Lines, TotalQty, OrderCode, BarCode, Case_Qty, Case_Id, Case_Factor, Case_Cost, Tax_Val, Freight_Val,Discount_Val,Hdr_Tax,Hdr_Freight,Hdr_Disc,LineNbr,InvType)" & " values (" & rec & ",1);"
                        Log(line, 15, 22)
                    Else
                        line = "insert into pilot_invoicestmp (" & "ExtRef, LinkedExtRef, ExtCode, LocationId, DateRef, Lines, TotalQty, OrderCode, BarCode, Case_Qty, Case_Id, Case_Factor, Case_Cost, Tax_Val, Freight_Val,Discount_Val,Hdr_Tax,Hdr_Freight,Hdr_Disc,LineNbr,InvType)" & " values (" & rec & ");"
                        Log(line, 15, 22)
                    End If
                    query.Append(line)
                End While

                Log(String.Format("Processed {0} line(s)", Counter), 3, 22)

                Reader.Close()

                'Backup current ASN file
                Try
                    Log(String.Format("Moving {0} to {1}", ASNFile, ASNFile & ".bak"), 2, 22)
                    File.Move(ASNFile, ASNFile & ".bak")
                Catch ex As Exception
                    Log(String.Format("Error moving {0}. Exception: {1}", ASNFile, ex.Message), 2, 9022)
                    Log(String.Format("Found ASN input file {0}", ASNFile), 2, 22)
                End Try

            Next

            query.Append("commit;")
            query.Append("execute procedure pilot_insertinvoices;")
            query.Append("commit;")

            Log("ASNInsert.sql file created", 2, 22)

            ProgressDialog.AddProgress("ASN isql started")
            Log("Call ASN isql started", 2, 22)
            InterbaseMultilineQuery(query.ToString())
            Log("Call ASN isql ended", 3, 22)
        Else
            Log("No ASN Input files found", 2, 22)
        End If
        Log("Exiting ProcessASNFiles() Sub", 10, 22)
    End Sub

    Private Function GetOldestFileDateTime(ByVal MyDir As String) As DateTime
        Log("Entering GetOldestFileDateTime(MyDir) Function (As DateTime)", 10, 23)
        'Dim dOldestDate As Date

        'If Directory.Exists(MyDir) Then
        '    Log(String.Format("Checking for oldest date file in {0}", MyDir), 10)
        '    Try
        '        dOldestDate = New DirectoryInfo(MyDir).GetFiles().Max(Function(file As FileInfo) file.CreationTime)
        '        Log(String.Format("Oldest date is {0}", dOldestDate.ToString), 10)
        '    Catch ex As Exception
        '        Log("   Exiting GetOldestFileDateTime(MyDir) Function (Exception)", 10)
        '        Return Now()
        '    End Try

        'Else
        '    Log("   Exiting GetOldestFileDateTime(MyDir) Function (No Directory)", 10)
        '    Return Now()
        'End If
        Dim creationTime As DateTime = DateTime.Now
        Dim fn As String = ""
        If Directory.Exists(MyDir) Then
            Try
                Dim di As New DirectoryInfo(MyDir)
                Dim fiArr As FileInfo() = di.GetFiles()
                Dim Updated As Boolean = False
                For Each fi As FileInfo In fiArr
                    If DateTime.Compare(fi.LastWriteTime(), creationTime) < 0 Then
                        creationTime = fi.LastWriteTime()
                        fn = fi.FullName
                        Updated = True
                    End If
                Next
                If Updated = True Then
                    Log(String.Format("Oldest file is {0}", fn), 10, 23)
                    Log(String.Format("Oldest date is {0}", creationTime.ToString), 10, 23)
                End If
            Catch ex As Exception
                Log(String.Format("Error obtaining oldest file date from {0}. Exception: ", MyDir, ex.Message), 5, 9023)
            End Try
        Else
            Log(String.Format("{0} does not exist", MyDir), 5, 9023)
        End If


        Log("Exiting GetOldestFileDateTime(MyDir) Function", 10, 23)
        Return creationTime
    End Function

    'Private Function GetFileDateTimeByType(ByVal MyDir As String, ByVal Type As Integer, ByVal Ext As String) As DateTime
    '    Log("   Entering GetFileDateTimebyType(MyDir,Type,Ext) Function (As DateTime)", 10)
    '    Dim Result As DateTime
    '    Dim di As DirectoryInfo = New DirectoryInfo(MyDir)
    '    Dim fileInfos As IOrderedEnumerable(Of FileInfo) = di.EnumerateFiles("*.*", SearchOption.TopDirectoryOnly).Where(Function(n) Path.GetExtension(n.Name) = Ext).OrderBy(Function(n) n.CreationTime)
    '    Dim oldest As FileInfo = fileInfos.FirstOrDefault()
    '    Dim latest As FileInfo = fileInfos.LastOrDefault()
    '    If Type = 0 Then
    '        If oldest IsNot Nothing Then
    '            Result = oldest.CreationTime
    '        End If
    '    Else
    '        If latest IsNot Nothing Then
    '            Result = latest.CreationTime
    '        End If
    '    End If
    '    Log("   Exiting GetFileDateTimebyType(MyDir,Type,Ext) Function", 10)
    '    Return Result
    'End Function

    Private Sub ProcessDatFiles()
        Log("Entering ProcessDatFiles() Sub", 10, 24)
        Dim DatFileStillFound As Boolean
        Dim ImportLoopCount As Integer = 0
        Dim id As Integer

        Checkfiles(RCV_FILE_PATH, mRcvDatFilesAgeWarningInDays)

        DatFileStillFound = Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length > 0

        While DatFileStillFound
            ImportLoopCount += 1
            If ImportLoopCount > mImportExeMaxRetryAttempts Then
                Log(String.Format("Import has run {0} times and still has .dat files. Sending e-mail", mImportExeMaxRetryAttempts), 1, 9024)
                EmailError(mMailRecipient, "DAT files not processed completely", String.Format(IMPORT_X_TRIES, mImportExeMaxRetryAttempts))
                Log("E-Mail sent (Dat files not completely processed)", 2, 9024)
                DatFileStillFound = False
            Else
                ProgressDialog.AddProgress("Additional dat files found; running Import.exe again")
                Log("Additional dat files found; running import again", 1, 24)
                While ProcessRunning("DbUpgrader")
                    Log(String.Format("DbUpgrader currently running; waiting {0} seconds", mDbUpgraderDelay), 5, 24)
                    Sleep(mDbUpgraderDelay * 1000)
                End While

                'Call the retalix import program to load HOST download files
                ProgressDialog.AddProgress("Calling Import.exe")
                Log("Call Retalix import started", 2, 24)
                If File.Exists(ImportExe) Then
                    If ShowProgress Then
                        id = Shell(ImportExe, AppWinStyle.MaximizedFocus, True)
                    Else
                        id = Shell(String.Format("{0} /M", ImportExe), AppWinStyle.Hide, True)
                    End If
                    ProgressDialog.AddProgress("Import.exe finished")
                    Log("Call Retalix import ended", 2, 24)
                Else
                    Log(String.Format("{0} does not exit. Unable to call Import", ImportExe), 5, 9024)
                End If


                'Delay because import seems to take a second or two to actually quit
                Log(String.Format("Sleeping for {0} seconds after import exit to allow time to properly quit", mDelayAfterImport), 5, 24)
                Sleep(mDelayAfterImport * 1000)
                HasDatFiles = True
                Try
                    DatFileStillFound = Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length > 0
                    Log(String.Format("Reading dat files from {0} returns {1}", RCV_FILE_PATH, DatFileStillFound), 5, 24)
                Catch ex As Exception
                    Log(String.Format("Error reading files from {0}. Exception: ", RCV_FILE_PATH, ex.Message), 5, 9024)
                End Try

            End If
        End While
        Log("Exiting ProcessDatFiles() Sub", 10, 24)
    End Sub

    Private Sub ProcessShippers()
        Log("Entering ProcessShippers() Sub ", 10, 25)
        Dim query As New Text.StringBuilder()
        ProgressDialog.AddProgress("Processing shipper files")
        Log("Import.exe processed a shipper file", 2, 25)

        If File.Exists(ShipFile) Then
            Log("Found Pilot Shipper file " & ShipFile, 2, 25)

            Dim Reader As StreamReader = New StreamReader(New FileStream(ShipFile, FileMode.Open, FileAccess.Read))
            Dim rec As String

            query.Append("delete From pilot_shippers;")
            query.Append("commit;")

            While Not Reader.EndOfStream
                rec = Reader.ReadLine()
                query.Append(rec)
            End While

            query.Append("commit;")

            Reader.Close()
            Log("ShipInsert.sql file created (delete)", 3, 25)
            Try
                Log(String.Format("Moving {0} to {1}", ShipFile, ShipBak), 5, 25)
                File.Move(ShipFile, ShipBak)
            Catch ex As Exception
                Log(String.Format("Error backing up {0}. Exception: ", ShipFile, ex.Message), 1, 9025)
            End Try

        Else
            Log(String.Format("Shipper File {0} not found", ShipFile), 2, 25)
            Log("ShipInsert.sql file created (exec only)", 3, 25)
        End If

        query.Append("execute procedure pilot_updateshippers;")
        query.Append("commit;")

        Log("Call Shipper isql started", 2, 25)
        Log(String.Format("Executing {0} query", query.ToString), 15, 25)
        InterbaseMultilineQuery(query.ToString())
        Log("Call Shipper isql ended", 3, 25)
        Log("Exiting ProcessShippers() Sub ", 10, 25)
    End Sub

    Private Sub ReadCommandLineSwitches()
        Log("Entering ReadCommandLineSwitches() Sub", 10, 26)
        If Environment.GetCommandLineArgs().Length > 1 Then
            Dim Args As New ArrayList(Environment.GetCommandLineArgs())
            For Counter As Integer = 1 To Args.Count - 1
                Log(String.Format("Command line switch {0} = {1}", Counter, Args(Counter).ToString().ToLower()), 15, 25)
                If Args(Counter).ToString().ToLower() = "-noexevent" Then
                    DoExternalEvent = False
                    Log("Command line switch -noexevent found. Setting DoExternalEvent to false.", 3, 26)
                ElseIf Args(Counter).ToString().ToLower() = "-norpofiles" Then
                    DoRPOBuild = False
                    Log("Command line switch -norpofiles. Setting DoRPOBuild to false.", 3, 26)
                ElseIf Args(Counter).ToString().ToLower() = "-nso" Then
                    ShowProgress = True
                    ProgressDialog.Show()
                    Log("Command line switch -nso found. Setting ShowProgress to True.", 3, 26)
                Else
                    Log("Unknown command line switch was passed.", 3, 9026)
                End If
            Next
        End If
        Log("Exiting ReadCommandLineSwitches() Sub", 10, 26)
    End Sub

    Private Function RegistryFunction(ByVal hive As String, ByVal key As String, ByVal value As String, ByVal readValue As Boolean, ByVal result As String) As Boolean
        Dim baseKey As RegistryKey
        Log("Entering RegistryFunction(hive,key,value,readvalue,result) Function (As Boolean)", 10, 31)
        Log(String.Format("Hive = {0}", hive), 10, 31)
        Log(String.Format("Key = {0}", key), 10, 31)
        Log(String.Format("Value = {0}", value), 10, 31)
        Log(String.Format("ReadValue = {0}", readValue), 10, 31)
        Try
            Select Case hive
                Case "HKLM"
                    baseKey = Registry.LocalMachine.OpenSubKey(key, True)
                    If baseKey Is Nothing Then
                        Log(String.Format("{0}{1} did not exist. Creating key", hive, key), 1, 31)
                        baseKey = Registry.LocalMachine.CreateSubKey(key, RegistryKeyPermissionCheck.Default)
                    End If
                Case "HKCU"
                    baseKey = Registry.CurrentUser.OpenSubKey(key, True)
                    If baseKey Is Nothing Then
                        Log(String.Format("{0}{1} did not exist. Creating key", hive, key), 1, 31)
                        baseKey = Registry.CurrentUser.CreateSubKey(key, RegistryKeyPermissionCheck.Default)
                    End If
                Case Else
                    Log(String.Format("Unsupported registry hive ({0}} was passed", hive), 1, 9031)
                    Log("Exiting RegistryFunction(hive,key,value,readvalue,result) Function (False)", 10, 31)
                    Return False
            End Select

            If readValue Then
                If baseKey.GetValue(value) Is Nothing Then
                    Log(String.Format("{0}{1}[{2}] did not exist. Creating value and setting to {3}", hive, key, value, result), 1, 31)
                    baseKey.SetValue(value, result, RegistryValueKind.String)
                End If
            Else
                Log(String.Format("{0}{1}[{2}] exists. Setting to {3}", hive, key, value, result), 1, 31)
                baseKey.SetValue(value, result, RegistryValueKind.String)
            End If
            result = CStr(baseKey.GetValue(value))
            Log(String.Format("Result = {0}", result), 10, 31)
            baseKey.Close()
            Log("Exiting RegistryFunction(hive,key,value,readvalue,result) Function (True)", 10, 31)
            Return True
        Catch ex As Exception
            Log(String.Format("Error processing registry key. Exception : {0}", ex.Message), 1, 9031)
            Log("Exiting RegistryFunction(hive,key,value,readvalue,result) Function (False)", 10, 31)
            Return False
        End Try
    End Function

    Private Function GetRegistryValue(ByVal hive As String, ByVal key As String, ByVal value As String, ByVal DefaultResult As String) As String
        Dim baseKey As RegistryKey
        Dim result As String
        Log("Entering GetRegistryValue(hive,key,value,defaultresults) Function (As String)", 10, 31)
        Log(String.Format("Hive = {0}", hive), 10, 31)
        Log(String.Format("Key = {0}", key), 10, 31)
        Log(String.Format("Value = {0}", value), 10, 31)
        Log(String.Format("Default Result = {0}", DefaultResult), 10, 31)
        Try
            Select Case hive
                Case "HKLM"
                    baseKey = Registry.LocalMachine.OpenSubKey(key, True)
                    If baseKey Is Nothing Then
                        Log(String.Format("{0}{1} did not exist. Creating key", hive, key), 1, 31)
                        baseKey = Registry.LocalMachine.CreateSubKey(key, RegistryKeyPermissionCheck.Default)
                    End If
                Case "HKCU"
                    baseKey = Registry.CurrentUser.OpenSubKey(key, True)
                    If baseKey Is Nothing Then
                        Log(String.Format("{0}{1} did not exist. Creating key", hive, key), 1, 31)
                        baseKey = Registry.CurrentUser.CreateSubKey(key, RegistryKeyPermissionCheck.Default)
                    End If
                Case Else
                    Log(String.Format("Unsupported registry hive ({0}} was passed", hive), 1, 9031)
                    Log("Exiting RegistryFunction(hive,key,value,readvalue,result) Function (False)", 10, 31)
                    Return ""
            End Select


            If baseKey.GetValue(value) Is Nothing Then
                Log(String.Format("{0}{1}[{2}] did not exist. Creating value and setting to {3}", hive, key, value, DefaultResult), 1, 31)
                baseKey.SetValue(value, DefaultResult, RegistryValueKind.String)
                Log(String.Format("Exiting GetRegistryValue(hive,key,value,readvalue,result) Function ({0})", DefaultResult), 10, 31)
                Return DefaultResult
            End If

            result = CStr(baseKey.GetValue(value))
            Log(String.Format("Result = {0}", result), 10, 31)
            baseKey.Close()
            Log(String.Format("Exiting GetRegistryValue(hive,key,value,readvalue,result) Function ({0})", DefaultResult), 10, 31)
            Return result
        Catch ex As Exception
            Log(String.Format("Error processing registry key. Exception : {0}", ex.Message), 1, 9031)
            Log(String.Format("Exiting GetRegistryValue(hive,key,value,readvalue,result) Function ({0})", DefaultResult), 10, 31)
            Return DefaultResult
        End Try
    End Function

    Private Sub RPOPrepareData()
        Log("Entering RPOPrepareData() Sub", 10, 27)
        If (Date.Now.Hour >= mRPOBuildFilesMinHour AndAlso Date.Now.Hour < mRPOBuildFilesMaxHour) Then
            If DoRPOBuild Then
                ProgressDialog.AddProgress("Calling RPO Prepare Data")
                Log("Calling RPO Prepare Data", 3, 27)
                CallExternalEventRunner("RPOpreparedata")
                Log("RPO Prepare Data complted", 3, 27)
            Else
                Log("DoRpoBuild is set to false. Rpo files will not be built", 5, 27)
            End If
        Else
            Log(String.Format("{0} is outside of the run time for building RPO files. Min hour is {1}. Max hour is {2}. RPO files will not be built",
                                      Date.Now.Hour, mRPOBuildFilesMinHour, mRPOBuildFilesMaxHour), 5, 27)
        End If
        Log("Exiting RPOPrepareData() Sub", 10, 27)
    End Sub

    Private Function SetDatabasePath() As Boolean
        Log("Entering SetDatabasePath() Function As Boolean", 10, 28)
        'look for the back office database, if not found then close the app
        Log("Looking for database path", 1, 28)
        If File.Exists("d:\office\db\office.gdb") Then
            OfficeDB = "d:\office\db\office.gdb"
        ElseIf File.Exists("c:\office\db\office.gdb") Then
            OfficeDB = "c:\office\db\office.gdb"
        ElseIf File.Exists("c:\office\db\office.ib") Then
            OfficeDB = "c:\office\db\office.ib"
        Else
            ProgressDialog.AddProgress("Could not find the back office database")
            ProgressDialog.AddProgress("... exiting")
            ProgressDialog.Done()
            Log("Could not find the back office database", 2, 9028)
            Log("Exiting SetDatabasePath() Function As Boolean (False)", 10, 28)
            Return False
        End If

        'fix to get db connectivity in Win 7 and greater
        If WinVersionInfo.Version.Major >= 6 Then
            Log("WinVersionInfo major version returned is greater than 6. Appending localhost to DB path.", 4, 28)
            OfficeDB = "localhost:" & OfficeDB
        Else
            Log("WinVersionInfo major version returned is <= 6. localhost will not be appended to DB path.", 4, 28)
        End If
        Log(String.Format("Database found at {0}", OfficeDB), 4, 28)
        Log("Exiting SetDatabasePath() Function As Boolean (True)", 10, 28)
        Return True
    End Function

    Private Sub SetPaths()
        Log("Entering SetPaths() Sub", 10, 29)
        ASNFileSpec = "ASN*.txt"
        ShipFile = Path.Combine(ASNDataPath, "pship.txt")
        ShipBak = Path.Combine(ASNDataPath, "pshipbak.txt")
        Dim RcvTblPath As String = Path.Combine(RCV_FILE_PATH, "RcvTbl")

        If Not Directory.Exists(RcvTblPath) Then
            Try
                Log(String.Format("{0} does not exist. Creating directory.", RcvTblPath), 3, 29)
                Directory.CreateDirectory(Path.Combine(RCV_FILE_PATH, "RcvTbl"))
            Catch ex As Exception
                Log(String.Format("Error creating {0}. Exception: {1}", RcvTblPath, ex.Message), 1, 9029)
            End Try
        Else
            Log(String.Format("{0} already exists. Directory will not be created.", RcvTblPath), 11, 29)
        End If
        Log("Exiting SetPaths() Sub", 10, 29)
    End Sub

    Private Sub UpdateItemStatus()
        Log("Entering UpdateItemStatus() Sub", 10, 30)
        ProgressDialog.AddProgress("Calling Update Item Status")
        CallExternalEventRunner("updateitemstatus")
        FixIt()
        Log("Exiting UpdateItemStatus() Sub", 10, 30)
    End Sub

    Private Sub FixIt()
        Log("Entering FixIt() Sub", 10, 31)
        ProgressDialog.AddProgress("Calling Fixit procedure")
        Log("Call Fixit procedure started", 2, 31)
        'InterbaseQuery(FIXIT_PROCEDURE)
        InterbaseExecuteProcedure("fixit")
        Log("Call Fixit procedure ended", 3, 31)
        Log("Exiting FixIt() Sub", 10, 31)
    End Sub

    Private Function KillProcess(ByVal MyProcessName As String) As Boolean

        Dim psList() As Process
        Dim PerformKill As Boolean = True
        Try
            psList = Process.GetProcessesByName(GetFileName(MyProcessName))
            'psList = Process.GetProcesses()
            For Each p As Process In psList
                Log(String.Format("{0} process was running ({1}).", MyProcessName, p.Id), 5, 33)
                If myProcessId = p.Id Then
                    Log(String.Format("{0} process ID is current application ID. Cannot terminate self.", myProcessId), 1, 9033)
                    PerformKill = False
                End If
                If PerformKill Then
                    Log(String.Format("Terminating {0} ({1}).", MyProcessName, p.Id), 1, 33)
                    p.Kill()
                End If
            Next p

        Catch ex As Exception
            Log("KillProcess Error: " & ex.Message, 1, 9033)
        End Try
        Return Process.GetProcessesByName(GetFileName(MyProcessName)).Count > 0


    End Function

    Public Function GetFileName(ByVal filepath As String) As String

        Dim slashindex As Integer = filepath.LastIndexOf("\")
        Dim dotindex As Integer = filepath.LastIndexOf(".")

        GetFileName = filepath.Substring(slashindex + 1, dotindex - slashindex - 1)
    End Function

End Module

Module ProcessExtensions
    Private Function FindIndexedProcessName(ByVal pid As Integer) As String
        Dim processName As String = Process.GetProcessById(pid).ProcessName
        Dim processesByName As Process() = Process.GetProcessesByName(processName)
        Dim processIndexdName As String = Nothing

        For index As Integer = 0 To processesByName.Length - 1
            processIndexdName = If(index = 0, processName, processName & "#" & index)
            Dim processId As PerformanceCounter = New PerformanceCounter("Process", "ID Process", processIndexdName)

            If processId.NextValue().Equals(pid) Then
                Return processIndexdName
            End If
        Next

        Return processIndexdName
    End Function

    Private Function FindPidFromIndexedProcessName(ByVal indexedProcessName As String) As Process
        Dim parentId As PerformanceCounter = New PerformanceCounter("Process", "Creating Process ID", indexedProcessName)
        Return Process.GetProcessById(CInt(parentId.NextValue()))
    End Function

    Function Parent(ByVal process As Process) As Process
        Return FindPidFromIndexedProcessName(FindIndexedProcessName(process.Id))
    End Function


End Module

Class OdbcHandling
    Public Overloads Shared Function BuildParameterizedCommand(ByVal command As Odbc.OdbcCommand, ByVal Param As String,
                                           ByVal ParamValue As String) As Odbc.OdbcCommand

        command.Parameters.Add(Param, Odbc.OdbcType.VarChar, 255).Value = ParamValue

        Return command
    End Function

    Public Overloads Shared Function BuildParameterizedCommand(ByVal command As Odbc.OdbcCommand, ByVal Param As String,
                                               ByVal ParamValue As Integer) As Odbc.OdbcCommand

        command.Parameters.Add(Param, Odbc.OdbcType.Int).Value = ParamValue

        Return command
    End Function
End Class