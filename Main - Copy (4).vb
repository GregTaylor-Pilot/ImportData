Imports System.IO
Imports System.Xml
Imports System.Configuration
Imports Microsoft.Win32
Imports System.Data
Imports System.Collections.Generic

Module mdlMain

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Private StoreNumber, OfficeDB As String
    Private ISQLExe As String
    Private WinVersionInfo As System.OperatingSystem = System.Environment.OSVersion
    Private AppPath As String = My.Application.Info.DirectoryPath
    Private ScriptLog As String = Path.Combine(AppPath, My.Application.Info.AssemblyName & ".log")
    Private ErrorLog As String = Path.Combine(AppPath, My.Application.Info.AssemblyName & "_err.log")
    Private MailSender As String = "autopostfailure@pilottravelcenters.com"
    Private mRelayServer As String
    Private DoExternalEvent As Boolean = True
    Private DoRPOBuild As Boolean = True
    Private ShowProgress As Boolean = False
    Private HasDatFiles As Boolean = False
    Private ImportMustRun As Boolean = False
    Private LogDebugMessages As Boolean = False
    Private ASNFileSpec, ASNFileBackup, ASNInsert, ASNInsertLog As String
    Private FixItFile, FixItLog, FixCrossSiteFile, FixCrossSiteLog, PilotMaskingFile, PilotMaskingLog, PendingPricebatchesFile, PendingPricebatchesLog As String
    Private ShipFile, ShipSQL, ShipBak, ShipLog, ASNDataPath As String
    Private mFirstAlert, mSecondAlert, mExternalEventTimer, mLogLevel As Integer
    Private ISQLCommandTemplate As String
    Private mPendingPriceBatchesCheckTimes As New HashSet(Of Integer)

    Private mLogger As LogWriter
    Private mMailRecipient, mPricebatchMailRecipient As String

    Const SEPARATOR As String = " --------------------------------------------------"
    Const POST_SALES_QUERY As String = "select count(Postsales) from daybatch where PostSales = 'F';"
    Const FIX_CROSS_SITE_PROCEDURE As String = "execute procedure fix_cross_site;"
    Const PILOT_MASKING_PROCEDURE As String = "execute procedure pilot_masking;"
    Const PENDING_PRICE_BATCHES_QUERY As String = "select count(*) from pricemodhdr where status = 'P' and (valdate > ""today"" - 7) and (valdate <= ""today"");"
    Const FIXIT_PROCEDURE As String = "execute procedure fixit;"
    Const RCV_FILE_PATH As String = "C:\Office\Rcv"
    Const IMPORT_LOCKED_UP As String = "Import.exe is locked up at this location.  Please stop maintenance (Pos Transmitter), kill import.exe and importdata.exe, then restart maintenance and re-launch importdata.exe."
    Const IMPORT_5_TRIES As String = "Import has run 5 times and still has .dat files to be processed. Please check."

    Public Sub Main()
        Dim HasShippers As Boolean = False
        Dim HasPriceChangeFiles As Boolean = False
        Dim HasCategoryFile As Boolean
        Dim Progress As New ProgressDialog()

        'MaskPrompts()

        GetConfigOptions()

        StoreNumber = GetStoreNumber()

        If Not ImportDataRunning() Then

            If Not MaintSrvRunning() Then
                ReadCommandLineSwitches()

                Dim LaunchWriter As New System.IO.StreamWriter(New System.IO.FileStream(Path.Combine(AppPath, "LaunchTime.txt"), FileMode.Create, FileAccess.Write))
                LaunchWriter.WriteLine(Date.Now.Ticks)
                LaunchWriter.Close()

                SetPaths()
                RollLargeLog()

                mLogger = New LogWriter(ScriptLog, mLogLevel)
                ProgressDialog.AddProgress("ImportData started")
                mLogger.Log(" Script ImportData started", 1)
                If LogDebugMessages Then
                    mLogger.Log("    *** Debug info - Store number = " & StoreNumber, 5)
                    mLogger.Log("    *** Debug info - e-mail recipient = " & mMailRecipient, 5)
                End If

                HasDatFiles = Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length > 0
                HasPriceChangeFiles = Directory.GetFiles(RCV_FILE_PATH, "CR*.dat").Length > 0
                HasCategoryFile = Directory.GetFiles(RCV_FILE_PATH, "caty*.dat").Length > 0 OrElse Directory.GetFiles(RCV_FILE_PATH, "category.dat").Length > 0
                ImportMustRun = Directory.GetFiles(RCV_FILE_PATH, "*.*").Length > 0 OrElse Directory.GetFiles(Path.Combine(RCV_FILE_PATH, "RcvTbl"), "*.*").Length > 0

                If Not SetDatabasePath() Then
                    Exit Sub
                End If

                If DoExternalEvent Then
                    If File.Exists("C:\office\exe\externaleventrunner.exe") <> True Then
                        ProgressDialog.AddProgress("ExternalEventRunner.exe not found")
                        ProgressDialog.AddProgress("... exiting")
                        ProgressDialog.Done()
                        mLogger.Log("   ExternalEventRunner.exe was not found", 2)
                        mLogger.Log(SEPARATOR, 2)
                        mLogger.Close()
                        Exit Sub
                    End If
                End If

                Try
                    'CreateFlagCountSql()

                    HasShippers = CheckForShippers()

                    If PendingPostSales() Then
                        Exit Sub
                    End If

                    'AutoPostSales()

                    mLogger.Log(String.Format("   {0} dat file(s) found", Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length), 2)

                    If ProcessRunning("Import") Then
                        If ImportStillRunning() Then
                            Exit Sub
                        End If
                    Else
                        LaunchImport()
                    End If

                    'Call Update Item Status and Fixit
                    If HasDatFiles Then
                        UpdateItemStatus()
                        AutoPostSales()
                        AutoPriceSrv()
                    End If

                    ProcessDatFiles()

                    'Check if there are still any pending pricebatches
                    CheckForPendingPricebatches()

                    'Execute fix_cross_site procedure if we had a category file
                    If HasCategoryFile Then
                        CallCrossSiteFix()
                    End If

                    ProcessASNFiles()

                    'checks if store has a shipper file produced by HOST
                    If HasShippers Then
                        ProcessShippers()
                    End If

                    CallPilotMasking()

                    ProgressDialog.AddProgress("ImportData ended")
                    mLogger.Log(" ImportData End. Script ended", 1)
                    mLogger.Log(SEPARATOR, 1)

                    'Notify home office of price changes
                    If HasPriceChangeFiles Then
                        mLogger.Log(" Price change files detected.... sending notification to price change service", 2)
                        Try
                            Dim ws As New MessengerWebService.MessengerWebService()
                            Dim Result As Boolean = ws.PriceBookDownloadCompletedAlert(CInt(StoreNumber))
                            mLogger.Log("   Web service result: " & Result, 3)
                        Catch ex As Exception
                            mLogger.Log("   Error: " & ex.Message, 2)
                        End Try
                        mLogger.Log(SEPARATOR, 1)
                    End If

                    RPOPrepareData()

                    mLogger.Close()
                    ProgressDialog.AddProgress("... done")
                    ProgressDialog.Done()

                    CheckForFailureLog()

                    If Not MaintSrvRunning() Then
                        MaskPrompts()
                    End If

                Catch ex As Exception
                    EmailError(mMailRecipient, ex.Message)
                    mLogger.Log(" E-Mail sent", 1)

                    mLogger.Log(" ImportData ended", 1)
                    mLogger.Log(SEPARATOR, 1)
                    mLogger.Close()
                End Try
            Else
                mLogger = New LogWriter(ScriptLog)
                mLogger.Log(" Script ImportData launch attempted, but MaintSrv appears to be running", 1)
                mLogger.Log("   (No task finished line detected at end of MaintSrv log)", 1)
                mLogger.Log("   ImportData will now exit", 1)
                mLogger.Log(SEPARATOR, 1)
                mLogger.Close()
            End If
        End If
    End Sub

    Private Sub EmailError(ByVal recipients As String, ByVal body As String)
        Dim SendMail As New Net.Mail.SmtpClient(mRelayServer)
        Dim mailMessage As New Net.Mail.MailMessage()

        mailMessage.From = New Net.Mail.MailAddress(MailSender)
        For Each recipient As String In recipients.Split(";"c)
            If recipient.IndexOf("@") < 0 Then
                recipient = recipient.Trim() & "@pilottravelcenters.com"
            End If
            mailMessage.To.Add(recipient.Trim())
        Next
        mailMessage.Subject = String.Format("Store{0}", StoreNumber)
        mailMessage.Body = body
        SendMail.EnableSsl = True
        Try
            SendMail.Send(mailMessage)
        Catch ex As System.Net.Mail.SmtpException
            If ex.Message.IndexOf("does not support secure connections") > 0 Then
                SendMail.EnableSsl = False
                SendMail.Send(mailMessage)
            End If
        End Try
    End Sub

    Private Function ReadFile(ByRef FileName As String) As String
        Dim Reader As New System.IO.StreamReader(New System.IO.FileStream(FileName, FileMode.Open, FileAccess.Read))
        Dim NextLine As String
        Dim Text As New System.Text.StringBuilder(500)

        While Not Reader.EndOfStream
            NextLine = Reader.ReadLine()
            NextLine = NextLine & vbCrLf
            Text.Append(NextLine)
        End While

        Reader.Close()

        Return Text.ToString()
    End Function

    Function CountCharOccurencesInStr(ByRef sStringToSearch As String, ByRef sCharacter As Char) As Integer
        'Dim iOccurenceCount As Short
        'Dim sLoopString As String
        'Dim iPosition As Short

        'iOccurenceCount = 0
        'iPosition = 0
        'sLoopString = sStringToSearch

        'Do While InStr(sLoopString, sCharacter) <> 0
        '    iOccurenceCount = iOccurenceCount + 1
        '    iPosition = iPosition + InStr(sLoopString, sCharacter)
        '    sLoopString = Right(sLoopString, Len(sLoopString) - iPosition)
        '    iPosition = 0
        'Loop
        'CountCharOccurencesInStr = iOccurenceCount

        Return sStringToSearch.Split(sCharacter).Length - 1
    End Function

    Private Sub RollLargeLog()
        Dim logSize As Long
        If File.Exists(ScriptLog) Then
            Dim LogFileStreamRead As New FileStream(ScriptLog, FileMode.Open, FileAccess.Read)
            logSize = LogFileStreamRead.Length()
            LogFileStreamRead.Close()
        End If

        If logSize > (5 * 1024 * 1024) Then
            ' don't let the log file grow too large; delete any existing backup, rename the current to the backup
            File.Delete(ScriptLog & "_bak")
            File.Move(ScriptLog, ScriptLog & "_bak")
        End If
    End Sub

    Public Function ProcessRunning(ByRef pName As String) As Boolean
        Dim Processes() As Process = Process.GetProcessesByName(pName)

        Return Processes.Length > 0
    End Function

    Private Sub MaskPrompts()
        Dim configPath As String = Path.Combine(AppPath, "ImportData.exe.config")
        Dim config As XmlDocument = New XmlDocument()
        Dim template As String = "select pmntcode_id, pmsubcode_id from pmnt where pmnt_name like '{0}';"
        Dim selectTemplate As String = "select * from tillpmnt_prompt where sernum_tillpmnt in (select sernum from tillpmnt where pmcode = {0} and pmsubcode = {1}) and prompt_name = '{2}' and sernum_tillpmnt > {3};"
        Dim maskSQL As String = "update tillpmnt_prompt set prompt_response = {0} where sernum_tillpmnt in (select sernum from tillpmnt where pmcode = {1} and pmsubcode = {2}) and prompt_name = '{3}' and sernum_tillpmnt > {4};"
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
        'Dim reader As StreamReader

        mLogger = New LogWriter(ScriptLog)
        mLogger.Log(" Script ImportData masking started", 1)

        Try
            File.Delete(queryFile)
            config.XmlResolver = Nothing
            config.Load(configPath)
            For Each node As XmlNode In config.SelectNodes("/configuration/userSettings/ImportData.My.MySettings/setting")
                cardname = node.Attributes("name").InnerText
                mLogger.Log(" Getting payment data for " & cardname, 2)
                prompt = node.ChildNodes(0).InnerText

                If prompt.IndexOf(","c) > 0 Then
                    If IsNumeric(prompt.Substring(prompt.IndexOf(","c) + 1)) Then
                        partialMaskLength = CInt(prompt.Substring(prompt.IndexOf(","c) + 1))
                    Else
                        partialMaskLength = 10
                    End If
                    prompt = prompt.Substring(0, prompt.IndexOf(","c))
                    maskText = String.Format("substr(prompt_response, 1, {0}) || substr('XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX', 1, strlen(prompt_response) - {0})", partialMaskLength)
                Else
                    partialMaskLength = 0
                    maskText = "'MASKED'"
                End If

                writer = New StreamWriter(New FileStream(querySql, FileMode.Create, FileAccess.Write))
                writer.WriteLine(String.Format(template, cardname))
                writer.Close()

                ' results should be obtained in less than 1 second, but allow 10 seconds
                ' a blank file will cause the masking to be skipped
                'Shell(String.Format(ISQLCommandTemplate, querySql, queryFile), AppWinStyle.Hide, True, 10000)
                result = InterbaseQuery(String.Format(template, cardname), querySql, queryFile, 10000)

                minSequence = "0"

                If result <> "" Then
                    If result.IndexOf(",") > -1 AndAlso result.IndexOf(",") < 10 Then
                        pmntcode_id = result.Split(","c)(0)
                        pmsubcode_id = result.Split(","c)(1)
                    Else
                        pmntcode_id = result.Substring(0, 5).Trim()
                        pmsubcode_id = result.Substring(6).Trim()
                    End If
                    'mLogger.Log(String.Format("   pmntcode_id = {0}, pmsubcode_id = {1}", pmntcode_id, pmsubcode_id), 3)
                    If IsNumeric(pmntcode_id) Then

                        'mLogger.Log(String.Format("  Getting max(sernum) from tillpmnt where pmcode = {0} and pmsubcode = {1}", pmntcode_id, pmsubcode_id), 6)
                        writer = New StreamWriter(New FileStream(querySql, FileMode.Create, FileAccess.Write))
                        writer.WriteLine(String.Format(maxSeq, pmntcode_id, pmsubcode_id))
                        writer.Close()

                        maxSequence = InterbaseQuery(String.Format(maxSeq, pmntcode_id, pmsubcode_id), querySql, queryFile)

                        maxSequence = maxSequence.Trim()
                        'mLogger.Log(String.Format("   MaxQueryResult = '{0}'", maxSequence), 7)

                        If maxSequence <> "" AndAlso maxSequence <> "<null>" Then
                            If RegistryFunction("HKLM", "Software\PilotInfo\ImportData", String.Format("{0}_{1}", cardname.Trim(), prompt), True, minSequence) Then
                                mLogger.Log(String.Format("  Masking {0} prompt for {1}", prompt, cardname), 2)

                                writer = New StreamWriter(New FileStream(querySql, FileMode.Create, FileAccess.Write))
                                writer.WriteLine(String.Format(maskSQL, maskText, pmntcode_id, pmsubcode_id, prompt, minSequence))
                                writer.Close()

                                result = InterbaseQuery(String.Format(maskSQL, maskText, pmntcode_id, pmsubcode_id, prompt, minSequence), querySql, queryFile)

                                'mLogger.Log("   Done", 3)
                                RegistryFunction("HKLM", "Software\PilotInfo\ImportData", String.Format("{0}_{1}", cardname.Trim(), prompt), False, maxSequence)
                                Sleep(1000)
                            End If

                        End If
                    Else
                        mLogger.Log(String.Format("   -- Tender {0} not found", cardname), 4)
                    End If
                Else

                End If

            Next
            File.Delete(querySql)
            File.Delete(queryFile)
        Catch ex As Exception
            mLogger.Log(" ERROR: " & ex.Message, 1)
        End Try
        mLogger.Log(" ImportData masking done", 1)
        mLogger.Log(SEPARATOR, 1)
        mLogger.Close()
    End Sub

    Private Sub AutoPostSales()
        If DoExternalEvent Then
            ProgressDialog.AddProgress("Calling Auto Post Sales")
            CallExternalEventRunner("autopostsales")
        End If
    End Sub

    Private Sub AutoPriceSrv()
        ProgressDialog.AddProgress("Calling Auto Price Srv")
        CallExternalEventRunner("autopricesrv")
    End Sub

    Private Sub CallCrossSiteFix()
        If File.Exists(FixCrossSiteLog) Then
            File.Delete(FixCrossSiteLog)
            mLogger.Log("   Deleted FixCrossSite.log", 2)
        End If
        Dim Writer As StreamWriter = New StreamWriter(New FileStream(FixCrossSiteFile, FileMode.Create, FileAccess.Write, FileShare.None))
        Writer.WriteLine(FIX_CROSS_SITE_PROCEDURE)
        Writer.WriteLine("commit;")
        Writer.Close()
        ProgressDialog.AddProgress("Calling Fix_cross_site procedure")
        mLogger.Log("   Call fix_cross_site procedure started", 2)
        InterbaseQuery(FIX_CROSS_SITE_PROCEDURE, FixCrossSiteFile, FixCrossSiteLog)
        mLogger.Log("   Call fix_cross_site procedure ended", 3)
    End Sub

    Private Sub CallExternalEventRunner(ByVal eventName As String)
        Dim counter As Integer = 0
        Dim returnVal As Integer

        If Not ProcessRunning("Import") Then
            mLogger.Log("   Import not running - good", 5)
            If ProcessRunning("ExternalEventRunner") Then
                mLogger.Log("   *** ExternalEventRunner currently running", 2)
                While ProcessRunning("ExternalEventRunner") AndAlso counter < 5
                    Threading.Thread.Sleep(60000)
                    mLogger.Log("         Waited 1 minute", 4)
                    counter += 1
                End While
                If ProcessRunning("ExternalEventRunner") Then
                    mLogger.Log("       Waited 5 minutes for ExternalEventRunner to complete.... exiting", 3)
                    Return
                End If
            End If

            mLogger.Log(String.Format("   Call ExternalEventRunner.exe /{0}", eventName), 1)
            mLogger.Flush()

            returnVal = Shell("C:\office\exe\externaleventrunner.exe /" & eventName, AppWinStyle.Hide, True, mExternalEventTimer * 60000)

            If returnVal <> 0 Then
                mLogger.Log(String.Format("    ExternalEventRunner failed to complete in {0} minutes.... terminating process", mExternalEventTimer), 1)
                mLogger.Log(String.Format("     (ProcessId: {0})", returnVal), 2)
                Dim aProcess As System.Diagnostics.Process
                aProcess = System.Diagnostics.Process.GetProcessById(returnVal)
                aProcess.Kill()
            End If

            mLogger.Log("   Call ExternalEventRunner.exe ended", 2)
        Else
            mLogger.Log("   Skipping call to ExternalEventRunner due to Import running", 2)
        End If
    End Sub

    Private Sub CallPilotMasking()
        If File.Exists(PilotMaskingLog) Then
            File.Delete(PilotMaskingLog)
            mLogger.Log("   Deleted PilotMasking.log", 2)
        End If
        Dim Writer As StreamWriter = New StreamWriter(New FileStream(PilotMaskingFile, FileMode.Create, FileAccess.Write, FileShare.None))
        Writer.WriteLine(PILOT_MASKING_PROCEDURE)
        Writer.WriteLine("commit;")
        Writer.Close()
        ProgressDialog.AddProgress("Calling Pilot_masking procedure")
        mLogger.Log("   Call Pilot_masking procedure started", 2)
        InterbaseQuery(PILOT_MASKING_PROCEDURE, PilotMaskingFile, PilotMaskingLog)
        mLogger.Log("   Call Pilot_masking procedure ended", 3)
        File.Delete(PilotMaskingFile)
    End Sub

    Private Sub CheckForFailureLog()
        Dim FileName As String
        Dim dt As DateTime
        Dim counter As Integer = -3

        Do
            dt = DateAdd(DateInterval.Minute, counter, Now)
            FileName = Path.Combine("C:\Office\Log", String.Format("ImpFailure{0}.log", dt.ToString("yyyyMMddHHmm")))
            If File.Exists(FileName) Then
                'Dim SendMail As New Net.Mail.SmtpClient(mRelayServer)
                Dim Message As New System.Text.StringBuilder(100)
                Message.Append(FileName & " was found at store " & StoreNumber & "." & Chr(13))
                Message.Append(ReadFile(FileName))

                'SendMail.EnableSsl = True
                'SendMail.Send(MailSender, mMailRecipient, "STORE" & StoreNumber, Message.ToString())
                EmailError(mMailRecipient, Message.ToString())

                Exit Sub
            End If

            counter += 1
        Loop While counter <= 1
    End Sub

    Private Sub CheckForPendingPricebatches()
        If mPendingPriceBatchesCheckTimes.Contains(Date.Now.Hour) Then
            Dim results As String

            If File.Exists(PendingPricebatchesLog) Then
                File.Delete(PendingPricebatchesLog)
                mLogger.Log("   Deleted PendingPricebatches.log", 2)
            End If

            Dim Writer As StreamWriter = New StreamWriter(New FileStream(PendingPricebatchesFile, FileMode.Create, FileAccess.Write, FileShare.None))

            Writer.WriteLine(PENDING_PRICE_BATCHES_QUERY)
            Writer.Close()
            ProgressDialog.AddProgress("Checking for pending pricebatches")
            mLogger.Log("   Call pending pricebatches sql", 2)
            results = InterbaseQuery(PENDING_PRICE_BATCHES_QUERY, PendingPricebatchesFile, PendingPricebatchesLog)
            mLogger.Log("   Call pending pricebatches sql ended", 3)
            File.Delete(PendingPricebatchesFile)

            results = results.Trim(" "c)

            If CInt(results) <> 0 Then
                EmailError(mPricebatchMailRecipient, "This location has " & results & " pending pricebatch(es) for the last 7 days.")
            End If
        Else
            mLogger.Log("  Not scheduled to check for pending pricebatches this hour", 5)
        End If
    End Sub

    Private Function CheckForShippers() As Boolean
        ProgressDialog.AddProgress("Checking for shippers")
        mLogger.Log("   Checking for shipper file", 3)
        If File.Exists(RCV_FILE_PATH & "\SHeader.dat") Then
            mLogger.Log("   -- Shipper file found", 4)
            Return True
        Else
            mLogger.Log("   -- No shipper file found", 4)
            Return False
        End If
    End Function

    'Private Sub CreateFlagCountSql()
    '    Dim Writer As StreamWriter
    '    Dim query As String

    '    'Look for the SQL Script file, if it's there, delete and recreate it
    '    mLogger.Log("   Checking for FlagCount.sql file", 1)
    '    If File.Exists(SQLFile) Then
    '        mLogger.Log("   FlagCount.sql Found", 3)
    '        'File.Delete(SQLFile)
    '    Else
    '        Try
    '            Writer = New StreamWriter(New FileStream(SQLFile, FileMode.Append, FileAccess.Write, FileShare.None))
    '            query = "select count(Postsales) from daybatch where PostSales = 'F';"
    '            Writer.WriteLine(query)
    '            Writer.Close()
    '            mLogger.Log("   FlagCount.sql created", 4)
    '        Catch ex As Exception
    '            mLogger.Log("   Could not create the Flags.sql File!", 2)
    '            mLogger.Log("     Error Number: " & Err.Number, 2)
    '            mLogger.Log("     Error Description: " & Err.Description, 2)
    '            EmailError(mMailRecipient, "Could not create the Flags.sql file" & Chr(13) & "Error Description: " & Err.Description)
    '            mLogger.Log(" E-Mail sent", 1)

    '            mLogger.Log(" ImportData ended", 1)
    '            mLogger.Log(SEPARATOR, 1)
    '            mLogger.Close()
    '            Exit Sub
    '        End Try
    '    End If
    'End Sub

    Private Function GetConfigOption(ByVal optionName As String, ByVal defaultValue As String) As String
        Dim result As String
        If ConfigurationManager.AppSettings(optionName) Is Nothing Then
            result = defaultValue
        ElseIf ConfigurationManager.AppSettings(optionName).Trim() = "0" Then
            result = defaultValue
        Else
            result = ConfigurationManager.AppSettings(optionName)
        End If
        Return result
    End Function

    Private Sub GetConfigOptions()
        ISQLExe = GetConfigOption("ISQLPath", "c:\program files\borland\interbase\bin\isql.exe")
        mMailRecipient = GetConfigOption("EmailRecipient", "TicketTracker.HelpDesk@pilottravelcenters.com")
        mPricebatchMailRecipient = GetConfigOption("PendingPricebatchEmailRecipient", "eric.olsen@pilottravelcenters.com;daniel.king@pilottravelcenters.com;TicketTracker.HelpDesk@pilottravelcenters.com")
        For Each hour As String In GetConfigOption("PendingPricebatchEmailHours", "2,8,12").Split(","c)
            mPendingPriceBatchesCheckTimes.Add(CInt(hour))
        Next
        mRelayServer = GetConfigOption("RelayServer", "pilotrelay")
        LogDebugMessages = CBool(GetConfigOption("Debug", "False"))
        mFirstAlert = CInt(GetConfigOption("FirstAlert", "3"))
        mSecondAlert = CInt(GetConfigOption("NextAlert", "6"))
        mExternalEventTimer = CInt(GetConfigOption("ExternalEventTimer", "15"))
        mLogLevel = CInt(GetConfigOption("LogLevel", "5"))
        ASNDataPath = GetConfigOption("ASNDataPath", "C:\Pilot\data\")
    End Sub

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
            'Windows.Forms.MessageBox.Show("ImportData process already running" & ControlChars.CrLf & "Click OK to close", "ImportData already running", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Dim mLogger As New LogWriter(ErrorLog)
            mLogger.Log(SEPARATOR, 0)
            mLogger.Log(" Script ImportData start attempted, but failed due to process already running", 0)

            Dim LaunchReader As New System.IO.StreamReader(New System.IO.FileStream(Path.Combine(AppPath, "LaunchTime.txt"), FileMode.Open, FileAccess.Read))
            Dim Line As String = LaunchReader.ReadLine()
            LaunchReader.Close()
            Dim Elapsed As New TimeSpan(Date.Now.Ticks - CLng(Line))
            Dim minutes As Integer = (Elapsed.Days * 24 + Elapsed.Hours) * 60 + Elapsed.Minutes
            If (minutes > (mFirstAlert * 60) AndAlso minutes < ((mFirstAlert + 1) * 60)) OrElse (minutes > mSecondAlert * 60) Then
                mLogger.Log(String.Format("  ImportData has been running for {0} minutes. Sending alert e-mail", minutes), 0)
                EmailError(mMailRecipient, "ImportData has been running for at least three hours. Please check")
            End If
            mLogger.Log(SEPARATOR, 0)
            mLogger.Close()
            Return True
        End If
        Return False
    End Function

    Private Function ImportStillRunning() As Boolean
        ProgressDialog.AddProgress("Retalix Import is running")
        mLogger.Log("   Retalix Import is running", 3)
        Dim Processes() As Process = Process.GetProcessesByName("Import")
        Dim Times() As Int32 = {5, 10, 10}
        Dim counter As Integer = 0

        While counter < 3 And Not Processes(0).HasExited
            ProgressDialog.AddProgress(" Waiting " & Times(counter) & " minutes for Import to finish")
            mLogger.Log("   Waiting " & Times(counter) & " minutes for Import to finish", 3)
            mLogger.Flush()
            Sleep(Times(counter) * 60000)
            counter += 1
        End While

        If counter = 3 And Not Processes(0).HasExited Then
            'Dim SendMail As New Net.Mail.SmtpClient(mRelayServer)
            'SendMail.EnableSsl = True
            'SendMail.Send(MailSender, mMailRecipient, "STORE" & StoreNumber, IMPORT_LOCKED_UP)
            EmailError(mMailRecipient, IMPORT_LOCKED_UP)
            mLogger.Log("   Sent import failure e-mail", 1)
            mLogger.Log(SEPARATOR, 1)
            mLogger.Close()
            ProgressDialog.AddProgress(" Import.exe appears locked up")
            ProgressDialog.AddProgress("... exiting")
            ProgressDialog.Done()
            Return True
        End If

        Return False
    End Function

    Private Function InterbaseQuery(ByVal query As String, ByVal source As String, ByVal destination As String, Optional ByVal timeout As Integer = 0) As String
        Dim result As String

        Try
            result = InterbaseQuery(query)
            'mLogger.Log("     DB query result = " & result, 5)
        Catch ex As Exception
            result = InterbaseQuery(source, destination, timeout)
            'mLogger.Log("     ISQL query result = " & result, 5)
        End Try


        Return result
    End Function

    Private Function InterbaseQuery(ByVal source As String, ByVal destination As String, Optional ByVal timeout As Integer = 0) As String
        Dim reader As StreamReader
        Dim line, result As String

        If timeout = 0 Then
            Shell(String.Format(ISQLCommandTemplate, source, destination), AppWinStyle.Hide, True)
        Else
            Shell(String.Format(ISQLCommandTemplate, source, destination), AppWinStyle.Hide, True, timeout)
        End If

        reader = New StreamReader(New FileStream(destination, FileMode.Open, FileAccess.Read))
        result = ""

        While Not reader.EndOfStream
            line = reader.ReadLine()
            If line.Trim(" "c) <> "" Then
                result = line
            End If
        End While
        reader.Close()

        File.Delete(destination)

        Return result.Trim(" "c)
    End Function

    Private Function InterbaseQuery(ByVal query As String) As String
        Dim connection As New Odbc.OdbcConnection()
        Dim connectionString As New Odbc.OdbcConnectionStringBuilder("DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA;PWD=masterkey;DBNAME=" & OfficeDB & ";")
        connection = New Odbc.OdbcConnection(connectionString.ToString)

        Dim adapter As New Odbc.OdbcDataAdapter(query, connectionString.ToString)
        Dim resultTable As New DataTable

        'Try
        connection.Open()
        adapter.Fill(resultTable)
        connection.Close()
        connection.Dispose()
        'Catch ex As Exception
        '    Return ""
        'End Try

        If resultTable.Rows.Count > 0 Then
            Dim count As Integer = resultTable.Rows(0).ItemArray.Count
            Dim result As String = ""
            For counter As Integer = 0 To count - 1
                result = result & "," & resultTable.Rows(0).Item(counter).ToString().Trim()
            Next
            Return result.Substring(1)
        End If

        Return ""
    End Function

    Private Sub LaunchImport()
        If Not ImportMustRun Then
            mLogger.Log(String.Format("  No files in {0} or {0}\RcvTbl - Import will not run", RCV_FILE_PATH), 3)
            Exit Sub
        End If

        'Make sure DbUpgrader isn't running (it shouldn't be)
        While ProcessRunning("DbUpgrader")
            mLogger.Log("   DbUpgrader currently running; waiting 1 minute", 2)
            Sleep(60000)
        End While

        'Call the retalix import program to load HOST download files
        ProgressDialog.AddProgress("Calling Import.exe")
        mLogger.Log("   Call Retalix import started", 1)
        mLogger.Flush()
        If ShowProgress Then
            Shell("C:\Office\Exe\Import.exe", AppWinStyle.MaximizedFocus, True)
        Else
            Shell("C:\Office\Exe\Import.exe /M", AppWinStyle.Hide, True)
        End If
        ProgressDialog.AddProgress("Import.exe finished")
        mLogger.Log("   Call Retalix import ended", 1)

        'Delay because import seems to take a second or two to actually quit
        Sleep(5000)
    End Sub

    Private Function MaintSrvRunning() As Boolean
        Dim MaintSrvLog As String = String.Format("C:\Office\Log\MaintSrv-{0}.log", Date.Today.ToString("yyyyMMdd"))
        Dim result As Boolean = True
        If File.Exists(MaintSrvLog) Then
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

            End Try
        Else
            result = False
        End If
        Return result
    End Function

    Private Function PendingPostSales() As Boolean
        Dim FlagTries As Integer = 1
        Dim success As Boolean = False
        Dim hasPendingSales As Boolean = True
        Dim results As String = ""
        Dim countSQL As String = Path.Combine(AppPath, "FlagCount.sql")
        Dim countResult As String = Path.Combine(AppPath, "FlagResult.txt")
        Dim writer As New StreamWriter(New FileStream(countSQL, FileMode.Create, FileAccess.Write))

        writer.WriteLine(POST_SALES_QUERY)
        writer.Close()

        While hasPendingSales AndAlso FlagTries < 4

            If File.Exists(countResult) Then
                File.Delete(countResult)
            End If

            'Execute FlagCount.sql and view results. If other than 0, try 3 times then send error msg
            ProgressDialog.AddProgress("Searching database, try " & FlagTries)
            mLogger.Log("   Searching Database, Try " & FlagTries, 1)
            results = InterbaseQuery(POST_SALES_QUERY, countSQL, countResult)

            'If (results <> "0" AndAlso FlagTries = 3) Then
            '    ProgressDialog.AddProgress("Database returned: '" & results & "'")
            '    ProgressDialog.AddProgress("... exiting")
            '    ProgressDialog.Done()
            '    mLogger.Log("   Database returned: '" & results & "' after 3 attempts", 2)
            '    mLogger.Log("   Sending email and exiting", 2)
            '    EmailError(mMailRecipient, "Error searching database" & Chr(13) & "Database returned '" & results & "'")
            '    mLogger.Log(" E-Mail sent", 1)

            '    mLogger.Log(" ImportData ended", 1)
            '    mLogger.Log(SEPARATOR, 1)
            '    mLogger.Close()
            '    File.Delete(countResult)
            '    Return True
            'ElseIf (Not IsNumeric(results) OrElse CDbl(results) <> 0) AndAlso FlagTries <> 3 Then
            '    ProgressDialog.AddProgress("Database return: '" & results & "'")
            '    ProgressDialog.AddProgress("Sleeping for 3 minutes")
            '    mLogger.Log("   Database returned: '" & results & "'", 3)
            '    mLogger.Log("   Sleeping for 3 minutes", 3)
            '    Sleep(180000)
            'Else
            '    ProgressDialog.AddProgress("Database returned: '" & results & "'")
            '    mLogger.Log("   Database returned: '" & results & "'", 2)
            '    Success = True
            'End If

            If results = "0" Then
                ProgressDialog.AddProgress("Database returned: '" & results & "'")
                mLogger.Log("   Database returned: '" & results & "'", 2)
                hasPendingSales = False
            ElseIf FlagTries <> 3 Then
                ProgressDialog.AddProgress("Database returned: '" & results & "'")
                ProgressDialog.AddProgress("Sleeping for 3 minutes")
                mLogger.Log("   Database returned: '" & results & "'", 3)
                mLogger.Log("   Sleeping for 3 mintes", 3)
                Sleep(180000)
            End If

            FlagTries += 1

        End While

        If hasPendingSales Then
            ProgressDialog.AddProgress("Database returned: '" & results & "'")
            ProgressDialog.AddProgress("... exiting")
            ProgressDialog.Done()
            mLogger.Log("   Database returned: '" & results & "' after 3 attempts", 2)
            mLogger.Log("   Sending email and exiting", 2)
            EmailError(mMailRecipient, "Error searching database" & Chr(13) & "Database returned '" & results & "'")
            mLogger.Log(" E-Mail sent", 1)

            mLogger.Log(" ImportData ended", 1)
            mLogger.Log(SEPARATOR, 1)
            mLogger.Close()
            File.Delete(countResult)
            Return True
        End If

        File.Delete(countSQL)
        File.Delete(countResult)

        Return False
    End Function

    Private Sub ProcessASNFiles()
        Dim ASNFileArray(), ASNBackupArray(), ASNFile, line, rec As String
        Dim CommaCount As Integer = 0
        Dim OldNumOfCommas As Integer = 19
        Dim NewNumOfCommas As Integer = 20

        If File.Exists(ASNFileBackup) Then
            File.Delete(ASNFileBackup)
        End If

        'Start the invoice load process

        'Create ASN file array
        If Directory.Exists(ASNDataPath) Then
            ASNFileArray = Directory.GetFiles(ASNDataPath, ASNFileSpec)
        Else
            Directory.CreateDirectory(ASNDataPath)
            ASNFileArray = Nothing
        End If

        If Not ASNFileArray Is Nothing AndAlso ASNFileArray.Length > 0 Then
            ProgressDialog.AddProgress("Processing ASN files")
            mLogger.Log("   Found " & ASNFileArray.Length & " ASN file(s)", 3)
            mLogger.Log("   Processing ASN files", 2)
            'Perform cleanup
            If File.Exists(ASNInsertLog) Then
                File.Delete(ASNInsertLog)
                mLogger.Log("   Deleted ASNInsert.log", 2)
            End If

            If File.Exists(ASNInsert) Then
                File.Delete(ASNInsert)
                mLogger.Log("   Deleted ASNInsert.sql", 2)
            End If

            ASNBackupArray = Directory.GetFiles(ASNDataPath, ASNFileSpec & ".bak")
            For Each ASNFile In ASNBackupArray
                File.Delete(ASNFile)
            Next
            mLogger.Log("   Deleted " & ASNBackupArray.Length & " backup files", 3)

            Dim Writer As StreamWriter = New StreamWriter(New FileStream(ASNInsert, FileMode.Create, FileAccess.Write, FileShare.None))

            Writer.WriteLine("delete from pilot_invoicestmp;")
            Writer.WriteLine("commit;")

            'Process ASN files
            For Each ASNFile In ASNFileArray
                mLogger.Log("   Found input file " & ASNFile, 2)
                Dim Counter As Integer = 0

                Dim Reader As StreamReader = New StreamReader(New FileStream(ASNFile, FileMode.Open, FileAccess.Read))

                'build file for use with isql commmand for issuing sql inserts into the office database
                While Not Reader.EndOfStream
                    rec = Reader.ReadLine()
                    Counter += 1
                    If CommaCount = 0 Then
                        CommaCount = CountCharOccurencesInStr(rec, ","c)
                    End If

                    'check if file is old type invoice or new type invoice (old type has OldNumOfCommas)
                    If CommaCount = OldNumOfCommas Then
                        line = "insert into pilot_invoicestmp (" & "ExtRef, LinkedExtRef, ExtCode, LocationId, DateRef, Lines, TotalQty, OrderCode, BarCode, Case_Qty, Case_Id, Case_Factor, Case_Cost, Tax_Val, Freight_Val,Discount_Val,Hdr_Tax,Hdr_Freight,Hdr_Disc,LineNbr,InvType)" & " values (" & rec & ",1);"
                        Writer.WriteLine(line)
                    Else
                        line = "insert into pilot_invoicestmp (" & "ExtRef, LinkedExtRef, ExtCode, LocationId, DateRef, Lines, TotalQty, OrderCode, BarCode, Case_Qty, Case_Id, Case_Factor, Case_Cost, Tax_Val, Freight_Val,Discount_Val,Hdr_Tax,Hdr_Freight,Hdr_Disc,LineNbr,InvType)" & " values (" & rec & ");"
                        Writer.WriteLine(line)
                    End If
                End While

                mLogger.Log(String.Format("     Processed {0} line(s)", Counter), 3)

                Reader.Close()

                'Backup current ASN file
                File.Move(ASNFile, ASNFile & ".bak")
            Next

            Writer.WriteLine("commit;")
            Writer.WriteLine("execute procedure pilot_insertinvoices;")
            Writer.WriteLine("commit;")

            Writer.Close()

            mLogger.Log("   ASNInsert.sql file created", 2)

            Sleep(2500)

            ProgressDialog.AddProgress("ASN isql started")
            mLogger.Log("   Call ASN isql started", 2)
            InterbaseQuery(ASNInsert, ASNInsertLog)
            mLogger.Log("   Call ASN isql ended", 3)
        Else
            mLogger.Log("   No ASN Input files found", 2)
        End If
    End Sub

    Private Sub ProcessDatFiles()
        Dim DatFileStillFound As Boolean
        Dim ImportLoopCount As Integer = 0

        DatFileStillFound = Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length > 0

        While DatFileStillFound
            ImportLoopCount += 1
            If ImportLoopCount > 4 Then
                mLogger.Log("  Import has run 5 times and still has .dat files. Sending e-mail", 1)
                'Dim SendMail As New Net.Mail.SmtpClient(mRelayServer)
                'SendMail.EnableSsl = True
                'SendMail.Send(MailSender, mMailRecipient, "STORE" & StoreNumber, IMPORT_5_TRIES)
                EmailError(mMailRecipient, IMPORT_5_TRIES)
                mLogger.Log("  E-Mail sent", 2)
                DatFileStillFound = False
            Else
                ProgressDialog.AddProgress("Additional dat files found; running Import.exe again")
                mLogger.Log("   Additional dat files found; running import again", 1)
                While ProcessRunning("DbUpgrader")
                    mLogger.Log("   DbUpgrader currently running; waiting 1 minute", 1)
                    Sleep(60000)
                End While
                'mLogger.Log("   Renaming DbUpgrader to DbUpgrader_hold")
                'File.Move("C:\Office\Exe\DbUpgrader.exe", "C:\Office\Exe\DbUpgrader_hold.exe")

                'Call the retalix import program to load HOST download files
                ProgressDialog.AddProgress("Calling Import.exe")
                mLogger.Log("   Call Retalix import started", 2)
                If ShowProgress Then
                    Shell("C:\Office\Exe\Import.exe", AppWinStyle.MaximizedFocus, True)
                Else
                    Shell("C:\Office\Exe\Import.exe /M", AppWinStyle.Hide, True)
                End If
                ProgressDialog.AddProgress("Import.exe finished")
                mLogger.Log("   Call Retalix import ended", 2)
                'mLogger.Log("   Renaming DbUpgrader_hold to DbUpgrader")
                'File.Move("C:\Office\Exe\DbUpgrader_hold.exe", "C:\Office\Exe\DbUpgrader.exe")
                'Delay because import seems to take a second or two to actually quit
                Sleep(5000)
                HasDatFiles = True
                DatFileStillFound = Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length > 0
            End If
        End While
    End Sub

    Private Sub ProcessShippers()
        ProgressDialog.AddProgress("Processing shipper files")
        mLogger.Log("   Import.exe processed a shipper file", 2)

        If File.Exists(ShipLog) Then
            File.Delete(ShipLog)
            mLogger.Log("   Deleted existing ShipInsert log", 3)
        End If

        If File.Exists(ShipSQL) Then
            File.Delete(ShipSQL)
            mLogger.Log("   Deleted existing ShipInsert SQL file", 3)
        End If

        If File.Exists(ShipBak) Then
            File.Delete(ShipBak)
            mLogger.Log("   Deleted existing Shipper backup file", 3)
        End If

        If File.Exists(ShipFile) Then
            mLogger.Log("   Found Pilot Shipper file " & ShipFile, 2)

            Dim Reader As StreamReader = New StreamReader(New FileStream(ShipFile, FileMode.Open, FileAccess.Read))
            Dim Writer As StreamWriter = New StreamWriter(New FileStream(ShipSQL, FileMode.Create, FileAccess.Write, FileShare.None))
            Dim rec As String

            Writer.WriteLine("delete from pilot_shippers;")
            Writer.WriteLine("commit;")
            While Not Reader.EndOfStream
                rec = Reader.ReadLine()
                Writer.WriteLine(rec)
            End While

            Writer.WriteLine("commit;")
            Writer.WriteLine("execute procedure pilot_updateshippers;")
            Writer.WriteLine("commit;")

            Reader.Close()
            Writer.Close()

            mLogger.Log("   ShipInsert.sql file created", 3)

            File.Move(ShipFile, ShipBak)
        Else
            mLogger.Log("   Shipper File " & ShipFile & " not found", 2)
            mLogger.Log("   ShipInsert.sql file created (exec only)", 3)
            Dim Writer As StreamWriter = New StreamWriter(New FileStream(ShipSQL, FileMode.Create, FileAccess.Write, FileShare.None))
            Writer.WriteLine("execute procedure pilot_updateshippers;")
            Writer.WriteLine("commit;")
            Writer.Close()
        End If

        mLogger.Log("   Call Shipper isql started", 2)
        InterbaseQuery(ShipSQL, ShipLog)
        mLogger.Log("   Call Shipper isql ended", 3)
    End Sub

    Private Sub ReadCommandLineSwitches()
        If Environment.GetCommandLineArgs().Length > 1 Then
            Dim Args As New ArrayList(Environment.GetCommandLineArgs())
            For Counter As Integer = 1 To Args.Count - 1
                If Args(Counter).ToString().ToLower() = "-noexevent" Then
                    DoExternalEvent = False
                ElseIf Args(Counter).ToString().ToLower() = "-norpofiles" Then
                    DoRPOBuild = False
                ElseIf Args(Counter).ToString().ToLower() = "-nso" Then
                    ShowProgress = True
                    ProgressDialog.Show()
                End If
            Next
        End If

    End Sub

    Private Function RegistryFunction(ByVal hive As String, ByVal key As String, ByVal value As String, ByVal readValue As Boolean, ByVal result As String) As Boolean
        Dim baseKey As RegistryKey

        Try
            Select Case hive
                Case "HKLM"
                    baseKey = Registry.LocalMachine.OpenSubKey(key, True)
                    If baseKey Is Nothing Then
                        baseKey = Registry.LocalMachine.CreateSubKey(key, RegistryKeyPermissionCheck.Default)
                    End If
                Case "HKCU"
                    baseKey = Registry.CurrentUser.OpenSubKey(key, True)
                    If baseKey Is Nothing Then
                        baseKey = Registry.CurrentUser.CreateSubKey(key, RegistryKeyPermissionCheck.Default)
                    End If
                Case Else
                    Return False
            End Select

            If readValue Then
                If baseKey.GetValue(value) Is Nothing Then
                    baseKey.SetValue(value, result, RegistryValueKind.String)
                End If
            Else
                baseKey.SetValue(value, result, RegistryValueKind.String)
            End If
            result = CStr(baseKey.GetValue(value))
            baseKey.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub RPOPrepareData()
        If (Date.Now.Hour >= 0 AndAlso Date.Now.Hour < 3) And DoRPOBuild Then
            ProgressDialog.AddProgress("Calling RPO Prepare Data")
            CallExternalEventRunner("RPOpreparedata")
            mLogger.Log(SEPARATOR & Chr(13), 1)
        End If
    End Sub

    Private Function SetDatabasePath() As Boolean
        'look for the back office database, if not found then close the app
        mLogger.Log("   Looking for database", 1)
        If File.Exists("d:\office\db\office.gdb") Then
            OfficeDB = "d:\office\db\office.gdb"
        ElseIf File.Exists("c:\office\db\office.gdb") Then
            OfficeDB = "c:\office\db\office.gdb"
        Else
            ProgressDialog.AddProgress("Could not find the back office database")
            ProgressDialog.AddProgress("... exiting")
            ProgressDialog.Done()
            mLogger.Log("   Could not find the back office database", 2)
            mLogger.Log(" ImportData ended", 0)
            mLogger.Log(SEPARATOR, 0)
            mLogger.Close()
            Return False
        End If

        ' lazy fix to get db connectivity in Win 7
        If WinVersionInfo.Version.Major >= 6 Then
            OfficeDB = "localhost:" & OfficeDB
        End If
        mLogger.Log("   Database found at " & OfficeDB, 4)
        ISQLCommandTemplate = ISQLExe & " " & OfficeDB & " -i {0} -o {1} -e -m -u sysdba -p masterkey"
        Return True
    End Function

    Private Sub SetPaths()
        ASNFileSpec = "ASN*.txt"
        ASNFileBackup = Path.Combine(ASNDataPath, "ASNBak.txt")
        ASNInsert = Path.Combine(ASNDataPath, "ASNInsert.sql")
        ASNInsertLog = Path.Combine(ASNDataPath, "ASNInsert.log")
        'SQLFile = Path.Combine(AppPath, "FlagCount.sql")
        FixItFile = Path.Combine(AppPath, "Fixit.sql")
        FixItLog = Path.Combine(AppPath, "Fixit.log")
        FixCrossSiteFile = Path.Combine(AppPath, "FixCrossSite.sql")
        FixCrossSiteLog = Path.Combine(AppPath, "FixCrossSite.log")
        PilotMaskingFile = Path.Combine(AppPath, "PilotMasking.sql")
        PilotMaskingLog = Path.Combine(AppPath, "PilotMasking.log")
        PendingPricebatchesFile = Path.Combine(AppPath, "PendingPB.sql")
        PendingPricebatchesLog = Path.Combine(AppPath, "PendingPB.log")

        ShipFile = Path.Combine(ASNDataPath, "pship.txt")
        ShipSQL = Path.Combine(ASNDataPath, "ShipInsert.sql")
        ShipBak = Path.Combine(ASNDataPath, "pshipbak.txt")
        ShipLog = Path.Combine(ASNDataPath, "ShipperInsert.log")

        If Not Directory.Exists(Path.Combine(RCV_FILE_PATH, "RcvTbl")) Then
            Directory.CreateDirectory(Path.Combine(RCV_FILE_PATH, "RcvTbl"))
        End If
    End Sub

    'Private Function ShellRun(ByVal command As String, ByVal parameters As String, ByVal wait As Boolean, Optional ByVal waitMinutes As Integer = 0) As Integer
    '    Dim procID As Integer
    '    Dim newProc As Process
    '    Dim procInfo As New ProcessStartInfo(command)
    '    procInfo.Arguments = parameters
    '    procInfo.WindowStyle = ProcessWindowStyle.Hidden

    '    Try
    '        newProc = Process.Start(procInfo)
    '    Catch ex As Exception
    '        mLogger.Log("  **** ShellRun exception: " & ex.Message, 4)
    '        Return -1
    '    End Try
    '    procID = newProc.Id
    '    If wait Then
    '        newProc.WaitForExit()
    '        Dim procEC As Integer = -1
    '        If newProc.HasExited Then
    '            procEC = newProc.ExitCode
    '        End If

    '        mLogger.Log("  Process exit code: " & procEC, 4)
    '        Return procEC
    '    End If
    '    Return 0
    'End Function

    Private Sub UpdateItemStatus()
        ProgressDialog.AddProgress("Calling Update Item Status")
        CallExternalEventRunner("updateitemstatus")

        ' Create and exec file to call Fixit procedure
        If File.Exists(FixItLog) Then
            File.Delete(FixItLog)
            mLogger.Log("   Deleted Fixit.log", 2)
        End If
        Dim Writer As StreamWriter = New StreamWriter(New FileStream(FixItFile, FileMode.Create, FileAccess.Write, FileShare.None))
        Writer.WriteLine(FIXIT_PROCEDURE)
        Writer.WriteLine("commit;")
        Writer.Close()
        ProgressDialog.AddProgress("Calling Fixit procedure")
        mLogger.Log("   Call Fixit procedure started", 2)
        InterbaseQuery(FIXIT_PROCEDURE, FixItFile, FixItLog)
        mLogger.Log("   Call Fixit procedure ended", 3)
        File.Delete(FixItFile)
        File.Delete(FixItLog)
    End Sub
End Module