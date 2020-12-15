Imports System.IO
Imports System.Xml
Imports System.Configuration.ConfigurationSettings
Imports Microsoft.Win32

Module mdllMain

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Private StoreNumber, OfficeDB As String
    Private ISQLExe As String = "c:\program files\borland\interbase\bin\isql.exe "
    Private WinVersionInfo As System.OperatingSystem = System.Environment.OSVersion
    Private ScriptLog As String = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".log"

    Private LogWriter As System.IO.StreamWriter
    Private MailRecipient As String '= "magic@pilottravelcenters.com"

    Const SEPARATOR As String = " --------------------------------------------------"
    Const RCV_FILE_PATH As String = "C:\Office\rcv"
    Const DATA_PATH As String = "C:\Pilot\data\"

    Public Sub Main()
        'Const DataDir As String = "c:\pilot\data\"
        Dim ASNFileSpec As String
        Dim ASNFile As String
        Dim ASNFileArray() As String
        Dim ASNBackupArray() As String
        Dim ASNFileBackup As String
        Dim ASNInsert As String
        Dim ASNInsertLog As String
        Dim FixItFile As String
        Dim FixItLog As String
        Dim FixCrossSiteFile As String
        Dim FixCrossSiteLog As String
        Dim ShipFile As String
        Dim ShipSQL As String
        Dim ShipBak As String
        Dim ShipLog As String
        Dim SQLFile As String
        Dim HasShippers As Boolean = False
        Dim HasPriceChangeFiles As Boolean = False
        Dim idProg, iExit As Integer
        'Dim OfficeDB As String
        Dim Rec As String
        Dim NewLine As String
        Dim CommaCount As Integer
        Dim OldNumOfCommas As Short
        Dim NewNumOfCommas As Short
        'Dim Connection As String
        'Dim EmailFile As String
        Dim FlagResults As String = ""
        Dim FlagTries As Integer
        Dim Success As Boolean
        Dim Reader As StreamReader
        Dim Writer As StreamWriter
        Dim DoExternalEvent As Boolean = True
        Dim DoRPOBuild As Boolean = True
        Dim ShowProgress As Boolean = False
        Dim DatFileFound As Boolean
        Dim DatFileStillFound, HasCategoryFile As Boolean
        Dim Counter As Integer
        Dim Progress As New ProgressDialog()
        Dim ImportLoopCount As Integer
        Dim StoreNumberLength As Integer
        Dim LogDebugMessages As Boolean
        'Dim ShellProcess As Diagnostics.Process

        'MaskPrompts()

        MailRecipient = AppSettings("EmailRecipient")
        Try
            LogDebugMessages = CBool(AppSettings("Debug"))
        Catch ex As Exception
            LogDebugMessages = False
        End Try

        For Counter = 5 To 0 Step -1
            If IsNumeric(Environment.MachineName.Substring(0, Counter)) Then
                StoreNumberLength = Counter
                Exit For
            End If
        Next

        StoreNumber = Environment.MachineName.Substring(0, StoreNumberLength)

        If Environment.GetCommandLineArgs().Length > 1 Then
            Dim Args As New ArrayList(Environment.GetCommandLineArgs())
            For Counter = 1 To Args.Count - 1
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

        If Process.GetProcessesByName(My.Application.Info.AssemblyName).Length > 1 Then
            'Windows.Forms.MessageBox.Show("ImportData process already running" & ControlChars.CrLf & "Click OK to close", "ImportData already running", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            LogWriter = New System.IO.StreamWriter(New System.IO.FileStream(ScriptLog, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            LogWriter.AutoFlush = True
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR)
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " Script ImportData start attempted, but failed due to process already running")

            Dim LaunchReader As New System.IO.StreamReader(New System.IO.FileStream(My.Application.Info.DirectoryPath & "\LaunchTime.txt", FileMode.Open, FileAccess.Read))
            Dim Line As String = LaunchReader.ReadLine()
            LaunchReader.Close()
            Dim Elapsed As New TimeSpan(Date.Now.Ticks - CLng(Line))
            If Elapsed.Minutes > 120 Then
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  ImportData has been running for " & Elapsed.Minutes & " minutes. Sending alert e-mail")
                EmailError("ImportData has been running for at least two hours. Please check")
            End If
            'LogWriter.Close()
            Exit Sub
        End If

        Dim LaunchWriter As New System.IO.StreamWriter(New System.IO.FileStream(My.Application.Info.DirectoryPath & "\LaunchTime.txt", FileMode.Create, FileAccess.Write))
        LaunchWriter.WriteLine(Date.Now.Ticks)
        LaunchWriter.Close()

        ASNFileSpec = "ASN*.txt"
        ASNFileBackup = DATA_PATH & "ASNBak.txt"
        ASNInsert = DATA_PATH & "ASNInsert.sql"
        ASNInsertLog = DATA_PATH & "ASNInsert.log"
        SQLFile = My.Application.Info.DirectoryPath & "\FlagCount.sql"
        FixItFile = My.Application.Info.DirectoryPath & "\Fixit.sql"
        FixItLog = My.Application.Info.DirectoryPath & "\Fixit.log"
        FixCrossSiteFile = My.Application.Info.DirectoryPath & "\FixCrossSite.sql"
        FixCrossSiteLog = My.Application.Info.DirectoryPath & "\FixCrossSite.log"

        ShipFile = DATA_PATH & "pship.txt"
        ShipSQL = DATA_PATH & "ShipInsert.sql"
        ShipBak = DATA_PATH & "pshipbak.txt"
        ShipLog = DATA_PATH & "ShipperInsert.log"

        'Added to allow for credit invoices
        CommaCount = 0
        OldNumOfCommas = 19
        NewNumOfCommas = 20

        LogWriter = New System.IO.StreamWriter(New System.IO.FileStream(ScriptLog, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
        LogWriter.AutoFlush = True
        ProgressDialog.AddProgress("ImportData started")
        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " Script ImportData started")
        If LogDebugMessages Then
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "    *** Debug info - Store number = " & StoreNumber)
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "    *** Debug info - e-mail recipient = " & MailRecipient)
        End If

        DatFileFound = Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length > 0
        HasPriceChangeFiles = Directory.GetFiles(RCV_FILE_PATH, "CR*.dat").Length > 0
        HasCategoryFile = Directory.GetFiles(RCV_FILE_PATH, "caty*.dat").Length > 0 OrElse Directory.GetFiles(RCV_FILE_PATH, "category.dat").Length > 0


        Try
            'Look for the SQL Script file, if it's there, delete and recreate it
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Checking for FlagCount.sql file")
            If File.Exists(SQLFile) Then
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   FlagCount.sql Found")
                'File.Delete(SQLFile)
            Else
                Try
                    Writer = New StreamWriter(New FileStream(SQLFile, FileMode.Append, FileAccess.Write, FileShare.None))
                    NewLine = "select count(Postsales) from daybatch where PostSales = 'F';"
                    Writer.WriteLine(NewLine)
                    Writer.Close()
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   FlagCount.sql created")
                Catch ex As Exception
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Could not create the Flags.sql File!")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "     Error Number: " & Err.Number)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "     Error Description: " & Err.Description)
                    EmailError("Could not create the Flags.sql file" & Chr(13) & "Error Description: " & Err.Description)
                    'LogWriter.Close()
                    Exit Sub
                End Try
            End If

            'look for the back office database, if not found then close the app
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Looking for database")
            If File.Exists("d:\office\db\office.gdb") Then
                OfficeDB = "d:\office\db\office.gdb"
            ElseIf File.Exists("c:\office\db\office.gdb") Then
                OfficeDB = "c:\office\db\office.gdb"
            Else
                ProgressDialog.AddProgress("Could not find the back office database")
                ProgressDialog.AddProgress("... exiting")
                ProgressDialog.Done()
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Could not find the back office database")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " ImportData ended")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR)
                LogWriter.Close()
                Exit Sub
            End If
            ' lazy fix to get db connectivity in Win 7
            If WinVersionInfo.Version.Major = 6 Then
                OfficeDB = "localhost:" & OfficeDB
            End If
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Database found at " & OfficeDB)

            'Check for a shipper file
            ProgressDialog.AddProgress("Checking for shippers")
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Checking for shipper file")
            If File.Exists(RCV_FILE_PATH & "\SHeader.dat") Then
                HasShippers = True
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   -- Shipper file found")
            Else
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   -- No shipper file found")
            End If

            If DoExternalEvent Then
                If File.Exists("C:\office\exe\externaleventrunner.exe") <> True Then
                    ProgressDialog.AddProgress("ExternalEventRunner.exe not found")
                    ProgressDialog.AddProgress("... exiting")
                    ProgressDialog.Done()
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   ExternalEventRunner.exe was not found")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR)
                    LogWriter.Close()
                    Exit Sub
                End If
            End If

            FlagTries = 1
            Success = False

            While Not Success

                If File.Exists(My.Application.Info.DirectoryPath & "\FlagResult.txt") Then
                    File.Delete(My.Application.Info.DirectoryPath & "\FlagResult.txt")
                End If

                'Execute FlagCount.sql and view results. If other than 0, try 3 times then send error msg
                ProgressDialog.AddProgress("Searching database, try " & FlagTries)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Searching Database, Try " & FlagTries)
                'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss     ") & ISQLExe & OfficeDB & " -i c:\pilot\FlagCount.sql -m -o c:\pilot\FlagResult.txt -e -u sysdba -p masterkey")
                Shell(ISQLExe & OfficeDB & " -i c:\pilot\FlagCount.sql -m -o c:\pilot\FlagResult.txt -e -u sysdba -p masterkey", AppWinStyle.Hide, True)

                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Searching Database ended")

                Sleep(1000)
                Reader = New StreamReader(New FileStream(My.Application.Info.DirectoryPath & "\FlagResult.txt", FileMode.Open, FileAccess.Read))

                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   FlagResult.txt:")

                'Read the results file, saving the last non-blank line as our results
                While Not Reader.EndOfStream
                    NewLine = Reader.ReadLine()
                    LogWriter.WriteLine("                         " & NewLine)
                    If NewLine.Trim(" "c) <> "" Then
                        FlagResults = NewLine
                    End If
                End While
                Reader.Close()

                FlagResults = FlagResults.Trim(" "c)

                If (FlagResults <> "0" And FlagTries = 3) Then
                    ProgressDialog.AddProgress("Database returned: '" & FlagResults & "'")
                    ProgressDialog.AddProgress("... exiting")
                    ProgressDialog.Done()
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Database returned: '" & FlagResults & "' after 3 attempts")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Sending email and exiting")
                    EmailError("Error searching database" & Chr(13) & "Database retured '" & FlagResults & "'")
                    'LogWriter.Close()
                    Exit Sub
                ElseIf (IsNumeric(FlagResults) <> True Or CDbl(FlagResults) <> 0) And FlagTries <> 3 Then
                    ProgressDialog.AddProgress("Database return: '" & FlagResults & "'")
                    ProgressDialog.AddProgress("Sleeping for 3 minutes")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Database returned: '" & FlagResults & "'")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Sleeping for 3 minutes")
                    Sleep(180000)
                Else
                    ProgressDialog.AddProgress("Database returned: '" & FlagResults & "'")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Database returned: '" & FlagResults & "'")
                    Success = True
                End If

                FlagTries += 1

            End While

            If DoExternalEvent Then
                ProgressDialog.AddProgress("Calling Auto Post Sales")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call ExternalEventRunner.exe /autopostsales")
                'idProg = Shell("C:\office\exe\externaleventrunner.exe /autopostsales")
                'iExit = fWait(idProg)
                iExit = Shell("C:\office\exe\externaleventrunner.exe /autopostsales", AppWinStyle.Hide, True)

                'Exit program if we did not received a successful error code.
                'If (iExit <> 0) Then
                '    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  -ExternalEventRunner.exe exited with a return code of: " & iExit & Chr(13))
                '    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPERATOR & Chr(13))
                '    LogWriter.Close()
                '    Exit Sub
                'Else
                '    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  -End ExternalEventRunner.exe ended" & Chr(13))
                'End If

                'If we did not receive a successful error code, continue with program.
                If (iExit <> 0) Then
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call ExternalEventRunner.exe exited with a return code of '" & iExit & "'")
                Else
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call ExternalEventRunner.exe ended")
                End If
            End If

            LogWriter.WriteLine(String.Format("{0}   {1} dat file(s) found", Now.ToString("MM/dd/yyyy HH:mm:ss"), Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length))

            ' Check to see if the retalix import is already running; if so, wait for it to complete, else call retalix to load files
            If ProcessRunning("Import") Then
                ProgressDialog.AddProgress("Retalix Import already running")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Retalix Import already running")
                Dim Processes() As Process = Process.GetProcessesByName("Import")
                Dim Times() As Int32 = {5, 10, 10}
                Counter = 0
                While Counter < 3 And Not Processes(0).HasExited
                    ProgressDialog.AddProgress(" Waiting " & Times(Counter) & " minutes for Import to finish")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Waiting " & Times(Counter) & " minutes for Import to finish")
                    Sleep(Times(Counter) * 60000)
                    Counter += 1
                End While
                If Counter = 3 And Not Processes(0).HasExited Then
                    Dim SendMail As New System.Net.Mail.SmtpClient("relay1")
                    SendMail.Send("importfailure@pilottravelcenters.com", MailRecipient, Environment.MachineName & "Import error", "Import.exe is locked up at this location.  Please stop maintenance (Pos Transmitter), kill import.exe and importdata.exe, then restart maintenance and re-launch importdata.exe.")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Sent import failure e-mail")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR)
                    LogWriter.Close()
                    ProgressDialog.AddProgress(" Import.exe appears locked up")
                    ProgressDialog.AddProgress("... exiting")
                    ProgressDialog.Done()
                    Exit Sub
                End If
            Else
                'Make sure DbUpgrader isn't running (it shouldn't be), then rename it
                While ProcessRunning("DbUpgrader")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   DbUpgrader currently running; waiting 1 minute")
                    Sleep(60000)
                End While
                'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Renaming DbUpgrader to DbUpgrader_hold")
                'File.Move("C:\Office\Exe\DbUpgrader.exe", "C:\Office\Exe\DbUpgrader_hold.exe")
                'Call the retalix import program to load HOST download files
                ProgressDialog.AddProgress("Calling Import.exe")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Retalix import started")
                If ShowProgress Then
                    Shell("C:\Office\Exe\Import.exe", AppWinStyle.MaximizedFocus, True)
                Else
                    Shell("C:\Office\Exe\Import.exe /M", AppWinStyle.Hide, True)
                End If
                ProgressDialog.AddProgress("Import.exe finished")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Retalix import ended")
                'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Renaming DbUpgrader_hold to DbUpgrader")
                'File.Move("C:\Office\Exe\DbUpgrader_hold.exe", "C:\Office\Exe\DbUpgrader.exe")
            End If

            'Delay because import seems to take a second or two to actually quit
            Sleep(5000)

            DatFileStillFound = Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length > 0

            While DatFileStillFound
                ImportLoopCount += 1
                If ImportLoopCount > 4 Then
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  Import has run 5 times and still has .dat files. Sending e-mail")
                    Dim SendMail As New System.Net.Mail.SmtpClient("relay1")
                    SendMail.Send("autopostfailure@pilotcorp.com", MailRecipient, "STORE" & StoreNumber, "Import has run 5 times and still has .dat files to be processed. Please check.")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  E-Mail sent")
                    DatFileStillFound = False
                Else
                    ProgressDialog.AddProgress("Additional dat files found; running Import.exe again")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Additional dat files found; running import again")
                    While ProcessRunning("DbUpgrader")
                        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   DbUpgrader currently running; waiting 1 minute")
                        Sleep(60000)
                    End While
                    'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Renaming DbUpgrader to DbUpgrader_hold")
                    'File.Move("C:\Office\Exe\DbUpgrader.exe", "C:\Office\Exe\DbUpgrader_hold.exe")

                    'Call the retalix import program to load HOST download files
                    ProgressDialog.AddProgress("Calling Import.exe")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Retalix import started")
                    If ShowProgress Then
                        Shell("C:\Office\Exe\Import.exe", AppWinStyle.MaximizedFocus, True)
                    Else
                        Shell("C:\Office\Exe\Import.exe /M", AppWinStyle.Hide, True)
                    End If
                    ProgressDialog.AddProgress("Import.exe finished")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Retalix import ended")
                    'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Renaming DbUpgrader_hold to DbUpgrader")
                    'File.Move("C:\Office\Exe\DbUpgrader_hold.exe", "C:\Office\Exe\DbUpgrader.exe")
                    'Delay because import seems to take a second or two to actually quit
                    Sleep(5000)
                    DatFileFound = True
                    DatFileStillFound = Directory.GetFiles(RCV_FILE_PATH, "*.dat").Length > 0
                End If
            End While

            'Execute fix_cross_site procedure if we had a category file
            If HasCategoryFile Then
                If File.Exists(FixCrossSiteLog) Then
                    File.Delete(FixCrossSiteLog)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted FixCrossSite.log")
                End If
                Writer = New StreamWriter(New FileStream(FixCrossSiteFile, FileMode.Create, FileAccess.Write, FileShare.None))
                Writer.WriteLine("execute procedure fix_cross_site;")
                Writer.WriteLine("commit;")
                Writer.Close()
                ProgressDialog.AddProgress("Calling Fix_cross_site procedure")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call fix_cross_site procedure started")
                Shell(ISQLExe & OfficeDB & " -i " & FixCrossSiteFile & " -o " & FixCrossSiteLog & " -e -m -u sysdba -p masterkey", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call fix_cross_site procedure ended")
                File.Delete(FixCrossSiteLog)
            End If

            'Call Update Item Status and Fixit
            If DatFileFound Then
                ProgressDialog.AddProgress("Calling Update Item Status")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Update Item Status started")
                Shell("C:\office\exe\externaleventrunner.exe /updateitemstatus", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Update Item Status ended")

                ' Create and exec file to call Fixit procedure
                If File.Exists(FixItLog) Then
                    File.Delete(FixItLog)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted Fixit.log")
                End If
                Writer = New StreamWriter(New FileStream(FixItFile, FileMode.Create, FileAccess.Write, FileShare.None))
                Writer.WriteLine("execute procedure fixit;")
                Writer.WriteLine("commit;")
                Writer.Close()
                ProgressDialog.AddProgress("Calling Fixit procedure")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Fixit procedure started")
                Shell(ISQLExe & OfficeDB & " -i " & FixItFile & " -o " & FixItLog & " -e -m -u sysdba -p masterkey", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Fixit procedure ended")
                File.Delete(FixItFile)
            End If

            'Call Auto Price Srv
            ProgressDialog.AddProgress("Calling Auto Price Srv")
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Auto Price Srv started")
            Shell("C:\office\exe\externaleventrunner.exe /autopricesrv", AppWinStyle.Hide, True)
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Auto Price Srv ended")

            If File.Exists(ASNFileBackup) Then
                File.Delete(ASNFileBackup)
            End If

            'Start the invoice load process

            'Create ASN file array
            If Directory.Exists(DATA_PATH) Then
                ASNFileArray = Directory.GetFiles(DATA_PATH, ASNFileSpec)
            Else
                ASNFileArray = Nothing
            End If

            If Not ASNFileArray Is Nothing AndAlso ASNFileArray.Length > 0 Then
                ProgressDialog.AddProgress("Processing ASN files")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Found " & ASNFileArray.Length & " ASN file(s)")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Processing ASN files")
                'Peform cleanup
                If File.Exists(ASNInsertLog) Then
                    File.Delete(ASNInsertLog)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted ASNInsert.log")
                End If

                If File.Exists(ASNInsert) Then
                    File.Delete(ASNInsert)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted ASNInsert.sql")
                End If

                ASNBackupArray = Directory.GetFiles(DATA_PATH, ASNFileSpec & ".bak")
                For Each ASNFile In ASNBackupArray
                    File.Delete(ASNFile)
                Next
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted " & ASNBackupArray.Length & " backup files")

                Writer = New StreamWriter(New FileStream(ASNInsert, FileMode.Create, FileAccess.Write, FileShare.None))

                Writer.WriteLine("delete from pilot_invoicestmp;")
                Writer.WriteLine("commit;")

                'Process ASN files
                For Each ASNFile In ASNFileArray
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Found input file " & ASNFile)
                    Counter = 0

                    Reader = New StreamReader(New FileStream(ASNFile, FileMode.Open, FileAccess.Read))

                    'build file for use with isql commmand for issuing sql inserts into the office database
                    While Not Reader.EndOfStream
                        Rec = Reader.ReadLine()
                        Counter += 1
                        If CommaCount = 0 Then
                            CommaCount = CountCharOccurencesInStr(Rec, ","c)
                        End If

                        'check if file is old type invoice or new type invoice (old type has OldNumOfCommas)
                        If CommaCount = OldNumOfCommas Then
                            NewLine = "insert into pilot_invoicestmp (" & "ExtRef, LinkedExtRef, ExtCode, LocationId, DateRef, Lines, TotalQty, OrderCode, BarCode, Case_Qty, Case_Id, Case_Factor, Case_Cost, Tax_Val, Freight_Val,Discount_Val,Hdr_Tax,Hdr_Freight,Hdr_Disc,LineNbr,InvType)" & " values (" & Rec & ",1);"
                            Writer.WriteLine(NewLine)
                        Else
                            NewLine = "insert into pilot_invoicestmp (" & "ExtRef, LinkedExtRef, ExtCode, LocationId, DateRef, Lines, TotalQty, OrderCode, BarCode, Case_Qty, Case_Id, Case_Factor, Case_Cost, Tax_Val, Freight_Val,Discount_Val,Hdr_Tax,Hdr_Freight,Hdr_Disc,LineNbr,InvType)" & " values (" & Rec & ");"
                            Writer.WriteLine(NewLine)
                        End If
                    End While

                    LogWriter.WriteLine(String.Format("{0}     Processed {1} line(s)", Now.ToString("MM/dd/yyyy HH:mm:ss"), Counter))

                    Reader.Close()

                    'Backup current ASN file
                    File.Move(ASNFile, ASNFile & ".bak")
                Next

                Writer.WriteLine("commit;")
                Writer.WriteLine("execute procedure pilot_insertinvoices;")
                Writer.WriteLine("commit;")

                Writer.Close()

                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   ASNInsert.sql file created")

                Sleep(2500)

                ProgressDialog.AddProgress("ASN isql started")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call ASN isql started")

                Shell(ISQLExe & OfficeDB & " -i " & ASNInsert & " -o " & ASNInsertLog & " -e -m -u sysdba -p masterkey", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call ASN isql ended")
            Else
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   No ASN Input files found")
            End If

            'checks if store has a shipper file produced by HOST
            If HasShippers Then
                ProgressDialog.AddProgress("Processing shipper files")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Import.exe processed a shipper file")

                If File.Exists(ShipLog) Then
                    File.Delete(ShipLog)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted existing ShipInsert log")
                End If

                If File.Exists(ShipSQL) Then
                    File.Delete(ShipSQL)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted existing ShipInsert SQL file")
                End If

                If File.Exists(ShipBak) Then
                    File.Delete(ShipBak)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted existing Shipper backup file")
                End If

                If File.Exists(ShipFile) Then
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Found Pilot Shipper file " & ShipFile)

                    Reader = New StreamReader(New FileStream(ShipFile, FileMode.Open, FileAccess.Read))
                    Writer = New StreamWriter(New FileStream(ShipSQL, FileMode.Create, FileAccess.Write, FileShare.None))

                    Writer.WriteLine("delete from pilot_shippers;")
                    Writer.WriteLine("commit;")
                    While Not Reader.EndOfStream
                        Rec = Reader.ReadLine()
                        Writer.WriteLine(Rec)
                    End While

                    Writer.WriteLine("commit;")
                    Writer.WriteLine("execute procedure pilot_updateshippers;")
                    Writer.WriteLine("commit;")

                    Reader.Close()
                    Writer.Close()

                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   ShipInsert.sql file created")

                    File.Move(ShipFile, ShipBak)
                Else
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Shipper File " & ShipFile & " not found")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   ShipInsert.sql file created (exec only)")
                    Writer = New StreamWriter(New FileStream(ShipSQL, FileMode.Create, FileAccess.Write, FileShare.None))
                    Writer.WriteLine("execute procedure pilot_updateshippers;")
                    Writer.WriteLine("commit;")
                    Writer.Close()
                End If

                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Shipper isql started")
                Shell(ISQLExe & OfficeDB & " -i " & ShipSQL & " -o " & ShipLog & " -e -m -u sysdba -p masterkey ", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Shipper isql ended")
            End If

            ProgressDialog.AddProgress("ImportData ended")
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " ImportData End. Script ended")
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR)

            'Notify home office of price changes
            If HasPriceChangeFiles Then
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " Price change files detected.... sending notification to price change service" & Chr(13))
                Try
                    Dim ws As New MessengerWebService.MessengerWebService()
                    Dim Result As Boolean = ws.PriceBookDownloadCompletedAlert(CInt(StoreNumber))
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Web service result: " & Result)
                Catch ex As Exception
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Error: " & ex.Message)
                End Try
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR)
            End If

            'Call RPO Prepare Data
            If (Date.Now.Hour >= 0 AndAlso Date.Now.Hour < 3) And DoRPOBuild Then
                ProgressDialog.AddProgress("Calling RPO Prepare Data")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call RPO Prepare Data started")
                idProg = Shell("C:\office\exe\ExternalEventRunner.exe /RPOpreparedata", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call RPO Prepare Data ended")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR & Chr(13))
            End If

            'sanity check to make sure DbUpgrader_hold got renamed back to DbUpgrader
            'If File.Exists("C:\Office\Exe\DbUpgrader_hold.exe") Then
            '    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Rename of DbUpgrader_hold to DbUpgrader failed; attempting again")
            '    File.Move("C:\Office\Exe\DbUpgrader_hold.exe", "C:\Office\Exe\DbUpgrader.exe")
            'End If

            'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR)
            LogWriter.Close()
            ProgressDialog.AddProgress("... done")
            ProgressDialog.Done()

            MaskPrompts()

            Dim FileName As String
            Dim dt As DateTime
            Dim i As Int32 = -3

            Do
                dt = DateAdd(DateInterval.Minute, i, Now)
                FileName = "C:\Office\Log\ImpFailure" & dt.ToString("yyyyMMddHHmm") & ".log"
                If File.Exists(FileName) Then
                    Dim SendMail As New System.Net.Mail.SmtpClient("relay1")
                    Dim Message As New System.Text.StringBuilder(100)
                    Message.Append(FileName & " was found at store " & StoreNumber & "." & Chr(13))
                    Message.Append(ReadFile(FileName))

                    SendMail.Send("importfailure@pilotcorp.com", MailRecipient, "STORE" & StoreNumber & " Import Failure Log Found", Message.ToString())

                    Exit Sub
                End If

                i = i + 1
            Loop While i <= 1
        Catch ex As Exception
            EmailError(ex.Message)
        End Try
    End Sub

    Private Sub EmailError(ByVal body As String)
        Dim SendMail As New System.Net.Mail.SmtpClient("relay1")
        SendMail.Send("autopostfailure@pilotcorp.com", MailRecipient, "STORE" & StoreNumber, body)
        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " E-Mail sent")

        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " ImportData ended")
        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR)
        LogWriter.Close()
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

    'Function fWait(ByVal lProgID As Integer) As Integer
    '    ' Wait until proggie exit code <>
    '    '     STILL_ACTIVE&
    '    Dim lExitCode, hdlProg As Integer
    '    ' Get proggie handle
    '    hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, lProgID)
    '    ' Get current proggie exit code
    '    GetExitCodeProcess(hdlProg, lExitCode)

    '    Do While lExitCode = STILL_ACTIVE
    '        System.Windows.Forms.Application.DoEvents()
    '        GetExitCodeProcess(hdlProg, lExitCode)
    '    Loop
    '    CloseHandle(hdlProg)
    '    fWait = lExitCode
    'End Function

    Public Function ProcessRunning(ByRef pName As String) As Boolean
        'Dim hSnapshot, lRet As Integer
        'Dim P As New PROCESSENTRY32
        'Dim Found As Boolean

        'P.dwSize = Len(P)
        'hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0)
        'Found = False
        'Dim procName As String
        'If hSnapshot Then
        '    'UPGRADE_WARNING: Couldn't resolve default property of object P. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    lRet = Process32First(hSnapshot, P)
        '    Do While lRet
        '        procName = Left(P.szExeFile, InStr(P.szExeFile, Chr(0)) - 1)
        '        If LCase(procName) = LCase(pName) Then Found = True : Exit Do
        '        'UPGRADE_WARNING: Couldn't resolve default property of object P. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        lRet = Process32Next(hSnapshot, P)
        '    Loop
        '    lRet = CloseHandle(hSnapshot)
        'End If
        'ProcessRunning = Found
        Dim Processes() As Process = Process.GetProcessesByName(pName)

        Return Processes.Length > 0
    End Function

    Private Sub MaskPrompts()
        Dim configPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location()) & "\ImportData.exe.config"
        Dim config As XmlDocument = New XmlDocument()
        Dim template As String = "select pmntcode_id, pmsubcode_id from pmnt where pmnt_name like '{0}';"
        Dim selectTemplate As String = "select * from tillpmnt_prompt where sernum_tillpmnt in (select sernum from tillpmnt where pmcode = {0} and pmsubcode = {1}) and prompt_name = '{2}' and sernum_tillpmnt > {3};"
        Dim update As String = "update tillpmnt_prompt set prompt_response = 'MASKED' where sernum_tillpmnt in (select sernum from tillpmnt where pmcode = {0} and pmsubcode = {1}) and prompt_name = '{2}' and sernum_tillpmnt > {3};"
        Dim maxSeq As String = "select max(sernum) from tillpmnt where pmcode = {0} and pmsubcode = {1};"
        Dim pmntcode_id, pmsubcode_id As String
        Dim line, result As String
        'Dim office As String
        Dim cardname As String
        Dim prompt As String
        Dim minSequence, maxSequence As String
        Dim writer As StreamWriter
        Dim reader As StreamReader

        'If File.Exists("d:\office\db\office.gdb") Then
        '    office = "d:\office\db\office.gdb"
        'ElseIf File.Exists("c:\office\db\office.gdb") Then
        '    office = "c:\office\db\office.gdb"
        'Else
        '    Exit Sub
        'End If

        LogWriter = New System.IO.StreamWriter(New System.IO.FileStream(ScriptLog, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
        LogWriter.AutoFlush = True
        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " Script ImportData masking started")

        ' lazy fix to get db connectivity in Win 7
        'If WinVersionInfo.Version.Major = 6 Then
        '    office = "localhost:" & office
        'End If

        Try
            config.Load(configPath)
            For Each node As XmlNode In config.SelectNodes("/configuration/userSettings/ImportData.My.MySettings/setting")
                cardname = node.Attributes("name").InnerText
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " Getting payment data for " & cardname)
                prompt = node.ChildNodes(0).InnerText
                writer = New StreamWriter(New FileStream("C:\Pilot\CodeQuery.sql", FileMode.Create, FileAccess.Write))
                writer.WriteLine(String.Format(template, cardname))
                writer.Close()

                ' results should be obtained in less than 1 second, but allow 10 seconds
                ' a blank file will cause the masking to be skipped
                Shell(ISQLExe & officedb & " -i c:\pilot\CodeQuery.sql -m -o c:\pilot\QueryResult.txt -e -u sysdba -p masterkey", AppWinStyle.Hide, True, 10000)

                result = ""
                line = ""
                minSequence = "0"
                reader = New StreamReader(New FileStream("C:\Pilot\QueryResult.txt", FileMode.Open, FileAccess.Read))
                While Not reader.EndOfStream()
                    line = reader.ReadLine().Trim()
                    If line <> "" Then
                        result = line
                    End If
                End While
                reader.Close()
                'MessageBox.Show("line = " & result, "Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
                File.Delete("C:\Pilot\QueryResult.txt")
                If result <> "" Then
                    pmntcode_id = result.Substring(0, 5).Trim()
                    pmsubcode_id = result.Substring(6).Trim()

                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & String.Format("  Getting max(sernum) from tillpmnt where pmcode = {0} and pmsubcode = {1}", pmntcode_id, pmsubcode_id))
                    writer = New StreamWriter(New FileStream("C:\Pilot\CodeQuery.sql", FileMode.Create, FileAccess.Write))
                    writer.WriteLine(String.Format(maxSeq, pmntcode_id, pmsubcode_id))
                    writer.Close()

                    Shell(ISQLExe & OfficeDB & " -i c:\pilot\CodeQuery.sql -m -o c:\pilot\MaxQueryResult.txt -e -u sysdba -p masterkey", AppWinStyle.Hide, True)

                    reader = New StreamReader(New FileStream("C:\Pilot\MaxQueryResult.txt", FileMode.Open, FileAccess.Read,FileShare.ReadWrite))
                    While Not reader.EndOfStream()
                        line = reader.ReadLine().Trim()
                        If line <> "" Then
                            result = line
                        End If
                    End While
                    reader.Close()
                    File.Delete("C:\Pilot\MaxQueryResult.txt")

                    maxSequence = result.Trim()
                    'MessageBox.Show("maxSequence = " & maxSequence, "Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & String.Format("   MaxQueryResult = '{0}'", maxSequence))

                    If maxSequence <> "" Then
                        If RegistryFunction("HKLM", "Software\PilotInfo\ImportData", String.Format("{0}_{1}", cardname.Trim(), prompt), True, minSequence) Then
                            'writer = New StreamWriter(New FileStream("C:\Pilot\CodeQuery.sql", FileMode.Create, FileAccess.Write))
                            'writer.WriteLine(String.Format(selectTemplate, pmntcode_id, pmsubcode_id, prompt, minSequence))
                            'writer.Close()

                            'Shell(ISQLExe & office & " -i c:\pilot\CodeQuery.sql -m -o c:\pilot\PreUpdate.txt -e -u sysdba -p masterkey", AppWinStyle.Hide, True)

                            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & String.Format(" Masking {0} prompt for {1}", prompt, cardname))

                            writer = New StreamWriter(New FileStream("C:\Pilot\CodeQuery.sql", FileMode.Create, FileAccess.Write))
                            writer.WriteLine(String.Format(update, pmntcode_id, pmsubcode_id, prompt, minSequence))
                            writer.Close()

                            Shell(ISQLExe & OfficeDB & " -i c:\pilot\CodeQuery.sql -m -o c:\pilot\PostUpdate.txt -e -u sysdba -p masterkey", AppWinStyle.Hide, True)
                            File.Delete("C:\Pilot\PostUpdate.txt")

                            'writer = New StreamWriter(New FileStream("C:\Pilot\CodeQuery.sql", FileMode.Create, FileAccess.Write))
                            'writer.WriteLine(String.Format(selectTemplate, pmntcode_id, pmsubcode_id, prompt, minSequence))
                            'writer.Close()

                            'Shell(ISQLExe & office & " -i c:\pilot\CodeQuery.sql -m -o c:\pilot\PostUpdate.txt -e -u sysdba -p masterkey", AppWinStyle.Hide, True)

                            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & String.Format("   Done"))
                            RegistryFunction("HKLM", "Software\PilotInfo\ImportData", String.Format("{0}_{1}", cardname.Trim(), prompt), False, maxSequence)
                            Sleep(2000)
                        End If

                    End If
                Else

                End If

            Next
        Catch ex As Exception
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " ERROR: " & ex.Message)
            'MessageBox.Show("Error!!" & ControlChars.CrLf & ex.Message, "Oops", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " ImportData masking done")
        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR)
        LogWriter.Close()
        'File.Delete("C:\Pilot\PreUpdate.txt")
        File.Delete("C:\Pilot\PostUpdate.txt")
        File.Delete("C:\Pilot\CodeQuery.sql")
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
            'MessageBox.Show("baseKey = " & baseKey.ToString(), "Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
            If readValue Then
                'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  Reading registry string value." & Chr(13))
                'MessageBox.Show(String.Format("Reading '{0}' registry value", value), "Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If baseKey.GetValue(value) Is Nothing Then
                    baseKey.SetValue(value, result, RegistryValueKind.String)
                End If
                'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  Returned value is " & baseKey.GetValue(value).ToString() & Chr(13))
                'MessageBox.Show(String.Format("Returned value is '{0}'", baseKey.GetValue(value)), "Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  Writing registry string value." & Chr(13))
                'MessageBox.Show(String.Format("Writing '{0}' registry value", value), "Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
                baseKey.SetValue(value, result, RegistryValueKind.String)
            End If
            result = CStr(baseKey.GetValue(value))
            baseKey.Close()
            Return True
        Catch ex As Exception
            'MessageBox.Show("Error!!" & ControlChars.CrLf & ex.Message, "Oops", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
End Module