Imports System.IO

Module mdllMain

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Private StoreNumber As String
    Private LogWriter As System.IO.StreamWriter
    Const MailRecipient As String = "magic@pilottravelcenters.com"
    'Const MailRecipient As String = "lawsonp@pilottravelcenters.com"

    Const SEPARATOR As String = " --------------------------------------------------"

    Public Sub Main()
        Const DataDir As String = "c:\pilot\data\"
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
        Dim ScriptLog As String
        Dim ShipFile As String
        Dim ShipSQL As String
        Dim ShipBak As String
        Dim ShipLog As String
        Dim SQLFile As String
        Dim HaveShippers As Boolean = False
        Dim idProg, iExit As Integer
        Dim OfficeDB As String
        Dim Rec As String
        Dim NewLine As String
        Dim ISQLExe As String
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
        Dim Counter As Int32
        Dim Progress As New ProgressDialog()
        Dim ImportLoopCount As Int32
        'Dim ShellProcess As Diagnostics.Process

        StoreNumber = Environment.MachineName.Substring(0, 3)
        ScriptLog = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".log"

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
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR & Chr(13))
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " Script ImportData start attempted, but failed due to process already running" & Chr(13))

            Dim LaunchReader As New System.IO.StreamReader(New System.IO.FileStream(My.Application.Info.DirectoryPath & "\LaunchTime.txt", FileMode.Open, FileAccess.Read))
            Dim Line As String = LaunchReader.ReadLine()
            LaunchReader.Close()
            Dim Elapsed As New TimeSpan(Date.Now.Ticks - CLng(Line))
            If Elapsed.Minutes > 120 Then
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  ImportData has been running for " & Elapsed.Minutes & " minutes. Sending alert e-mail" & Chr(13))
                EmailError("ImportData has been running for at least two hours. Please check" & Chr(13))
            End If
            'LogWriter.Close()
            Exit Sub
        End If

        Dim LaunchWriter As New System.IO.StreamWriter(New System.IO.FileStream(My.Application.Info.DirectoryPath & "\LaunchTime.txt", FileMode.Create, FileAccess.Write))
        LaunchWriter.WriteLine(Date.Now.Ticks)
        LaunchWriter.Close()

        ISQLExe = "c:\program files\borland\interbase\bin\isql.exe "

        ASNFileSpec = "ASN*.txt"
        ASNFileBackup = DataDir & "ASNBak.txt"
        ASNInsert = DataDir & "ASNInsert.sql"
        ASNInsertLog = DataDir & "ASNInsert.log"
        SQLFile = My.Application.Info.DirectoryPath & "\FlagCount.sql"
        FixItFile = My.Application.Info.DirectoryPath & "\Fixit.sql"
        FixItLog = My.Application.Info.DirectoryPath & "\Fixit.log"
        FixCrossSiteFile = My.Application.Info.DirectoryPath & "\FixCrossSite.sql"
        FixCrossSiteLog = My.Application.Info.DirectoryPath & "\FixCrossSite.log"

        ShipFile = DataDir & "pship.txt"
        ShipSQL = DataDir & "ShipInsert.sql"
        ShipBak = DataDir & "pshipbak.txt"
        ShipLog = DataDir & "ShipperInsert.log"

        'Added to allow for credit invoices
        CommaCount = 0
        OldNumOfCommas = 19
        NewNumOfCommas = 20

        LogWriter = New System.IO.StreamWriter(New System.IO.FileStream(ScriptLog, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
        LogWriter.AutoFlush = True
        ProgressDialog.AddProgress("ImportData started")
        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " Script ImportData started" & Chr(13))

        DatFileFound = Directory.GetFiles("c:\office\rcv", "*.dat").Length > 0
        HasCategoryFile = Directory.GetFiles("c:\office\rcv", "caty*.dat").Length > 0 OrElse Directory.GetFiles("c:\office\rcv", "category.dat").Length > 0

        Try
            'Look for the SQL Script file, if it's there, delete and recreate it
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Checking for FlagCount.sql file" & Chr(13))
            If File.Exists(SQLFile) Then
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   FlagCount.sql Found" & Chr(13))
                'File.Delete(SQLFile)
            Else
                Try
                    Writer = New StreamWriter(New FileStream(SQLFile, FileMode.Append, FileAccess.Write, FileShare.None))
                    NewLine = "select count(Postsales) from daybatch where PostSales = 'F';"
                    Writer.WriteLine(NewLine)
                    Writer.Close()
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   FlagCount.sql created" & Chr(13))
                Catch ex As Exception
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Could not create the Flags.sql File!" & Chr(13))
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "     Error Number: " & Err.Number & Chr(13))
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "     Error Description: " & Err.Description & Chr(13))
                    EmailError("Could not create the Flags.sql file" & Chr(13) & "Error Description: " & Err.Description & Chr(13))
                    'LogWriter.Close()
                    Exit Sub
                End Try
            End If

            'look for the back office database, if not found then close the app
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Looking for database" & Chr(13))
            If File.Exists("d:\office\db\office.gdb") Then
                OfficeDB = "d:\office\db\office.gdb"
            ElseIf File.Exists("c:\office\db\office.gdb") Then
                OfficeDB = "c:\office\db\office.gdb"
            Else
                ProgressDialog.AddProgress("Could not find the back office database")
                ProgressDialog.AddProgress("... exiting")
                ProgressDialog.Done()
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Could not find the back office database" & Chr(13))
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " ImportData ended" & Chr(13))
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR & Chr(13))
                LogWriter.Close()
                Exit Sub
            End If
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Database found at " & OfficeDB & Chr(13))

            'Check for a shipper file
            ProgressDialog.AddProgress("Checking for shippers")
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Checking for shipper file" & Chr(13))
            If File.Exists("C:\Office\Rcv\SHeader.dat") Then
                HaveShippers = True
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   -- Shipper file found" & Chr(13))
            Else
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   -- No shipper file found" & Chr(13))
            End If

            If DoExternalEvent Then
                If File.Exists("C:\office\exe\externaleventrunner.exe") <> True Then
                    ProgressDialog.AddProgress("ExternalEventRunner.exe not found")
                    ProgressDialog.AddProgress("... exiting")
                    ProgressDialog.Done()
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   ExternalEventRunner.exe was not found" & Chr(13))
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR & Chr(13))
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
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Searching Database, Try " & FlagTries & Chr(13))
                Shell(ISQLExe & OfficeDB & " -i c:\pilot\FlagCount.sql -m -o c:\pilot\FlagResult.txt -e -u sysdba -p masterkey", AppWinStyle.Hide, True)

                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Searching Database ended" & Chr(13))

                Sleep(1000)
                Reader = New StreamReader(New FileStream(My.Application.Info.DirectoryPath & "\FlagResult.txt", FileMode.Open, FileAccess.Read))

                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   FlagResult.txt:" & Chr(13))

                'Read the results file, saving the last non-blank line as our results
                While Not Reader.EndOfStream
                    NewLine = Reader.ReadLine()
                    LogWriter.WriteLine("                         " & NewLine & Chr(13))
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
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Database returned: '" & FlagResults & "' after 3 attempts" & Chr(13))
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Sending email and exiting" & Chr(13))
                    EmailError("Error searching database" & Chr(13) & "Database retured '" & FlagResults & "'" & Chr(13))
                    'LogWriter.Close()
                    Exit Sub
                ElseIf (IsNumeric(FlagResults) <> True Or CDbl(FlagResults) <> 0) And FlagTries <> 3 Then
                    ProgressDialog.AddProgress("Database return: '" & FlagResults & "'")
                    ProgressDialog.AddProgress("Sleeping for 3 minutes")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Database returned: '" & FlagResults & "'" & Chr(13))
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Sleeping for 3 minutes")
                    Sleep(180000)
                Else
                    ProgressDialog.AddProgress("Database returned: '" & FlagResults & "'")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Database returned: '" & FlagResults & "'" & Chr(13))
                    Success = True
                End If

                FlagTries += 1

            End While

            If DoExternalEvent Then
                ProgressDialog.AddProgress("Calling Auto Post Sales")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call ExternalEventRunner.exe /autopostsales" & Chr(13))
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
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call ExternalEventRunner.exe exited with a return code of '" & iExit & "'" & Chr(13))
                Else
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call ExternalEventRunner.exe ended" & Chr(13))
                End If
            End If

            LogWriter.WriteLine(String.Format("{0}   {1} dat file(s) found", Now.ToString("MM/dd/yyyy HH:mm:ss"), Directory.GetFiles("c:\office\rcv\", "*.dat").Length))

            ' Check to see if the retalix import is already running; if so, wait for it to complete, else call retalix to load files
            If ProcessRunning("Import") Then
                ProgressDialog.AddProgress("Retalix Import already running")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Retalix Import already running" & Chr(13))
                Dim Processes() As Process = Process.GetProcessesByName("Import")
                Dim Times() As Int32 = {5, 10, 10}
                Counter = 0
                While Counter < 3 And Not Processes(0).HasExited
                    ProgressDialog.AddProgress(" Waiting " & Times(Counter) & " minutes for Import to finish")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Waiting " & Times(Counter) & " minutes for Import to finish" & Chr(13))
                    Sleep(Times(Counter) * 60000)
                    Counter += 1
                End While
                If Counter = 3 And Not Processes(0).HasExited Then
                    Dim SendMail As New System.Net.Mail.SmtpClient("relay1")
                    SendMail.Send("importfailure@pilottravelcenters.com", MailRecipient, Environment.MachineName & "Import error", "Import.exe is locked up at this location.  Please stop maintenance (Pos Transmitter), kill import.exe and importdata.exe, then restart maintenance and re-launch importdata.exe.")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Sent import failure e-mail")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR & Chr(13))
                    LogWriter.Close()
                    ProgressDialog.AddProgress(" Import.exe appears locked up")
                    ProgressDialog.AddProgress("... exiting")
                    ProgressDialog.Done()
                    Exit Sub
                End If
            Else
                'Make sure DbUpgrader isn't running (it shouldn't be), then rename it
                While ProcessRunning("DbUpgrader")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   DbUpgrader currently running; waiting 1 minute" & Chr(13))
                    Sleep(60000)
                End While
                'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Renaming DbUpgrader to DbUpgrader_hold" & Chr(13))
                'File.Move("C:\Office\Exe\DbUpgrader.exe", "C:\Office\Exe\DbUpgrader_hold.exe")
                'Call the retalix import program to load HOST download files
                ProgressDialog.AddProgress("Calling Import.exe")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Retalix import started" & Chr(13))
                If ShowProgress Then
                    Shell("C:\Office\Exe\Import.exe", AppWinStyle.MaximizedFocus, True)
                Else
                    Shell("C:\Office\Exe\Import.exe /M", AppWinStyle.Hide, True)
                End If
                ProgressDialog.AddProgress("Import.exe finished")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Retalix import ended" & Chr(13))
                'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Renaming DbUpgrader_hold to DbUpgrader" & Chr(13))
                'File.Move("C:\Office\Exe\DbUpgrader_hold.exe", "C:\Office\Exe\DbUpgrader.exe")
            End If

            'Delay because import seems to take a second or two to actually quit
            Sleep(5000)

            DatFileStillFound = Directory.GetFiles("c:\office\rcv", "*.dat").Length > 0

            While DatFileStillFound
                ImportLoopCount += 1
                If ImportLoopCount > 4 Then
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  Import has run 5 times and still has .dat files. Sending e-mail" & Chr(13))
                    Dim SendMail As New System.Net.Mail.SmtpClient("relay1")
                    SendMail.Send("autopostfailure@pilotcorp.com", MailRecipient, "STORE" & StoreNumber, "Import has run 5 times and still has .dat files to be processed. Please check.")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "  E-Mail sent" & Chr(13))
                    DatFileStillFound = False
                Else
                    ProgressDialog.AddProgress("Additional dat files found; running Import.exe again")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Additional dat files found; running import again" & Chr(13))
                    While ProcessRunning("DbUpgrader")
                        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   DbUpgrader currently running; waiting 1 minute" & Chr(13))
                        Sleep(60000)
                    End While
                    'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Renaming DbUpgrader to DbUpgrader_hold" & Chr(13))
                    'File.Move("C:\Office\Exe\DbUpgrader.exe", "C:\Office\Exe\DbUpgrader_hold.exe")

                    'Call the retalix import program to load HOST download files
                    ProgressDialog.AddProgress("Calling Import.exe")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Retalix import started" & Chr(13))
                    If ShowProgress Then
                        Shell("C:\Office\Exe\Import.exe", AppWinStyle.MaximizedFocus, True)
                    Else
                        Shell("C:\Office\Exe\Import.exe /M", AppWinStyle.Hide, True)
                    End If
                    ProgressDialog.AddProgress("Import.exe finished")
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Retalix import ended" & Chr(13))
                    'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Renaming DbUpgrader_hold to DbUpgrader" & Chr(13))
                    'File.Move("C:\Office\Exe\DbUpgrader_hold.exe", "C:\Office\Exe\DbUpgrader.exe")
                    'Delay because import seems to take a second or two to actually quit
                    Sleep(5000)
                    DatFileFound = True
                    DatFileStillFound = Directory.GetFiles("c:\office\rcv", "*.dat").Length > 0
                End If
            End While

            'Execute fix_cross_site procedure if we had a category file
            If HasCategoryFile Then
                If File.Exists(FixCrossSiteLog) Then
                    File.Delete(FixCrossSiteLog)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted FixCrossSite.log" & Chr(13))
                End If
                Writer = New StreamWriter(New FileStream(FixCrossSiteFile, FileMode.Create, FileAccess.Write, FileShare.None))
                Writer.WriteLine("execute procedure fix_cross_site;")
                Writer.WriteLine("commit;")
                Writer.Close()
                ProgressDialog.AddProgress("Calling Fix_cross_site procedure")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call fix_cross_site procedure started" & Chr(13))
                Shell(ISQLExe & OfficeDB & " -i " & FixCrossSiteFile & " -o " & FixCrossSiteLog & " -e -m -u sysdba -p masterkey", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call fix_cross_site procedure ended" & Chr(13))
                File.Delete(FixCrossSiteLog)
            End If

            'Call Update Item Status and Fixit
            If DatFileFound Then
                ProgressDialog.AddProgress("Calling Update Item Status")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Update Item Status started" & Chr(13))
                Shell("C:\office\exe\externaleventrunner.exe /updateitemstatus", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Update Item Status ended" & Chr(13))

                ' Create and exec file to call Fixit procedure
                If File.Exists(FixItLog) Then
                    File.Delete(FixItLog)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted Fixit.log" & Chr(13))
                End If
                Writer = New StreamWriter(New FileStream(FixItFile, FileMode.Create, FileAccess.Write, FileShare.None))
                Writer.WriteLine("execute procedure fixit;")
                Writer.WriteLine("commit;")
                Writer.Close()
                ProgressDialog.AddProgress("Calling Fixit procedure")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Fixit procedure started" & Chr(13))
                Shell(ISQLExe & OfficeDB & " -i " & FixItFile & " -o " & FixItLog & " -e -m -u sysdba -p masterkey", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Fixit procedure ended" & Chr(13))
                File.Delete(FixItFile)
            End If

            'Call Auto Price Srv
            ProgressDialog.AddProgress("Calling Auto Price Srv")
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Auto Price Srv started" & Chr(13))
            Shell("C:\office\exe\externaleventrunner.exe /autopricesrv", AppWinStyle.Hide, True)
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Auto Price Srv ended" & Chr(13))

            If File.Exists(ASNFileBackup) Then
                File.Delete(ASNFileBackup)
            End If

            'Start the invoice load process

            'Create ASN file array
            If Directory.Exists(DataDir) Then
                ASNFileArray = Directory.GetFiles(DataDir, ASNFileSpec)
            Else
                ASNFileArray = Nothing
            End If

            If Not ASNFileArray Is Nothing AndAlso ASNFileArray.Length > 0 Then
                ProgressDialog.AddProgress("Processing ASN files")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Found " & ASNFileArray.Length & " ASN file(s)" & Chr(13))
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Processing ASN files" & Chr(13))
                'Peform cleanup
                If File.Exists(ASNInsertLog) Then
                    File.Delete(ASNInsertLog)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted ASNInsert.log" & Chr(13))
                End If

                If File.Exists(ASNInsert) Then
                    File.Delete(ASNInsert)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted ASNInsert.sql" & Chr(13))
                End If

                ASNBackupArray = Directory.GetFiles(DataDir, ASNFileSpec & ".bak")
                For Each ASNFile In ASNBackupArray
                    File.Delete(ASNFile)
                Next
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted " & ASNBackupArray.Length & " backup files" & Chr(13))

                Writer = New StreamWriter(New FileStream(ASNInsert, FileMode.Create, FileAccess.Write, FileShare.None))

                Writer.WriteLine("delete from pilot_invoicestmp;" & Chr(13))
                Writer.WriteLine("commit;" & Chr(13))

                'Process ASN files
                For Each ASNFile In ASNFileArray
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Found input file " & ASNFile & Chr(13))
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

                    LogWriter.WriteLine(String.Format("{0}     Processed {1} line(s)", Now.ToString("MM/dd/yyyy HH:mm:ss"), Counter) & Chr(13))

                    Reader.Close()

                    'Backup current ASN file
                    File.Move(ASNFile, ASNFile & ".bak")
                Next

                Writer.WriteLine("commit;")
                Writer.WriteLine("execute procedure pilot_insertinvoices;")
                Writer.WriteLine("commit;")

                Writer.Close()

                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   ASNInsert.sql file created" & Chr(13))

                Sleep(2500)

                ProgressDialog.AddProgress("ASN isql started")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call ASN isql started" & Chr(13))

                Shell(ISQLExe & OfficeDB & " -i " & ASNInsert & " -o " & ASNInsertLog & " -e -m -u sysdba -p masterkey", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call ASN isql ended" & Chr(13))
            Else
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   No ASN Input files found" & Chr(13))
            End If

            'checks if store has a shipper file produced by HOST
            If HaveShippers Then
                ProgressDialog.AddProgress("Processing shipper files")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Import.exe processed a shipper file" & Chr(13))

                If File.Exists(ShipLog) Then
                    File.Delete(ShipLog)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted existing ShipInsert log" & Chr(13))
                End If

                If File.Exists(ShipSQL) Then
                    File.Delete(ShipSQL)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted existing ShipInsert SQL file" & Chr(13))
                End If

                If File.Exists(ShipBak) Then
                    File.Delete(ShipBak)
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Deleted existing Shipper backup file" & Chr(13))
                End If

                If File.Exists(ShipFile) Then
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Found Pilot Shipper file " & ShipFile & Chr(13))

                    Reader = New StreamReader(New FileStream(ShipFile, FileMode.Open, FileAccess.Read))
                    Writer = New StreamWriter(New FileStream(ShipSQL, FileMode.Create, FileAccess.Write, FileShare.None))

                    Writer.WriteLine("delete from pilot_shippers;" & Chr(13))
                    Writer.WriteLine("commit;" & Chr(13))
                    While Not Reader.EndOfStream
                        Rec = Reader.ReadLine()
                        Writer.WriteLine(Rec)
                    End While

                    Writer.WriteLine("commit;" & Chr(13))
                    Writer.WriteLine("execute procedure pilot_updateshippers;" & Chr(13))
                    Writer.WriteLine("commit;" & Chr(13))

                    Reader.Close()
                    Writer.Close()

                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   ShipInsert.sql file created" & Chr(13))

                    File.Move(ShipFile, ShipBak)
                Else
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Shipper File " & ShipFile & " not found" & Chr(13))
                    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   ShipInsert.sql file created (exec only)" & Chr(13))
                    Writer = New StreamWriter(New FileStream(ShipSQL, FileMode.Create, FileAccess.Write, FileShare.None))
                    Writer.WriteLine("execute procedure pilot_updateshippers;" & Chr(13))
                    Writer.WriteLine("commit;" & Chr(13))
                    Writer.Close()
                End If

                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Shipper isql started" & Chr(13))
                Shell(ISQLExe & OfficeDB & " -i " & ShipSQL & " -o " & ShipLog & " -e -m -u sysdba -p masterkey ", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call Shipper isql ended" & Chr(13))
            End If

            ProgressDialog.AddProgress("ImportData ended")
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " ImportData End. Script ended" & Chr(13))
            LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR & Chr(13))

            'Call RPO Prepare Data
            If (Date.Now.Hour >= 0 AndAlso Date.Now.Hour < 3) And DoRPOBuild Then
                ProgressDialog.AddProgress("Calling RPO Prepare Data")
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call RPO Prepare Data started" & Chr(13))
                idProg = Shell("C:\office\exe\ExternalEventRunner.exe /RPOpreparedata", AppWinStyle.Hide, True)
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Call RPO Prepare Data ended" & Chr(13))
                LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR & Chr(13))
            End If

            'sanity check to make sure DbUpgrader_hold got renamed back to DbUpgrader
            'If File.Exists("C:\Office\Exe\DbUpgrader_hold.exe") Then
            '    LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & "   Rename of DbUpgrader_hold to DbUpgrader failed; attempting again" & Chr(13))
            '    File.Move("C:\Office\Exe\DbUpgrader_hold.exe", "C:\Office\Exe\DbUpgrader.exe")
            'End If

            'LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR & Chr(13))
            LogWriter.Close()
            ProgressDialog.AddProgress("... done")
            ProgressDialog.Done()


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
        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " E-Mail sent" & Chr(13))

        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & " ImportData ended" & Chr(13))
        LogWriter.WriteLine(Now.ToString("MM/dd/yyyy HH:mm:ss") & SEPARATOR & Chr(13))
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

End Module