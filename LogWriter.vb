Imports System.IO

Public Class LogWriter
    Private textOutput As System.Windows.Forms.RichTextBox
    Private writer As StreamWriter
    Private mLogLevel As Integer
    Const SEPARATOR As String = "****************************************************************************************************"

    Public Sub New(ByVal logName As String, Optional ByVal maxLogLevel As Integer = 10, Optional ByVal append As Boolean = True, Optional funcNum As Integer = -1)
        Try
            If append Then
                writer = New StreamWriter(New FileStream(logName, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            Else
                writer = New StreamWriter(New FileStream(logName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            End If
            writer.AutoFlush = True
            textOutput = Nothing
            If maxLogLevel < 21 AndAlso maxLogLevel > -1 Then
                mLogLevel = maxLogLevel
            Else
                mLogLevel = 10
            End If
            'If funcNum = -1 Then
            '    writer.Write(SEPARATOR & vbCrLf)
            'Else
            '    writer.Write(String.Format("({0}){1}{2}", funcNum.ToString.PadLeft(4, "0"c), SEPARATOR, vbCrLf))
            'End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub New(ByVal logName As String, ByRef rtfOutput As System.Windows.Forms.RichTextBox, Optional ByVal maxLogLevel As Integer = 10, Optional ByVal append As Boolean = True, Optional funcNum As Integer = -1)
        Try
            If append Then
                writer = New StreamWriter(New FileStream(logName, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            Else
                writer = New StreamWriter(New FileStream(logName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            End If
            writer.AutoFlush = True
            textOutput = rtfOutput
            If maxLogLevel < 21 AndAlso maxLogLevel > -1 Then
                mLogLevel = maxLogLevel
            Else
                mLogLevel = 10
            End If
            'If funcNum = -1 Then
            '    writer.Write(SEPARATOR & vbCrLf)
            'Else
            '    writer.Write(String.Format("({0}){1}{2}", funcNum.ToString.PadLeft(4, "0"c), SEPARATOR, vbCrLf))
            'End If

        Catch ex As Exception

        End Try

    End Sub

    Public Property LogLevel() As Integer
        Get
            Return mLogLevel
        End Get
        Set(ByVal value As Integer)
            If value < 21 AndAlso value > -1 Then
                mLogLevel = value
            Else
                mLogLevel = 10
            End If
            Log("Setting log level to " & mLogLevel, 1, 0)
        End Set
    End Property

    Public Sub Log(ByVal message As String, ByVal loglevel As Integer, ByVal FuncNum As Integer)
        If loglevel <= mLogLevel AndAlso loglevel > 0 Then
            Dim logLine As String = String.Format("{0}{1}({2}){3}", DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff"),
                                                  vbTab, FuncNum.ToString.PadLeft(4, "0"c), message) & vbCrLf
            If textOutput IsNot Nothing Then
                textOutput.Text &= logLine
            End If
            Try
                writer.Write(logLine)
            Catch ex As Exception

            End Try

        End If
    End Sub

    Public Sub Flush()
        Try
            writer.Flush()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub Close()
        Try
            writer.Close()
        Catch ex As Exception

        End Try

    End Sub
End Class
