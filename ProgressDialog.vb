Imports System.Windows.Forms

Public Class ProgressDialog

    Public Sub AddProgress(ByVal progress As String)
        Try
            rtfProgress.Text = rtfProgress.Text & Now.ToString("HH:mm:ss ") & progress & ControlChars.CrLf
            Me.Refresh()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub Done()
        Try
            Me.Refresh()
        Catch ex As Exception

        End Try

        Threading.Thread.Sleep(3000)
    End Sub

End Class
