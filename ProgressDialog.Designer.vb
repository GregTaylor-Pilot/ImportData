<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProgressDialog
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.rtfProgress = New System.Windows.Forms.RichTextBox
        Me.lblNotice = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'rtfProgress
        '
        Me.rtfProgress.Enabled = False
        Me.rtfProgress.Location = New System.Drawing.Point(12, 12)
        Me.rtfProgress.Name = "rtfProgress"
        Me.rtfProgress.Size = New System.Drawing.Size(408, 256)
        Me.rtfProgress.TabIndex = 1
        Me.rtfProgress.Text = ""
        '
        'lblNotice
        '
        Me.lblNotice.AutoSize = True
        Me.lblNotice.Location = New System.Drawing.Point(12, 280)
        Me.lblNotice.Name = "lblNotice"
        Me.lblNotice.Size = New System.Drawing.Size(270, 26)
        Me.lblNotice.TabIndex = 2
        Me.lblNotice.Text = "This window will close when ImportData completes." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "See C:\Pilot\ImportData.log up" & _
            "on completion for results." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'ProgressDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(435, 315)
        Me.Controls.Add(Me.lblNotice)
        Me.Controls.Add(Me.rtfProgress)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ProgressDialog"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Import Data Progress"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rtfProgress As System.Windows.Forms.RichTextBox
    Friend WithEvents lblNotice As System.Windows.Forms.Label

End Class
