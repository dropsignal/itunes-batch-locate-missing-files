<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.ButtonScan = New System.Windows.Forms.Button()
        Me.TextBoxLog = New System.Windows.Forms.TextBox()
        Me.TextBoxMediaFolder = New System.Windows.Forms.TextBox()
        Me.ButtonBrowse = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ButtonScan
        '
        Me.ButtonScan.Enabled = False
        Me.ButtonScan.Location = New System.Drawing.Point(755, 41)
        Me.ButtonScan.Name = "ButtonScan"
        Me.ButtonScan.Size = New System.Drawing.Size(75, 23)
        Me.ButtonScan.TabIndex = 3
        Me.ButtonScan.Text = "&Scan"
        Me.ButtonScan.UseVisualStyleBackColor = True
        '
        'TextBoxLog
        '
        Me.TextBoxLog.Location = New System.Drawing.Point(12, 41)
        Me.TextBoxLog.Multiline = True
        Me.TextBoxLog.Name = "TextBoxLog"
        Me.TextBoxLog.ReadOnly = True
        Me.TextBoxLog.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBoxLog.Size = New System.Drawing.Size(737, 397)
        Me.TextBoxLog.TabIndex = 2
        '
        'TextBoxMediaFolder
        '
        Me.TextBoxMediaFolder.Location = New System.Drawing.Point(12, 12)
        Me.TextBoxMediaFolder.Name = "TextBoxMediaFolder"
        Me.TextBoxMediaFolder.Size = New System.Drawing.Size(737, 23)
        Me.TextBoxMediaFolder.TabIndex = 0
        '
        'ButtonBrowse
        '
        Me.ButtonBrowse.Location = New System.Drawing.Point(755, 12)
        Me.ButtonBrowse.Name = "ButtonBrowse"
        Me.ButtonBrowse.Size = New System.Drawing.Size(75, 23)
        Me.ButtonBrowse.TabIndex = 1
        Me.ButtonBrowse.Text = "&Browse"
        Me.ButtonBrowse.UseVisualStyleBackColor = True
        '
        'FormMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(838, 450)
        Me.Controls.Add(Me.ButtonBrowse)
        Me.Controls.Add(Me.TextBoxMediaFolder)
        Me.Controls.Add(Me.TextBoxLog)
        Me.Controls.Add(Me.ButtonScan)
        Me.Name = "FormMain"
        Me.Text = "Batch Locate Missing Files in iTunes"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ButtonScan As Button
    Friend WithEvents TextBoxLog As TextBox
    Friend WithEvents TextBoxMediaFolder As TextBox
    Friend WithEvents ButtonBrowse As Button
End Class
