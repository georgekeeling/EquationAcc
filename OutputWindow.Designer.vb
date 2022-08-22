<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OutputWindow
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
        Me.ResultsBox = New System.Windows.Forms.RichTextBox()
        Me.BtnCopyAll = New System.Windows.Forms.Button()
        Me.BtnClose = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ResultsBox
        '
        Me.ResultsBox.Location = New System.Drawing.Point(12, 12)
        Me.ResultsBox.Name = "ResultsBox"
        Me.ResultsBox.Size = New System.Drawing.Size(940, 620)
        Me.ResultsBox.TabIndex = 1
        Me.ResultsBox.Text = ""
        '
        'BtnCopyAll
        '
        Me.BtnCopyAll.Location = New System.Drawing.Point(600, 660)
        Me.BtnCopyAll.Name = "BtnCopyAll"
        Me.BtnCopyAll.Size = New System.Drawing.Size(150, 44)
        Me.BtnCopyAll.TabIndex = 2
        Me.BtnCopyAll.Text = "Copy All"
        Me.BtnCopyAll.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(770, 660)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(172, 44)
        Me.BtnClose.TabIndex = 3
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'OutputWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(974, 729)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.BtnCopyAll)
        Me.Controls.Add(Me.ResultsBox)
        Me.Name = "OutputWindow"
        Me.Text = "Results"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ResultsBox As Windows.Forms.RichTextBox
    Friend WithEvents BtnCopyAll As Windows.Forms.Button
    Friend WithEvents BtnClose As Windows.Forms.Button
End Class
