<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.lblClient = New System.Windows.Forms.Label()
        Me.ddlClients = New System.Windows.Forms.ComboBox()
        Me.lblSource = New System.Windows.Forms.Label()
        Me.ddlSource = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnSelect
        '
        Me.btnSelect.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelect.Location = New System.Drawing.Point(141, 127)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(231, 37)
        Me.btnSelect.TabIndex = 0
        Me.btnSelect.Text = "Select Files To Upload"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'lblClient
        '
        Me.lblClient.AutoSize = True
        Me.lblClient.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClient.Location = New System.Drawing.Point(110, 34)
        Me.lblClient.Name = "lblClient"
        Me.lblClient.Size = New System.Drawing.Size(99, 16)
        Me.lblClient.TabIndex = 2
        Me.lblClient.Text = "Select Client:"
        '
        'ddlClients
        '
        Me.ddlClients.FormattingEnabled = True
        Me.ddlClients.Location = New System.Drawing.Point(215, 33)
        Me.ddlClients.Name = "ddlClients"
        Me.ddlClients.Size = New System.Drawing.Size(158, 21)
        Me.ddlClients.TabIndex = 3
        '
        'lblSource
        '
        Me.lblSource.AutoSize = True
        Me.lblSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSource.Location = New System.Drawing.Point(148, 76)
        Me.lblSource.Name = "lblSource"
        Me.lblSource.Size = New System.Drawing.Size(61, 16)
        Me.lblSource.TabIndex = 4
        Me.lblSource.Text = "Source:"
        '
        'ddlSource
        '
        Me.ddlSource.FormattingEnabled = True
        Me.ddlSource.Items.AddRange(New Object() {"EMAIL", "PORTAL", "SFTP", "SDrive"})
        Me.ddlSource.Location = New System.Drawing.Point(215, 75)
        Me.ddlSource.Name = "ddlSource"
        Me.ddlSource.Size = New System.Drawing.Size(158, 21)
        Me.ddlSource.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(141, 193)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 13)
        Me.Label1.TabIndex = 6
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(511, 262)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ddlSource)
        Me.Controls.Add(Me.lblSource)
        Me.Controls.Add(Me.ddlClients)
        Me.Controls.Add(Me.lblClient)
        Me.Controls.Add(Me.btnSelect)
        Me.Name = "Form1"
        Me.Text = "EMSDBUploadFiles"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents lblClient As System.Windows.Forms.Label
    Friend WithEvents ddlClients As System.Windows.Forms.ComboBox
    Friend WithEvents lblSource As System.Windows.Forms.Label
    Friend WithEvents ddlSource As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
