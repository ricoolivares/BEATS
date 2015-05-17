<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBeats
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
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.txtLog = New System.Windows.Forms.TextBox()
        Me.btnValidate = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.openFD = New System.Windows.Forms.OpenFileDialog()
        Me.lblLogErrors = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.gbSaveGames = New System.Windows.Forms.GroupBox()
        Me.btnNewLoad = New System.Windows.Forms.Button()
        Me.btnSentToDB = New System.Windows.Forms.Button()
        Me.gbGettingGames = New System.Windows.Forms.GroupBox()
        Me.txtFilePath = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.lblDate2 = New System.Windows.Forms.Label()
        Me.lblDate1 = New System.Windows.Forms.Label()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.pbProgress = New System.Windows.Forms.ProgressBar()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.gbSaveGames.SuspendLayout()
        Me.gbGettingGames.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Black
        Me.lblTitle.Location = New System.Drawing.Point(112, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(382, 25)
        Me.lblTitle.TabIndex = 17
        Me.lblTitle.Text = "** BEATS 2014 **"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblTitle.Visible = False
        '
        'txtLog
        '
        Me.txtLog.Location = New System.Drawing.Point(24, 119)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLog.Size = New System.Drawing.Size(546, 144)
        Me.txtLog.TabIndex = 18
        '
        'btnValidate
        '
        Me.btnValidate.Location = New System.Drawing.Point(328, 235)
        Me.btnValidate.Name = "btnValidate"
        Me.btnValidate.Size = New System.Drawing.Size(176, 45)
        Me.btnValidate.TabIndex = 1
        Me.btnValidate.Text = "GENERAR SIMULACIONES"
        Me.btnValidate.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(303, 450)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(84, 48)
        Me.btnExit.TabIndex = 19
        Me.btnExit.Text = "SALIR"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'openFD
        '
        Me.openFD.FileName = "OpenFileDialog1"
        '
        'lblLogErrors
        '
        Me.lblLogErrors.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLogErrors.ForeColor = System.Drawing.Color.Black
        Me.lblLogErrors.Location = New System.Drawing.Point(25, 91)
        Me.lblLogErrors.Name = "lblLogErrors"
        Me.lblLogErrors.Size = New System.Drawing.Size(109, 25)
        Me.lblLogErrors.TabIndex = 17
        Me.lblLogErrors.Text = "Log errores:"
        Me.lblLogErrors.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 37)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(655, 353)
        Me.TabControl1.TabIndex = 22
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.gbSaveGames)
        Me.TabPage1.Controls.Add(Me.gbGettingGames)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(647, 327)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Cargado Juegos"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'gbSaveGames
        '
        Me.gbSaveGames.Controls.Add(Me.btnNewLoad)
        Me.gbSaveGames.Controls.Add(Me.btnSentToDB)
        Me.gbSaveGames.Location = New System.Drawing.Point(27, 33)
        Me.gbSaveGames.Name = "gbSaveGames"
        Me.gbSaveGames.Size = New System.Drawing.Size(420, 164)
        Me.gbSaveGames.TabIndex = 20
        Me.gbSaveGames.TabStop = False
        Me.gbSaveGames.Text = "2. Guardar Partidos en la Base de Datos."
        Me.gbSaveGames.Visible = False
        '
        'btnNewLoad
        '
        Me.btnNewLoad.Location = New System.Drawing.Point(89, 83)
        Me.btnNewLoad.Name = "btnNewLoad"
        Me.btnNewLoad.Size = New System.Drawing.Size(165, 40)
        Me.btnNewLoad.TabIndex = 0
        Me.btnNewLoad.Text = "NUEVA CARGA"
        Me.btnNewLoad.UseVisualStyleBackColor = True
        Me.btnNewLoad.Visible = False
        '
        'btnSentToDB
        '
        Me.btnSentToDB.Location = New System.Drawing.Point(89, 26)
        Me.btnSentToDB.Name = "btnSentToDB"
        Me.btnSentToDB.Size = New System.Drawing.Size(165, 40)
        Me.btnSentToDB.TabIndex = 0
        Me.btnSentToDB.Text = "ENVIO A BASE DE DATOS2"
        Me.btnSentToDB.UseVisualStyleBackColor = True
        '
        'gbGettingGames
        '
        Me.gbGettingGames.Controls.Add(Me.txtFilePath)
        Me.gbGettingGames.Controls.Add(Me.Button1)
        Me.gbGettingGames.Controls.Add(Me.btnBrowse)
        Me.gbGettingGames.Controls.Add(Me.lblFile)
        Me.gbGettingGames.Controls.Add(Me.lblLogErrors)
        Me.gbGettingGames.Controls.Add(Me.txtLog)
        Me.gbGettingGames.Location = New System.Drawing.Point(27, 33)
        Me.gbGettingGames.Name = "gbGettingGames"
        Me.gbGettingGames.Size = New System.Drawing.Size(586, 279)
        Me.gbGettingGames.TabIndex = 18
        Me.gbGettingGames.TabStop = False
        Me.gbGettingGames.Text = "1. Obtener Partidos"
        '
        'txtFilePath
        '
        Me.txtFilePath.Location = New System.Drawing.Point(24, 44)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(443, 20)
        Me.txtFilePath.TabIndex = 24
        Me.txtFilePath.Text = "E:\Projects\BEATS\ExcelFiles\JustTesting.xlsx"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(225, 82)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(99, 31)
        Me.Button1.TabIndex = 21
        Me.Button1.Text = "VALIDAR"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(495, 44)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(75, 23)
        Me.btnBrowse.TabIndex = 23
        Me.btnBrowse.Text = "Examinar..."
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'lblFile
        '
        Me.lblFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFile.ForeColor = System.Drawing.Color.Black
        Me.lblFile.Location = New System.Drawing.Point(24, 16)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(70, 25)
        Me.lblFile.TabIndex = 22
        Me.lblFile.Text = "Archivo:"
        Me.lblFile.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(23, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(382, 16)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "** Simulaciones **"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.lblDate2)
        Me.TabPage2.Controls.Add(Me.lblDate1)
        Me.TabPage2.Controls.Add(Me.DateTimePicker2)
        Me.TabPage2.Controls.Add(Me.DateTimePicker1)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.btnValidate)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(647, 327)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Simulaciones"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'lblDate2
        '
        Me.lblDate2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate2.Location = New System.Drawing.Point(272, 96)
        Me.lblDate2.Name = "lblDate2"
        Me.lblDate2.Size = New System.Drawing.Size(50, 23)
        Me.lblDate2.TabIndex = 22
        Me.lblDate2.Text = "Final:"
        '
        'lblDate1
        '
        Me.lblDate1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate1.Location = New System.Drawing.Point(67, 100)
        Me.lblDate1.Name = "lblDate1"
        Me.lblDate1.Size = New System.Drawing.Size(63, 23)
        Me.lblDate1.TabIndex = 21
        Me.lblDate1.Text = "Inicial:"
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker2.Location = New System.Drawing.Point(328, 96)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(107, 20)
        Me.DateTimePicker2.TabIndex = 20
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(136, 100)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(107, 20)
        Me.DateTimePicker1.TabIndex = 19
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(96, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(382, 25)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "** Simulaciones **"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pbProgress
        '
        Me.pbProgress.Location = New System.Drawing.Point(20, 396)
        Me.pbProgress.Name = "pbProgress"
        Me.pbProgress.Size = New System.Drawing.Size(647, 23)
        Me.pbProgress.TabIndex = 24
        '
        'lblProgress
        '
        Me.lblProgress.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress.ForeColor = System.Drawing.Color.Black
        Me.lblProgress.Location = New System.Drawing.Point(20, 422)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(643, 25)
        Me.lblProgress.TabIndex = 23
        Me.lblProgress.Text = "Progreso..."
        Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmBeats
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(679, 511)
        Me.Controls.Add(Me.pbProgress)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.lblTitle)
        Me.Name = "frmBeats"
        Me.Text = "BEATS 2014"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.gbSaveGames.ResumeLayout(False)
        Me.gbGettingGames.ResumeLayout(False)
        Me.gbGettingGames.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents txtLog As System.Windows.Forms.TextBox
    Friend WithEvents btnValidate As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents openFD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblLogErrors As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents gbGettingGames As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents pbProgress As System.Windows.Forms.ProgressBar
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents gbSaveGames As System.Windows.Forms.GroupBox
    Friend WithEvents btnNewLoad As System.Windows.Forms.Button
    Friend WithEvents btnSentToDB As System.Windows.Forms.Button
    Friend WithEvents lblDate2 As System.Windows.Forms.Label
    Friend WithEvents lblDate1 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
End Class
