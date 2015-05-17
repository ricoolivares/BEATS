Imports BEATS.CLS_BEATS
Imports BEATS.CLS_GAME
Public Class frmBeats

    Private mySimulation As CLS_SIMULATION
    Private resulSimul() As CLS_GAME
    Private MiApuesta As New CLS_BEATS


    Private Sub btnValidate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnValidate.Click
        Dim listSimulation(2) As String
        Dim tmpSim() As CLS_GAME
        Dim i, n As Integer


        'listSimulation(1) = "DAL-LOCMED1"
        listSimulation(0) = "EMPMED1"
        listSimulation(1) = "LOCMED1"
        listSimulation(2) = "TEST"
        'listSimulation(2) = "EMPMED2"
        'listSimulation(4) = "BT-LOCMED1"
        'listSimulation(5) = "BT-VISMED1"
        'listSimulation(6) = "BT-EMPMED1"
        'listSimulation(7) = "BT-EMPMED2"
        'listSimulation(8) = "BT-EMPFAV1"

         
        i = 0
        For Each simul As String In listSimulation
            mySimulation = New CLS_SIMULATION(simul)

            tmpSim = Nothing
            Select Case UCase(mySimulation.Strategy)
                Case "ORMONDLIGHT"
                    tmpSim = mySimulation.generateSimulationOrmond(pbProgress, lblProgress)
                Case "D'ALEMBERT"
                    tmpSim = mySimulation.generateSimulationDAlambert(pbProgress, lblProgress)
                Case "D'ALEMBERTINV"
                    tmpSim = mySimulation.generateSimulationDAlambertInv(pbProgress, lblProgress)
            End Select

            'tmpSim = mySimulation.generateSimulationOrmond(pbProgress, lblProgress)
            If resulSimul Is Nothing Then
                resulSimul = tmpSim
            Else
                n = UBound(resulSimul) + 1
                ReDim Preserve resulSimul(n + UBound(tmpSim))
                tmpSim.CopyTo(resulSimul, n)

                n = resulSimul.Length

            End If

            'Concatenate Final
            i = i + 1

        Next


        'Display result
        'mySimulation.displayResulSim("", resulSimul, pbProgress, lblProgress)

        MsgBox("Simulacion generada exitosamente!!!", MsgBoxStyle.Information)

    End Sub


    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        'End program and applications hanged (Bartender and EXCEL)
        Call closingApplication()

    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        openFD.InitialDirectory = "E:\Projects\BEATS\ExcelFiles"
        openFD.Title = "Abrir Archivo de EXCEL"
        openFD.Filter = "Archivos Excel|*.xlsx|Archivos Excel|*.xls"
        openFD.ShowDialog()


        If Not DialogResult.Cancel Then
            'txtFilePath.Text = openFD.FileName
        End If



    End Sub



    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Call MiApuesta.uploadGames(pbProgress, lblProgress, txtLog)


    End Sub

    Private Sub frmLoadBeats_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim errorMsg As String

        'Opening configuration file and initialization an creating SETTING's class object (miConfiguracion)
        errorMsg = openConfigurationFile()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim resultValidate As Boolean

        resultValidate = False

        If txtFilePath.Text <> "" Then
            'Show ProgressBar()
            lblProgress.Visible = True : pbProgress.Visible = True

            'Carga partidos
            resultValidate = MiApuesta.validateGames(txtFilePath.Text, pbProgress, lblProgress, txtLog)

            If resultValidate Then
                gbSaveGames.Visible = True
                gbGettingGames.Visible = False
            End If

            'Hide ProgressBar()
            lblProgress.Visible = False : pbProgress.Visible = False

        Else
            'Error Message
            MsgBox("No se ha cargado archivo de EXCEL!!", MsgBoxStyle.Critical)

        End If
    End Sub

    Private Sub btnBrowse_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        openFD.InitialDirectory = "E:\Projects\BEATS\ExcelFiles"
        openFD.Title = "Abrir Archivo de EXCEL"
        openFD.Filter = "Archivos Excel|*.xlsx|Archivos Excel|*.xls"
        openFD.ShowDialog()


        If Not DialogResult.Cancel Then
            txtFilePath.Text = openFD.FileName
        End If

    End Sub

     
    

     
    Private Sub btnSentToDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSentToDB.Click
        'Show ProgressBar()
        lblProgress.Visible = True : pbProgress.Visible = True

        Call MiApuesta.uploadGames(pbProgress, lblProgress, txtLog)
        btnSentToDB.Enabled = False
        btnNewLoad.Visible = True

        'Hide ProgressBar()
        lblProgress.Visible = False : pbProgress.Visible = False

    End Sub

    
    Private Sub btnNewLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewLoad.Click

        btnSentToDB.Enabled = True
        btnNewLoad.Visible = False

        gbSaveGames.Visible = False
        gbGettingGames.Visible = True

    End Sub
End Class