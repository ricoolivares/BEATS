Imports System.IO
Module Rutinas

    'General variables
    'Public gralInfo As GENERAL_INFO
    Public connstring As String

    'Connections database variables
    Public conn As Data.SqlClient.SqlConnection
    Public objCommand As Data.SqlClient.SqlCommand

   


    ''' -----------------------------------------------------------------------------
    ''' Project    : STEP
    ''' Function   : openConfigurationFile
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The openConfigurationFile function find the file (config.dat) wich contains
    ''' the string for connection to the database server, also contains some user information.
    ''' Finally initializate the the object for conect with ther database server (conn)
    ''' If there are any error trying to find the file, display a message box and end the program
    ''' </summary>
    ''' 
    ''' <returns>Returns true for a sucessful search and false for a unsucess search.</returns>
    ''' 
    ''' <remarks>
    ''' </remarks>
    ''' 
    ''' <history>
    '''    [Dario Rico]   11/02/2011   Created
    ''' </history>
    ''' 
    ''' -----------------------------------------------------------------------------
    Public Function openConfigurationFile() As String

        'Variable declaration
        Dim path As String, msg As String
        Dim result As String

         
        'Varibles for openning file
        Dim fs As FileStream
        Dim s As StreamReader

        'Initizalitation variables
        result = False
        msg = ""

        'Path for get the file
        path = Application.StartupPath.ToString + "\config.dat"

        'Opening the configurationfile for getting the string connection
        If IO.File.Exists(path) Then

            'REading the file
            fs = New FileStream(path, FileMode.Open, FileAccess.Read)
            s = New StreamReader(fs)

            'Getting information from configurtation file
            connstring = s.ReadLine            
            
            'Initialization of Objects for connections
            conn = New SqlClient.SqlConnection(connstring)

            'Testing connection (get Lines from database)
            result = testConnection()

            'Initialization object MiConfiguracion if there was NOT an error during test connection
            If result = "" Then
                'miConfiguracion = New SETTINGS(lineaID, ModelCode, kit, workStationID, userID, printer, scanComponents, labelPath)                 
            End If


            'Closing file
            s.Close()
            fs.Close()

        Else
            'If file is not found display a failed message
            msg = "          Error critico, no se localizo el archivo de configuración." + vbCrLf + _
                    "Contactar al departamento de informática para asistencia técnica."
            result = msg

        End If

        'Returning the value
        openConfigurationFile = result


    End Function



    Public Function saveConfigurationFile() As Boolean

        'Variable declaration
        Dim path As String, msg As String
        Dim result As Boolean
        'Initizalitation variables
        result = False
        msg = ""

        'Path for get the file
        path = Application.StartupPath.ToString + "\config.dat"
        Try
            Dim fs As New FileStream(path, FileMode.OpenOrCreate, FileAccess.Write)
            Dim s As New StreamWriter(fs)

            'Getting information from configurtation file
            s.WriteLine(connstring)
           
            s.Close()
            fs.Close()


            result = True

        Catch ex As Exception
            result = False
            MsgBox(ex.Message)
        End Try

        'Returning the value
        saveConfigurationFile = result


    End Function



    Public Sub displayResults(ByVal status As String, ByRef actPicture As PictureBox, ByRef actLblMessage As Label, ByVal actMessage As String)
        Dim player As New System.Media.SoundPlayer()

        If status = "GOOD" Then
            actPicture.Image = My.Resources.OK
            player.Stream = My.Resources.Good
            actLblMessage.ForeColor = Color.Green
        Else
            actPicture.Image = My.Resources.Bad
            player.Stream = My.Resources.Bad1
            actLblMessage.ForeColor = Color.Red
        End If

        'Show image, message and play sound
        actPicture.Visible = True
        actLblMessage.Visible = True
        actLblMessage.Text = actMessage
        player.Play()

    End Sub

    Public Sub hideResults(ByRef actPicture As PictureBox, ByRef actLblMessage As Label)
        'Show image, message and play sound
        actPicture.Visible = False
        actLblMessage.Visible = False
    End Sub


    Public Function testConnection() As String
        'Variable declaration
        Dim adapter As SqlClient.SqlDataAdapter
        Dim sql As String
        Dim result As String
        Dim dtTest As DataTable

        'Initialization variables
        result = ""
        adapter = Nothing

        Try
            '
            '--1. Openning connection on database
            '
            conn.Open()

            'Openning table and getting a test information
            sql = "SELECT LeagueID FROM Catalog_Leagues"
            adapter = New SqlClient.SqlDataAdapter(sql, conn)
            dtTest = New DataTable()
            adapter.Fill(dtTest)

            'Return an sucessuful result, no error message
            result = ""


        Catch ex As Exception
            'Returning a posible error conection
            result = "Error al establecer conexion en la base de datos."
        End Try

        'Closing connection and adapter
        If conn.State = ConnectionState.Open Then conn.Close()
        If Not adapter Is Nothing Then adapter.Dispose()

        'Returning the result
        testConnection = result

    End Function

    Public Sub closingApplication()
        'Closing opened forms        
        If Not frmLogin Is Nothing Then frmLogin.Dispose()


        'Closing Bartender applications hanged
        Dim pProcess() As Process = System.Diagnostics.Process.GetProcessesByName("bartend")
        For Each p As Process In pProcess
            p.Kill()
        Next

        'Closing Excel instances hanged
        Dim pProcess2() As Process = System.Diagnostics.Process.GetProcessesByName("Excel")
        For Each p As Process In pProcess2
            p.Kill()
        Next

        'Closing connection
        If conn.State = ConnectionState.Open Then conn.Close()

        'Quit from application
        Application.Exit()

    End Sub

End Module
