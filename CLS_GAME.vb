Public Class CLS_GAME        

    'Games info
    Public HomeBeats As String      'MM-->MisMarcadores;    CL-->Caliente
    Public ActionType As String     '0--> Add Beats;  1--> Update Coutes;   2-->Update Results
    Public LeagueID As String
    Public League As String
    Public TeamIDLoc As String
    Public TeamIDVis As String
    Public Local As String
    Public Visita As String
    Public DatePlayed As Date
    Public TimePlayed As Date
    Public ResLoc As Integer
    Public ResVis As Integer
    Public MM_CuoLoc As Double
    Public MM_CuoEmp As Double
    Public MM_CuoVis As Double
    Public CuoteFilter As Double    
    Public WinLoc As Boolean
    Public WinEmp As Boolean
    Public WinVis As Boolean
    Public WinOvr As Boolean
    Public WinUnd As Boolean

    'SimulInfo
    Public SimulationID As String
    Public SimulNo As Integer
    Public GameNo As Integer
    Public GamePlayed As Boolean
    Public partido As String
    Public Marcador As String
    Public Estrategia As String
    Public tipoApuesta As String
    Public cuota As Double
    Public ApGanada As Boolean
    Public noFallos As Integer
    Public limNoFallos As Integer
    Public ApTotal As Double
    Public GanNeta As Double
    Public GanTotal As Double
    Public Impuesto As Double
    Public ganAcum As Double
    Public ApBase As Integer
    Public BankRoll As Integer




    'Constructor
    Public Sub New()
        MyBase.New()
    End Sub

    Public Function mttoGames() As String
        'Variable declaration
        Dim result As String

        Try
            'Openning connection with database
            If conn.State = ConnectionState.Closed Then conn.Open()

            objCommand = New SqlClient.SqlCommand()

            'Setting Parameters for store procedure: spInsertSerialComponents
            With objCommand
                .CommandText = "spMttoGamesData"
                .CommandType = CommandType.StoredProcedure
                .Connection = conn

                'Setting parameters
                With objCommand.Parameters
                    'Input parameters
                    .Add(New SqlClient.SqlParameter("@homeBeats", HomeBeats))
                    .Add(New SqlClient.SqlParameter("@actionType", ActionType))
                    .Add(New SqlClient.SqlParameter("@LeagueID", LeagueID))
                    .Add(New SqlClient.SqlParameter("@TeamIDLoc", TeamIDLoc))
                    .Add(New SqlClient.SqlParameter("@TeamIDVis", TeamIDVis))
                    .Add(New SqlClient.SqlParameter("@DatePlayed", DatePlayed))
                    .Add(New SqlClient.SqlParameter("@TimePlayed", TimePlayed))
                    .Add(New SqlClient.SqlParameter("@ResLoc", ResLoc))
                    .Add(New SqlClient.SqlParameter("@ResVis", ResVis))
                    .Add(New SqlClient.SqlParameter("@MM_CuoLoc", MM_CuoLoc))
                    .Add(New SqlClient.SqlParameter("@MM_CuoEmp", MM_CuoEmp))
                    .Add(New SqlClient.SqlParameter("@MM_CuoVis", MM_CuoVis))
                    .Add(New SqlClient.SqlParameter("@CL_CuoLoc", 0))
                    .Add(New SqlClient.SqlParameter("@CL_CuoEmp", 0))
                    .Add(New SqlClient.SqlParameter("@CL_CuoVis", 0))
                    .Add(New SqlClient.SqlParameter("@CL_CuoOvr", 0))
                    .Add(New SqlClient.SqlParameter("@CL_CuoUnd", 0))
                    .Add(New SqlClient.SqlParameter("@msg", " "))
                End With

                '----Setting Ouput parameters                
                'Msg
                .Parameters(17).Direction = ParameterDirection.Output
                .Parameters(17).Size = 100
                .Parameters(17).DbType = DbType.String


            End With

            'Executing store procedure for insertion
            objCommand.ExecuteNonQuery()

            'Getting results from store procedure
            result = Trim(objCommand.Parameters("@msg").Value)

        Catch ex As Exception
            'If an error occurs with conection with the server database
            result = "Error al ejecutar la accion:" + ex.Message
        Finally
            'closing objects
            If Not objCommand Is Nothing Then objCommand.Dispose()
            If conn.State = ConnectionState.Open Then conn.Close()

        End Try


        'Returning the result
        mttoGames = result

    End Function



End Class
