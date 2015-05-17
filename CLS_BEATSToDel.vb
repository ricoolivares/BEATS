Imports BEATS.CLS_GAME
Imports Microsoft.Office.Interop


Public Class CLS_BEATS

    Public myGames() As CLS_GAME

#Region "ReadingGames"

    Public Function validateGames(ByVal _fileName As String, ByRef _ProgressBar As ProgressBar, ByRef _lblProgress As Label, ByRef _txtLogErrors As TextBox) As Boolean
        'Vars for openning excel File
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        'Declaracion de variables                
        Dim i, totSheets, row, rowRes, totalRows, emptyRow As Integer
        Dim newGame, newDate, foundValidRows, foundInfo, validGame As Boolean
        Dim actualCell, logError, sheetName As String
        Dim myGame As CLS_GAME
        Dim posFound As Integer
        'Games Info
        Dim idLiga, idLocal, idVisita, equipoLoc, equipoVis, ResLoc, ResVis As String
        Dim miDate As Date
        Dim miTime, tempTime As String
        Dim tmpTime As Double
        Dim CuoLoc, CuoEmp, CuoVis As String
        Dim ts As TimeSpan
        'For control total
        Dim totNewGames, totUpdGamesRes, totUpdGamesCuo, totErrors As Double
        Dim hola As String
        Dim hola2 As String




        'Opening Excel Files
        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(_fileName)

        'Variable initialization
        totNewGames = 0              'For store the total of new games saved
        totUpdGamesRes = 0              'For store the total of updates games - Results
        totUpdGamesCuo = 0              'For store the total of updates games - Cuotes
        i = 1
        totSheets = 2
        sheetName = "Sheet1"
        rowRes = 0
        ReDim myGames(rowRes)

        'Go throw the Sheet1 to Sheet2 of Excel File
        For i = 1 To totSheets
            sheetName = "Sheet" + CStr(i)
            xlWorkSheet = xlWorkBook.Worksheets(sheetName)

            foundValidRows = True       'If five empty rows are found, then foundValidRows==False        
            emptyRow = 0
            row = 1                     'Row for start getting Games on template sheet        
            idLiga = ""
            foundInfo = True

            '
            '1. Gettting TotalRows from Sheet
            '
            'Loop throw all rows of actual sheet until not found games (Five consecutives empty rows)
            While (foundInfo = True)
                'Check if actual row is not empty
                actualCell = Trim(xlWorkSheet.Cells(row, "A").Value)

                'If emptyCell, increase counter, else reset emptyRow
                If actualCell = "" Then
                    emptyRow = emptyRow + 1

                    'If five consecutive emptyRows, exit from While
                    If emptyRow > 5 Then
                        foundInfo = False
                    End If
                Else
                    emptyRow = 0
                End If

                row = row + 1

            End While
            totalRows = row - 5

            'If found valid games then activate flag for get games
            If totalRows > 5 Then
                foundValidRows = True
                row = 1
            End If


            'Updating settings progressBar
            _ProgressBar.Minimum = 0
            _ProgressBar.Maximum = totalRows
            _ProgressBar.Value = 0

            'Un cambio importante
            If hola = hola2 Then
                hola = hola2
            End If


            'Loop throw all rows of actual sheet until not found games (Five consecutives empty rows)
            While (foundValidRows = True)
                'Check if actual row is not empty
                actualCell = Trim(xlWorkSheet.Cells(row, "A").Value)
                emptyRow = IIf(actualCell = "", emptyRow + 1, 0)


                'If found row with data, evaluate it!
                If actualCell <> "" Then
                    'Check if new Date
                    newDate = IIf(InStr(actualCell, "Todos") > 0, True, False)
                    If newDate Then
                        'Get Actual Date
                        row = row + 6
                        actualCell = Trim(xlWorkSheet.Cells(row, "A").Value)
                        miDate = gettingDate(actualCell)
                        row = row + 1
                    End If

                    'Get actual data
                    'row = row + 1
                    actualCell = Trim(UCase(xlWorkSheet.Cells(row, "A").Value))

                    'If cell is not a Game, then the league name, else, loop throw valid games                    
                    newGame = IIf(IsNumeric(actualCell) = True, True, False)
                    If Not newGame Then
                        'Getting LeagueID
                        actualCell = actualCell.Replace("'", "''")
                        idLiga = getLeagueID(actualCell)

                        If idLiga = "No se encontro liga!" Then
                            'Imprime error en Log
                            logError = "Line(" + CStr(row) + "): No se encontro Liga:" + actualCell + vbCrLf
                            _txtLogErrors.Text = _txtLogErrors.Text + logError
                            totErrors = totErrors + 1
                        End If

                    Else

                        'Loop throw all the valid games (cell with 'am' or 'pm')
                        While newGame
                            'Getting Game Information
                            If row = 42 Then
                                row = 42
                            End If


                            'Converting time
                            tmpTime = xlWorkSheet.Cells(row, "A").Value
                            ts = TimeSpan.FromDays(tmpTime)
                            tempTime = String.Format("{0}:{1}:{2}", ts.Hours.ToString("00"), ts.Minutes.ToString("00"), ts.Seconds.ToString("00"))
                            miTime = CDate(CStr(miDate) + " " + tempTime)

                            equipoLoc = Trim(xlWorkSheet.Cells(row, "B").Value)
                            equipoLoc = Trim(equipoLoc.Replace(Chr(160), Chr(32)))  'Delete blank spaces

                            equipoVis = Trim(xlWorkSheet.Cells(row, "C").Value)
                            equipoVis = Trim(equipoVis.Replace(Chr(160), Chr(32)))  'Delete blank spaces

                            idLocal = getTeamID(idLiga, equipoLoc)
                            If idLocal = "No se encontro equipo!" Then
                                'Imprime error en Log
                                logError = "Line(" + CStr(row) + "): No se encontro EquipoLocal: " + equipoLoc + " en Liga: " + idLiga + vbCrLf
                                _txtLogErrors.Text = _txtLogErrors.Text + logError
                                totErrors = totErrors + 1

                            End If

                            idVisita = getTeamID(idLiga, equipoVis)
                            If idVisita = "No se encontro equipo!" Then
                                'Imprime error en Log
                                logError = "Line(" + CStr(row) + "): No se encontro EquipoVisita: " + equipoVis + " en Liga: " + idLiga + vbCrLf
                                _txtLogErrors.Text = _txtLogErrors.Text + logError
                                totErrors = totErrors + 1
                            End If


                            'Process row just if no errors
                            If idVisita <> "No se encontro equipo!" And idLocal <> "No se encontro equipo!" And idLiga <> "No se encontro liga!" Then

                                ResLoc = getResul(Trim(xlWorkSheet.Cells(row, "D").Value), "LOCAL")
                                ResVis = getResul(Trim(xlWorkSheet.Cells(row, "D").Value), "VISITA")
                                'Getting cuotes information
                                CuoLoc = IIf(Trim(xlWorkSheet.Cells(row, "E").Value) = "-", 0, Trim(xlWorkSheet.Cells(row, "E").Value))
                                CuoEmp = IIf(Trim(xlWorkSheet.Cells(row, "F").Value) = "-", 0, Trim(xlWorkSheet.Cells(row, "F").Value))
                                CuoVis = IIf(Trim(xlWorkSheet.Cells(row, "G").Value) = "-", 0, Trim(xlWorkSheet.Cells(row, "G").Value))


                                'If game don't have coutes, then is NOT a valid game!
                                If CuoLoc = 0 Or CuoEmp = 0 Or CuoVis = 0 Then
                                    validGame = False
                                Else
                                    validGame = True
                                End If


                                'If valid game, evaluate it NOT game exist on dababase
                                If validGame And Not gameInsWithResOnDB(idLiga, idLocal, idVisita, miDate) Then

                                    'Create new Game
                                    myGame = New CLS_GAME
                                    myGame.HomeBeats = "MM"

                                    'Check if game exists on Vector MyGame()
                                    posFound = gameExistVector(idLiga, idLocal, idVisita, miDate)
                                    'If game already exist on actual Vector - myGames(), update just cuotes or Results
                                    If posFound <> -1 Then
                                        'If found game DON'T have results, update it!
                                        If myGames(posFound).ResLoc = -1 Or myGames(posFound).ResVis = -1 Then
                                            'Update results
                                            myGames(posFound).ResLoc = ResLoc
                                            myGames(posFound).ResVis = ResVis

                                            'Insert Game
                                            myGames(posFound).ActionType = 2       'ActionType=2 -->UpdateResults
                                            totUpdGamesRes = totUpdGamesRes + 1

                                        Else
                                            'Just update coutes
                                            If myGames(posFound).MM_CuoLoc = 0 Or myGames(posFound).MM_CuoVis = 0 Then
                                                myGames(posFound).MM_CuoLoc = CuoLoc
                                                myGames(posFound).MM_CuoEmp = CuoEmp
                                                myGames(posFound).MM_CuoVis = CuoVis

                                                'Insert Game
                                                myGames(posFound).ActionType = 1       'ActionType=1 -->UpdateCuotes
                                                totUpdGamesCuo = totUpdGamesCuo + 1

                                            End If

                                        End If
                                    Else
                                        'If game not was found on Vector

                                        'If Game doesn't exists on Database!
                                        If Not gameExist(idLiga, idLocal, idVisita, miDate) Then
                                            'Insert Game
                                            myGame.ActionType = 0       'ActionType=0 -->InsertGame

                                            'Update Counters                            
                                            totNewGames = totNewGames + 1

                                        Else




                                            'Check type of Update
                                            If ResLoc = -1 Or ResVis = -1 Then
                                                'If actGame NOT has RESULTS then-- Update <<CUOTES>> ActionType=1
                                                myGame.ActionType = 1
                                                totUpdGamesCuo = totUpdGamesCuo + 1
                                            Else
                                                'If actGame has RESULTS then-- Update <<RESULTS>> ActionType=2
                                                myGame.ActionType = 2
                                                totUpdGamesRes = totUpdGamesRes + 1
                                            End If                                             

                                        End If  'End If gameExist on Database

                                        myGame.LeagueID = idLiga
                                        myGame.TeamIDLoc = idLocal
                                        myGame.TeamIDVis = idVisita
                                        myGame.DatePlayed = miDate
                                        myGame.TimePlayed = miTime
                                        myGame.ResLoc = ResLoc
                                        myGame.ResVis = ResVis
                                        myGame.MM_CuoLoc = CuoLoc
                                        myGame.MM_CuoEmp = CuoEmp
                                        myGame.MM_CuoVis = CuoVis

                                        'Update counters
                                        ReDim Preserve myGames(rowRes)
                                        myGames(rowRes) = myGame
                                        rowRes = rowRes + 1

                                    End If  'End If gameExist on Vector - posFound<>-1


                                End If 'End if ValidGame condition


                            End If 'End idVisita <> "No se encontro equipo!" condition




                            'Check if next row is a valid Game
                            row = row + 1
                            actualCell = Trim(UCase(xlWorkSheet.Cells(row, "A").Value))
                            newGame = IIf(IsNumeric(actualCell) = True, True, False)

                            'If not new game, subs one row
                            If Not newGame Then row = row - 1


                            'Update progressBar
                            If rowRes <= (totalRows - 2) Then
                                _ProgressBar.Value = row + 2
                                _lblProgress.Text = "Processing Games: NewGames:[" + CStr(totNewGames) + _
                                                " ] -- UpdGamesResults[" + CStr(totUpdGamesRes) + "]-- UpdGamesCuotes:[" + CStr(totUpdGamesCuo) + "] -- Errors:[" + CStr(totErrors) + "]..." & Format((row / (totalRows - 2)), "0%")
                            End If

                        End While  'End while  found GAMES

                    End If   'End found Game condition

                Else
                    'If found emptyRow, check if is the third one, and End theLoop while
                    If emptyRow = 10 Then foundValidRows = False
                End If

                'Increase counter for row
                row = row + 1

            End While   'End foundValidRows while 

            row = i
           

        Next



        'Closing ExcelFiles
        xlWorkBook.Close()
        xlApp.Quit()
        xlWorkSheet = Nothing
        xlWorkBook = Nothing
        xlApp = Nothing


        'Returning results
        If totErrors = 0 Then
            validateGames = True
        Else
            validateGames = False
        End If

        MsgBox("Juegos Leidos Exitosamente!", MsgBoxStyle.Information)

    End Function

    Private Function getTotalRows(ByRef _sheet As Excel.Worksheet) As Integer
        Dim foundInfo As Boolean
        Dim row, emptyRow As Integer
        Dim actualCell As String


        foundInfo = True

        'Get TotalRows
        'Loop throw all rows of sheet [dataMM] until not found games (Five consecutives empty rows)
        While (foundInfo = True)
            'Check if actual row is not empty
            actualCell = Trim(_sheet.Cells(row, "A").Value)

            'If emptyCell, increase counter, else reset emptyRow
            If actualCell = "" Then
                emptyRow = emptyRow + 1

                'If five consecutive emptyRows, exit from While
                If emptyRow > 5 Then
                    foundInfo = False
                End If
            Else
                emptyRow = 0
            End If

            row = row + 1

        End While

        getTotalRows = row - 5

    End Function

    Private Function gettingDate(ByVal tempStr As String) As Date
        Dim result, monthStr, dayStr, yearStr As String
        Dim pos1, pos2 As Integer

        'Getting Day
        pos1 = InStr(1, tempStr, "/")
        dayStr = Trim(Mid(tempStr, 1, pos1 - 1))
        'Getting Month
        pos2 = InStr(pos1 + 1, tempStr, " ")
        monthStr = Trim(Mid(tempStr, pos1 + 1, pos2 - pos1))

        'Getting final DATE
        yearStr = CStr(Today.Year)
        'If CInt(monthStr) = 1 Then
        '    yearStr = "2015"
        'Else
        '    yearStr = "2014"
        'End If
        result = monthStr + "/" + dayStr + "/" + yearStr

        'Returning the Final Date
        gettingDate = CDate(result)

    End Function

    Private Function getHomeBeats(ByVal _fileName As String) As String
        Dim result As String
        Dim pos As Integer

        result = ""
        _fileName = Trim(UCase(_fileName))

        'If fileName contains "MisMarcadores" then results=MM
        pos = InStr("MISMARCADORES", _fileName, CompareMethod.Text)
        If pos > 0 Then
            result = "MM"
        End If

        'If fileName contains "Caliente" then results=MM
        pos = InStr("CALIENTE", _fileName, CompareMethod.Text)
        If pos > 0 Then
            result = "CL"
        End If

        getHomeBeats = result

    End Function

    Private Function getLeagueID(ByVal _League As String) As String
        'Variable declaration
        Dim result As String, sql As String
        'Variable declaration        
        Dim cmd As SqlClient.SqlCommand


        'Variable initializations
        '_League = UCase(Trim(Application.WorksheetFunction.Substitute(_League, Chr(160), Chr(32)))) 'Function for elimate blanck space web <&nbsp>
        '_League = UCase(Trim(_League.Replace(Chr(160), Chr(32))))    'Function for elimate blanck space web <&nbsp>

        result = ""
        cmd = Nothing
        sql = "SELECT  LeagueID  FROM    Catalog_Leagues " + _
              "WHERE     RTRIM(League) LIKE  '%" + Trim(_League) + "%' OR  RTRIM(League2) LIKE  '%" + Trim(_League) + "%' OR  RTRIM(League3) LIKE  '%" + Trim(_League) + "%'"

        Try
            'Openning connection with database
            If conn.State = ConnectionState.Closed Then conn.Open()
            'Executing Scalar SQL Command
            cmd = New SqlClient.SqlCommand(sql, conn)

            'Getting result
            result = cmd.ExecuteScalar()

        Catch ex As Exception
            result = "No se encontro liga!"
        Finally
            'Closing objectCommand  and connection      
            If Not cmd Is Nothing Then cmd.Dispose()
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try

        If result = "" Then
            result = "No se encontro liga!"
        End If

        getLeagueID = result

    End Function

    Private Function getTeamID(ByVal _LeagueID As String, ByVal _Team As String) As String
        'Variable declaration
        Dim result As String, sql As String
        'Variable declaration        
        Dim cmd As SqlClient.SqlCommand


        'Variable initializations
        result = ""
        _Team = _Team.Replace("'", "''")
        cmd = Nothing
        sql = "SELECT     TeamID FROM Catalog_Teams " + _
              "WHERE     RTrim(LeagueID) = '" + Trim(_LeagueID) + "' AND RTrim(Team) = '" + Trim(_Team) + "'"

        Try
            'Openning connection with database
            If conn.State = ConnectionState.Closed Then conn.Open()
            'Executing Scalar SQL Command
            cmd = New SqlClient.SqlCommand(sql, conn)

            'Getting result
            result = cmd.ExecuteScalar()

        Catch ex As Exception
            result = "No se encontro equipo!"
        Finally
            'Closing objectCommand  and connection      
            If Not cmd Is Nothing Then cmd.Dispose()
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try

        If result = "" Then
            result = "No se encontro equipo!"
        End If

        getTeamID = result

    End Function

    Private Function getResul(ByVal tmpValue As String, ByVal tipo As String) As Integer
        Dim pos As Integer
        Dim subStr As String

        subStr = ""
        tmpValue = Trim(tmpValue)
        pos = InStr(1, tmpValue, ":")

        'teamToSearch = UCase(Trim(Application.WorksheetFunction.Substitute(teamToSearch, Chr(160), Chr(32)))) 'Function for elimate blanck space web <&nbsp>

        If pos > 0 Then
            'If result to get is Local, get Left part from "-"
            If UCase(tipo) = "LOCAL" Then
                subStr = Trim(Left(tmpValue, pos - 2))
                getResul = CInt(subStr)
            Else
                'If result to get is Visita, get Right part from "-"
                subStr = Mid(tmpValue, (pos + 2), Len(tmpValue))
                getResul = CInt(subStr)
            End If

        Else
            getResul = -1
        End If


    End Function

    Public Function gameExist(ByVal _LeagueID As String, ByVal _TeamIDLoc As String, _
                              ByVal _TeamIDVis As String, ByVal _datePlay As Date) As Boolean
        'Variable declaration        
        Dim cmd As SqlClient.SqlCommand
        Dim qtyRecords As New Integer
        Dim sql As String
        Dim result As Boolean

        'Initialization variables
        result = False
        cmd = Nothing
        sql = ""

        Try
            '
            '--1. Openning connection on database
            '
            conn.Open()

            'Setting query  according to the number of serial code


            sql = "SELECT      COUNT(LeagueID) FROM GamesData " + _
                  "WHERE       LeagueID = '" + _LeagueID + "' AND  TeamIDLoc = '" + _TeamIDLoc + "' AND " + _
                  "            TeamIDVis = '" + _TeamIDVis + "' AND  DatePlayed = '" + CStr(_datePlay) + "'"

            cmd = New SqlClient.SqlCommand(sql, conn)
            qtyRecords = cmd.ExecuteScalar()

            If qtyRecords > 0 Then
                'Found Game
                result = True
                'End Add by Dario
            Else
                result = False
            End If

        Catch ex As Exception
            'Returning a posible error conection
            'result = "Error al establecer conexion en la base de datos."
            result = False
        End Try

        'Closing connection and adapter
        If conn.State = ConnectionState.Open Then conn.Close()
        If Not cmd Is Nothing Then cmd.Dispose()

        'Returning the result
        gameExist = result

    End Function

    Public Function gameInsWithResOnDB(ByVal _LeagueID As String, ByVal _TeamIDLoc As String, _
                              ByVal _TeamIDVis As String, ByVal _datePlay As Date) As Boolean
        'Variable declaration        
        Dim cmd As SqlClient.SqlCommand
        Dim qtyRecords As New Integer
        Dim sql As String
        Dim result As Boolean

        'Initialization variables
        result = False
        cmd = Nothing
        sql = ""

        Try
            '
            '--1. Openning connection on database
            '
            conn.Open()

            'Setting query  according to the number of serial code


            sql = "SELECT      COUNT(LeagueID) FROM GamesData " + _
                  "WHERE       LeagueID = '" + _LeagueID + "' AND  TeamIDLoc = '" + _TeamIDLoc + "' AND " + _
                  "            TeamIDVis = '" + _TeamIDVis + "' AND  DatePlayed = '" + CStr(_datePlay) + "' AND ReslOC<>-1 and ResVis<>-1"

            cmd = New SqlClient.SqlCommand(sql, conn)
            qtyRecords = cmd.ExecuteScalar()

            If qtyRecords > 0 Then
                'Found Game
                result = True
                'End Add by Dario
            Else
                result = False
            End If

        Catch ex As Exception
            'Returning a posible error conection
            'result = "Error al establecer conexion en la base de datos."
            result = False
        End Try

        'Closing connection and adapter
        If conn.State = ConnectionState.Open Then conn.Close()
        If Not cmd Is Nothing Then cmd.Dispose()

        'Returning the result
        gameInsWithResOnDB = result

    End Function

    
    Public Function gameExistVector(ByVal _LeagueID As String, ByVal _TeamIDLoc As String, _
                             ByVal _TeamIDVis As String, ByVal _datePlay As Date) As Integer
        'Variable declaration                        
        Dim result, i, totRecords As Integer
        Dim actGame As New CLS_GAME

        'Initialization variables
        result = -1
        i = 0
        totRecords = myGames.Length


        'Search trough all the actual vector games-- Mygames()
        For i = 0 To totRecords - 1

            'actGame = New CLS_GAME
            actGame = myGames(i)

            If Not actGame Is Nothing Then 'First element is nothing

                'If found record
                If actGame.LeagueID = _LeagueID And actGame.TeamIDLoc = _TeamIDLoc And _
                   actGame.TeamIDVis = _TeamIDVis And actGame.DatePlayed = _datePlay Then

                    'Exit with a TRUE value (record was found)
                    result = i

                    Exit For

                End If
            End If



        Next


        'Returning the result
        gameExistVector = result

    End Function


#End Region

#Region "UploadGames"

    Public Function uploadGames(ByRef _ProgressBar As ProgressBar, _
                              ByRef _lblProgress As Label, ByRef _txtLogErrors As TextBox) As Boolean

        Dim totalGames, totNewGames, totUpdGames, totErrors, row As Integer
        Dim logError, errorMsg As String



        totalGames = myGames.Length
        row = 1
        totErrors = 0


        'Updating settings progressBar
        _ProgressBar.Minimum = 0
        _ProgressBar.Maximum = totalGames
        _ProgressBar.Value = 0



        For Each game As CLS_GAME In myGames

            errorMsg = ""

            'Upload / Update Games's Data on Database BEATS
            errorMsg = game.mttoGames()


            'Updating counters
            'If no errors
            If errorMsg = "" Then
                If game.ActionType = 0 Then
                    totNewGames = totNewGames + 1
                Else
                    totUpdGames = totUpdGames + 1
                End If
            Else
                'If errors
                logError = "Index(" + CStr(row) + "): NoBeats:" + game.LeagueID + "-" + game.TeamIDLoc + "-" + game.TeamIDVis + "-" + CStr(game.DatePlayed) + "   ErrorDesc:" + errorMsg
                _txtLogErrors.Text = _txtLogErrors.Text + logError

                totErrors = totErrors + 1
            End If


            'Update progressBar
            If row <= totalGames Then
                _ProgressBar.Value = row
                _lblProgress.Text = "Processing Games: NewGames[" + CStr(totNewGames) + _
                                " ] -- UpdatedGames[" + CStr(totUpdGames) + "] -- Errors[" + CStr(totErrors) + "]..." & Format((row / totalGames), "0%")
            End If



            'Update Counters
            row = row + 1
        Next

        MsgBox("Juegos guardados Exitosamente!", MsgBoxStyle.Information)


    End Function
#End Region

    




End Class
