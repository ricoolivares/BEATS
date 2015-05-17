Imports BEATS.CLS_GAME
Imports Microsoft.Office.Interop

Public Class CLS_SIMULATION
    'Information for Simulation
    Private SimulationID As String
    Private TipoApuesta As String
    Private RangeCoutes As String
    Public Strategy As String
    Private Description As String
    Private FilterSql As String
    Private ApBase As Integer
    Private LimNoFails As Integer
    Private QtyIncrease As Integer
    Private BankRoll As Integer
    Private FiltroTipoApuesta As String


    Public Sub New(ByVal _simulID As String)

        'Update Simulation Information, based on SimulationID came from _simulID
        updateSimulInfo(_simulID)


    End Sub

    '
    'Descripcion estrategia Ormond:  
    '                   Apostar la misma cantidad inicial (A) cada vez que se gane
    '                   Apostar la cantidad (A + Increm) cada vez que se pierda
    '                       Condicion:  si NoFallos llega al limNoFallos, entonces se reinicia apuesta (A)
    '
    Public Function generateSimulationOrmond(ByRef _ProgressBar As ProgressBar, ByRef _lblProgress As Label) As CLS_GAME()
        'Variable declaration
        Dim simInfo, resulSimul(), Partidos() As CLS_GAME
        Dim actualLeague As String
        Dim i, idPartido, noFallos, rowRes, totalPartidos As Integer
        Dim apTotal, ganNeta, Impuesto, GanTotal, ganAcum As Double
        Dim newLeague, apAntGanada As Boolean

        'Variable Initialization
        simInfo = Nothing
        resulSimul = Nothing

        Partidos = getValidGames()
        newLeague = False


        'Obteniendo el total de partidos generados
        i = LBound(Partidos)
        totalPartidos = UBound(Partidos)

        '--Delete simulGames if found validGames
        If totalPartidos > 0 Then
            delSimulGames(SimulationID)
        End If

        'Updating settings progressBar
        _ProgressBar.Minimum = 0
        _ProgressBar.Maximum = totalPartidos
        _ProgressBar.Value = 0



        'Recorre todos los partidos generados
        actualLeague = ""

        While (i <= totalPartidos)

            simInfo = New CLS_GAME
            simInfo.SimulNo = i
            simInfo.LeagueID = Partidos(i).LeagueID


            If Partidos(i).LeagueID <> actualLeague Then
                'New League-->Initializate Beats
                actualLeague = Partidos(i).LeagueID
                apTotal = 0
                ganNeta = 0
                Impuesto = 0
                GanTotal = 0
                ganAcum = 0
                newLeague = True
                idPartido = 0
                noFallos = 0
            End If


            '
            '---- 1. Obteniedo la cantidad a Apostar
            '

            'Checa si es la primer apuesta
            If simInfo.SimulNo = 0 Or newLeague Then
                'Si es la primer Apuesta, asigna la apBase
                apTotal = ApBase
                noFallos = 0
                newLeague = False
            Else
                'Verifica el resultado de la apuesta anterior
                Select Case UCase(TipoApuesta)
                    Case "LOCAL"
                        apAntGanada = Partidos(i - 1).WinLoc

                    Case "EMPATE"
                        apAntGanada = Partidos(i - 1).WinEmp

                    Case "VISITA"
                        apAntGanada = Partidos(i - 1).WinVis

                End Select



                'Si no es la primer apuesta, verifica el resultado de la ApAnterior
                If (apAntGanada) Or (noFallos = LimNoFails) Then
                    'Si la apuesta Ant fue ganada OR noFallos llego al limite de noFallos, entonces  apTotal=apBase
                    apTotal = ApBase
                    noFallos = 0
                Else
                    'Si la apuesta Ant fue perdida agrega Incremento (QtyIncrease)
                    noFallos = noFallos + 1
                    apTotal = apTotal + QtyIncrease
                End If

            End If

            '
            '---- 2. Verificando resultado para calculo de Ganancias o perdidas
            '

            'Actualizando datos de la apuesta, con la informacion del partido


            'Call updateInfoSim(Partidos(i), simInfo)
            idPartido = idPartido + 1
            simInfo.GameNo = idPartido
            simInfo.DatePlayed = Partidos(i).DatePlayed
            simInfo.TimePlayed = Partidos(i).TimePlayed
            simInfo.TeamIDLoc = Partidos(i).TeamIDLoc
            simInfo.TeamIDVis = Partidos(i).TeamIDVis
            simInfo.Local = Partidos(i).Local
            simInfo.Visita = Partidos(i).Visita
            simInfo.partido = Partidos(i).Local + " VS " + Partidos(i).Visita
            simInfo.Marcador = CStr(Partidos(i).ResLoc) + " -- " + CStr(Partidos(i).ResVis)

            'Get apGanada and Coutes
            Select Case UCase(TipoApuesta)
                Case "LOCAL"
                    simInfo.ApGanada = Partidos(i).WinLoc
                    simInfo.cuota = Partidos(i).MM_CuoLoc

                Case "EMPATE"
                    simInfo.ApGanada = Partidos(i).WinEmp
                    simInfo.cuota = Partidos(i).MM_CuoEmp

                Case "VISITA"
                    simInfo.ApGanada = Partidos(i).WinVis
                    simInfo.cuota = Partidos(i).MM_CuoVis
            End Select

            'Get FilterCuotes
            Select Case UCase(FiltroTipoApuesta)
                Case "LOCAL"
                    simInfo.CuoteFilter = Partidos(i).MM_CuoLoc
                Case "EMPATE"
                    simInfo.CuoteFilter = Partidos(i).MM_CuoEmp
                Case "VISITA"
                    simInfo.CuoteFilter = Partidos(i).MM_CuoVis
            End Select



            'Calculos de ganancias y perdidas
            If simInfo.ApGanada Then
                '-- Si la apuesta fue ganada                
                GanTotal = apTotal * (simInfo.cuota - 1)
                ganAcum = ganAcum + GanTotal
            Else
                '-- Si la apuesta fue perdida                
                GanTotal = apTotal * -1
                ganAcum = ganAcum + GanTotal
            End If

            'Si el juego no se ha jugado, no calcular ganancias
            If Partidos(i).ResLoc = -1 Then
                simInfo.GamePlayed = False
                ganNeta = 0
                Impuesto = 0
                GanTotal = 0
                ganAcum = 0
                noFallos = 0
            Else
                simInfo.GamePlayed = True
            End If




            '
            '---- 3. Agrega simulacion a Vector de resultados
            '


            'Initialization
            If resulSimul Is Nothing Then
                ReDim resulSimul(i)
            Else
                i = UBound(resulSimul) + 1
                ReDim Preserve resulSimul(i)
            End If

            resulSimul(i) = New CLS_GAME
            resulSimul(i) = simInfo            
            resulSimul(i).SimulationID = SimulationID            
            resulSimul(i).tipoApuesta = TipoApuesta
            resulSimul(i).Estrategia = Strategy
            resulSimul(i).limNoFallos = LimNoFails
            resulSimul(i).ApBase = ApBase
            resulSimul(i).BankRoll = BankRoll
            resulSimul(i).noFallos = IIf(Not simInfo.ApGanada, noFallos + 1, noFallos)
            resulSimul(i).ApTotal = apTotal
            resulSimul(i).GanNeta = ganNeta
            resulSimul(i).Impuesto = Impuesto
            resulSimul(i).GanTotal = GanTotal
            resulSimul(i).ganAcum = ganAcum

            'Add simulation to Database
            mttoSimulGames(resulSimul(i))


            'Counters increase
            i = i + 1
            'row = row + 1
            rowRes = rowRes + 1



            'Update progressBar
            If rowRes <= totalPartidos Then
                _ProgressBar.Value = i
                _lblProgress.Text = "Juegos Procesados: [" + CStr(i) + "] Juegos --  Porcentaje:[" + Format((i / totalPartidos), "0%") + "]"
            End If


        End While


        'Display Result

        generateSimulationOrmond = resulSimul



    End Function

    '
    'Descripcion estrategia D'Alembert:  
    '                   Restar (apBase) cuando se gane
    '                   Añadir (apBase) cuando se pierda 
    '                       Condicion:  
    '                       Nota: Metodo apropiado para apostar a "FAVORITOS" -- cuotas bajas
    '
    Public Function generateSimulationDAlambert(ByRef _ProgressBar As ProgressBar, ByRef _lblProgress As Label) As CLS_GAME()
        'Variable declaration
        Dim simInfo, resulSimul(), Partidos() As CLS_GAME
        Dim actualLeague As String
        Dim i, idPartido, noFallos, rowRes, totalPartidos As Integer
        Dim apTotal, ganNeta, Impuesto, GanTotal, ganAcum, ganAcumTmp As Double
        Dim newLeague, apAntGanada As Boolean

        'Variable Initialization
        simInfo = Nothing
        resulSimul = Nothing

        Partidos = getValidGames()
        newLeague = False


        'Obteniendo el total de partidos generados
        i = LBound(Partidos)
        totalPartidos = UBound(Partidos)
        '--Delete simulGames if found validGames
        If totalPartidos > 0 Then
            delSimulGames(SimulationID)
        End If


        'Updating settings progressBar
        _ProgressBar.Minimum = 0
        _ProgressBar.Maximum = totalPartidos
        _ProgressBar.Value = 0



        'Recorre todos los partidos generados
        actualLeague = ""

        While (i <= totalPartidos)

            simInfo = New CLS_GAME
            simInfo.SimulNo = i
            simInfo.LeagueID = Partidos(i).LeagueID


            If Partidos(i).LeagueID <> actualLeague Then
                'New League-->Initializate Beats
                actualLeague = Partidos(i).LeagueID
                apTotal = 0
                ganNeta = 0
                Impuesto = 0
                GanTotal = 0
                ganAcum = 0
                ganAcumTmp = 0
                newLeague = True
                idPartido = 0
                noFallos = 0
            End If


            '
            '---- 1. Obteniedo la cantidad a Apostar
            '

            'Checa si es la primer apuesta
            If simInfo.SimulNo = 0 Or newLeague Then
                'Si es la primer Apuesta, asigna la apBase
                apTotal = ApBase
                noFallos = 0
                newLeague = False
                ganAcumTmp = 0
            Else
                'Verifica el resultado de la apuesta anterior
                Select Case UCase(TipoApuesta)
                    Case "LOCAL"
                        apAntGanada = Partidos(i - 1).WinLoc

                    Case "EMPATE"
                        apAntGanada = Partidos(i - 1).WinEmp

                    Case "VISITA"
                        apAntGanada = Partidos(i - 1).WinVis

                End Select



                'Si la ApAnt fue ganada
                If apAntGanada Then
                    'Si ApAnt fue ganada, restar apBase
                    apTotal = apTotal - ApBase
                    If apTotal <= ApBase Then
                        apTotal = ApBase
                    End If
                    noFallos = 0
                Else
                    'si ApAnt fue perdida, añade apBase
                    If apTotal >= (LimNoFails * ApBase) Then
                        apTotal = ApBase
                        ganAcumTmp = 0
                    Else
                        apTotal = apTotal + ApBase
                    End If
                    noFallos = noFallos + 1

                End If


            End If

            '
            '---- 2. Verificando resultado para calculo de Ganancias o perdidas
            '

            'Actualizando datos de la apuesta, con la informacion del partido



            'Call updateInfoSim(Partidos(i), simInfo)
            idPartido = idPartido + 1
            simInfo.GameNo = idPartido
            simInfo.DatePlayed = Partidos(i).DatePlayed
            simInfo.TimePlayed = Partidos(i).TimePlayed
            simInfo.TeamIDLoc = Partidos(i).TeamIDLoc
            simInfo.TeamIDVis = Partidos(i).TeamIDVis
            simInfo.Local = Partidos(i).Local
            simInfo.Visita = Partidos(i).Visita
            simInfo.partido = Partidos(i).Local + " VS " + Partidos(i).Visita
            simInfo.Marcador = CStr(Partidos(i).ResLoc) + " -- " + CStr(Partidos(i).ResVis)


            'Get apGanada and Coutes
            Select Case UCase(TipoApuesta)

                Case "LOCAL"
                    simInfo.ApGanada = Partidos(i).WinLoc
                    simInfo.cuota = Partidos(i).MM_CuoLoc

                Case "EMPATE"
                    simInfo.ApGanada = Partidos(i).WinEmp
                    simInfo.cuota = Partidos(i).MM_CuoEmp

                Case "VISITA"
                    simInfo.ApGanada = Partidos(i).WinVis
                    simInfo.cuota = Partidos(i).MM_CuoVis


            End Select

            'Get FilterCuotes
            Select Case UCase(FiltroTipoApuesta)
                Case "LOCAL"
                    simInfo.CuoteFilter = Partidos(i).MM_CuoLoc
                Case "EMPATE"
                    simInfo.CuoteFilter = Partidos(i).MM_CuoEmp
                Case "VISITA"
                    simInfo.CuoteFilter = Partidos(i).MM_CuoVis
            End Select



            'Calculos de ganancias y perdidas
            If simInfo.ApGanada Then
                '-- Si la apuesta fue ganada
                ganNeta = (apTotal * simInfo.cuota) - apTotal
                GanTotal = ganNeta
                ganAcum = ganAcum + GanTotal
                ganAcumTmp = ganAcumTmp + GanTotal
            Else
                '-- Si la apuesta fue perdida
                ganNeta = apTotal * -1
                GanTotal = ganNeta
                ganAcum = ganAcum + GanTotal
                ganAcumTmp = ganAcumTmp + GanTotal
            End If



            'Si el juego no se ha jugado, no calcular ganancias
            If Partidos(i).ResLoc = -1 Then
                ganNeta = 0
                Impuesto = 0
                GanTotal = 0
                ganAcum = 0
                noFallos = 0
            End If




            '
            '---- 3. Agrega simulacion a Vector de resultados
            '


            'Initialization
            If resulSimul Is Nothing Then
                ReDim resulSimul(i)
            Else
                i = UBound(resulSimul) + 1
                ReDim Preserve resulSimul(i)
            End If

            resulSimul(i) = New CLS_GAME
            resulSimul(i) = simInfo
            resulSimul(i).SimulationID = SimulationID
            resulSimul(i).tipoApuesta = TipoApuesta
            resulSimul(i).Estrategia = Strategy
            resulSimul(i).limNoFallos = ganAcumTmp
            resulSimul(i).ApBase = ApBase
            resulSimul(i).BankRoll = BankRoll
            resulSimul(i).noFallos = noFallos
            resulSimul(i).ApTotal = apTotal
            resulSimul(i).GanNeta = ganNeta
            resulSimul(i).Impuesto = Impuesto
            resulSimul(i).GanTotal = GanTotal
            resulSimul(i).ganAcum = ganAcum

            'Add simulation to Database
            mttoSimulGames(resulSimul(i))

            'Counters increase
            i = i + 1
            'row = row + 1
            rowRes = rowRes + 1



            'Update progressBar
            If rowRes <= totalPartidos Then
                _ProgressBar.Value = i
                _lblProgress.Text = "Juegos Procesados: [" + CStr(i) + "] Juegos --  Porcentaje:[" + Format((i / totalPartidos), "0%") + "]"
            End If


        End While


        'Display Result

        generateSimulationDAlambert = resulSimul



    End Function



    '
    'Descripcion estrategia D'Alembert INVERSO:  
    '                   Añadir (apBase) cuando se gane
    '                   Restar (apBase) cuando se pierda 
    '                       Condicion:  
    '                       Nota: Metodo apropiado para apostar a "NO FAVORITOS" -- cuotas altas
    '
    Public Function generateSimulationDAlambertInv(ByRef _ProgressBar As ProgressBar, ByRef _lblProgress As Label) As CLS_GAME()
        'Variable declaration
        Dim simInfo, resulSimul(), Partidos() As CLS_GAME
        Dim actualLeague As String
        Dim i, idPartido, noFallos, rowRes, totalPartidos As Integer
        Dim apTotal, ganNeta, Impuesto, GanTotal, ganAcum, ganAcumTmp As Double
        Dim newLeague, apAntGanada As Boolean

        'Variable Initialization
        simInfo = Nothing
        resulSimul = Nothing

        Partidos = getValidGames()
        newLeague = False


        'Obteniendo el total de partidos generados
        i = LBound(Partidos)
        totalPartidos = UBound(Partidos)
        '--Delete simulGames if found validGames
        If totalPartidos > 0 Then
            delSimulGames(SimulationID)
        End If


        'Updating settings progressBar
        _ProgressBar.Minimum = 0
        _ProgressBar.Maximum = totalPartidos
        _ProgressBar.Value = 0



        'Recorre todos los partidos generados
        actualLeague = ""

        While (i <= totalPartidos)

            simInfo = New CLS_GAME
            simInfo.SimulNo = i
            simInfo.LeagueID = Partidos(i).LeagueID


            If Partidos(i).LeagueID <> actualLeague Then
                'New League-->Initializate Beats
                actualLeague = Partidos(i).LeagueID
                apTotal = 0
                ganNeta = 0
                Impuesto = 0
                GanTotal = 0
                ganAcum = 0
                ganAcumTmp = 0
                newLeague = True
                idPartido = 0
                noFallos = 0
            End If


            '
            '---- 1. Obteniedo la cantidad a Apostar
            '

            'Checa si es la primer apuesta
            If simInfo.SimulNo = 0 Or newLeague Then
                'Si es la primer Apuesta, asigna la apBase
                apTotal = ApBase
                noFallos = 0
                newLeague = False
                ganAcumTmp = 0
            Else
                'Verifica el resultado de la apuesta anterior
                Select Case UCase(TipoApuesta)
                    Case "LOCAL"
                        apAntGanada = Partidos(i - 1).WinLoc

                    Case "EMPATE"
                        apAntGanada = Partidos(i - 1).WinEmp

                    Case "VISITA"
                        apAntGanada = Partidos(i - 1).WinVis

                End Select



                'Si la ApAnt fue ganada
                If apAntGanada Then
                    'Si ApAnt fue ganada, añade (apBase)                    
                    If apTotal >= (LimNoFails * ApBase) Then
                        apTotal = ApBase
                        ganAcumTmp = 0
                    Else
                        apTotal = apTotal + ApBase
                    End If
                    noFallos = 0
                Else
                    'Si ApAnt fue perdida, restar (apBase)
                    apTotal = apTotal - ApBase
                    If apTotal <= ApBase Then
                        apTotal = ApBase
                    End If
                    noFallos = noFallos + 1
                End If


            End If

            '
            '---- 2. Verificando resultado para calculo de Ganancias o perdidas
            '

            'Actualizando datos de la apuesta, con la informacion del partido

            'Call updateInfoSim(Partidos(i), simInfo)
            idPartido = idPartido + 1
            simInfo.GameNo = idPartido
            simInfo.DatePlayed = Partidos(i).DatePlayed
            simInfo.TimePlayed = Partidos(i).TimePlayed
            simInfo.TeamIDLoc = Partidos(i).TeamIDLoc
            simInfo.TeamIDVis = Partidos(i).TeamIDVis
            simInfo.Local = Partidos(i).Local
            simInfo.Visita = Partidos(i).Visita
            simInfo.partido = Partidos(i).Local + " VS " + Partidos(i).Visita
            simInfo.Marcador = CStr(Partidos(i).ResLoc) + " -- " + CStr(Partidos(i).ResVis)


            'Get apGanada and Coutes
            Select Case UCase(TipoApuesta)

                Case "LOCAL"
                    simInfo.ApGanada = Partidos(i).WinLoc
                    simInfo.cuota = Partidos(i).MM_CuoLoc

                Case "EMPATE"
                    simInfo.ApGanada = Partidos(i).WinEmp
                    simInfo.cuota = Partidos(i).MM_CuoEmp

                Case "VISITA"
                    simInfo.ApGanada = Partidos(i).WinVis
                    simInfo.cuota = Partidos(i).MM_CuoVis


            End Select

            'Get FilterCuotes
            Select Case UCase(FiltroTipoApuesta)
                Case "LOCAL"
                    simInfo.CuoteFilter = Partidos(i).MM_CuoLoc
                Case "EMPATE"
                    simInfo.CuoteFilter = Partidos(i).MM_CuoEmp
                Case "VISITA"
                    simInfo.CuoteFilter = Partidos(i).MM_CuoVis
            End Select


            'Calculos de ganancias y perdidas
            If simInfo.ApGanada Then
                '-- Si la apuesta fue ganada
                ganNeta = (apTotal * simInfo.cuota) - apTotal
                GanTotal = ganNeta
                ganAcum = ganAcum + GanTotal
                ganAcumTmp = ganAcumTmp + GanTotal
            Else
                '-- Si la apuesta fue perdida
                ganNeta = apTotal * -1                
                GanTotal = ganNeta
                ganAcum = ganAcum + GanTotal
                ganAcumTmp = ganAcumTmp + GanTotal
            End If


            'Si el juego no se ha jugado, no calcular ganancias
            If Partidos(i).ResLoc = -1 Then
                simInfo.GamePlayed = False
                ganNeta = 0
                Impuesto = 0
                GanTotal = 0
                ganAcum = 0
                noFallos = 0
            Else
                simInfo.GamePlayed = True
            End If



            '
            '---- 3. Agrega simulacion a Vector de resultados
            '

 
            'Initialization
            If resulSimul Is Nothing Then
                ReDim resulSimul(i)
            Else
                i = UBound(resulSimul) + 1
                ReDim Preserve resulSimul(i)
            End If

            resulSimul(i) = New CLS_GAME
            resulSimul(i) = simInfo
            resulSimul(i).SimulationID = SimulationID
            resulSimul(i).tipoApuesta = TipoApuesta
            resulSimul(i).Estrategia = Strategy
            resulSimul(i).limNoFallos = ganAcumTmp
            resulSimul(i).ApBase = ApBase
            resulSimul(i).BankRoll = BankRoll
            resulSimul(i).noFallos = noFallos
            resulSimul(i).ApTotal = apTotal
            resulSimul(i).GanNeta = ganNeta
            resulSimul(i).Impuesto = Impuesto
            resulSimul(i).GanTotal = GanTotal
            resulSimul(i).ganAcum = ganAcum

            'Add simulation to Database
            mttoSimulGames(resulSimul(i))


            'Counters increase
            i = i + 1
            'row = row + 1
            rowRes = rowRes + 1



            'Update progressBar
            If rowRes <= totalPartidos Then
                _ProgressBar.Value = i
                _lblProgress.Text = "Juegos Procesados: [" + CStr(i) + "] Juegos --  Porcentaje:[" + Format((i / totalPartidos), "0%") + "]"
            End If


        End While


        'Display Result

        generateSimulationDAlambertInv = resulSimul



    End Function

    

    Private Function getValidGames() As CLS_GAME()
        'Declaracion de variables
        Dim totalPart, partSel As Integer
        Dim i, j As Integer
        Dim cuotaAct, cuotaSig As Double
        Dim DifMinutes As Integer
        Dim sameDate As Boolean
        Dim validGames(), tmpGames() As CLS_GAME

        'Get games from database
        tmpGames = getGames()

        'Variable Initialization        
        i = LBound(tmpGames)
        totalPart = UBound(tmpGames)
        validGames = Nothing


        '
        '2. Obtencion de partidos validos... filtrar partidos con rangos de 2 hrs de separacion
        '        
        While (i <= totalPart)

            '---Seleccione el primer registro de la fechaActual como como PartidoSeleccionado
            j = i
            partSel = i


            'Checa si el partido actual y el siguiente estan en las mismas fechas,
            '(siempre que no sea el ultimo registro).
            If (i <= totalPart - 1) Then
                sameDate = IIf(tmpGames(partSel).DatePlayed = tmpGames(j + 1).DatePlayed, True, False)

                'Si el partido no esta en la misma fecha, agregalo
                If Not sameDate Then
                    'Agrega la informacion del partido al arreglo dinamico                    
                    addGameToVector(validGames, tmpGames(partSel))

                    'Avanza apuntador al siguiente partido
                    j = j + 1
                End If
            Else
                'Si es el ultimo registro
                sameDate = False

                'Agrega la informacion del partido al arreglo dinamico                    
                addGameToVector(validGames, tmpGames(partSel))

                'Avanza apuntador para finalizar ciclo
                j = j + 1

            End If


            'Repetir mientras que sean partidos con la misma fecha
            While (sameDate And j <= totalPart)
                '-- Verifica si las hrs de los dos partidos a comparar NO se solapan (2 hrs)

                'Parameter on DateDiff ("n") means compare minutes, so 2hrs-->120 mins
                DifMinutes = DateDiff("n", tmpGames(partSel).TimePlayed, tmpGames(j + 1).TimePlayed)

                'Si el horario de los partidos se SOLAPAN --> (partidos con Menos de 2 Hrs(120Mins) de Diferencia)
                If DifMinutes < 120 Then
                    'Getting better Coutes 
                    Select Case UCase(FiltroTipoApuesta)
                        Case "LOCAL"
                            cuotaAct = tmpGames(partSel).MM_CuoLoc
                            cuotaSig = tmpGames(j + 1).MM_CuoLoc
                        Case "EMPATE"
                            cuotaAct = tmpGames(partSel).MM_CuoEmp
                            cuotaSig = tmpGames(j + 1).MM_CuoEmp

                        Case "VISITA"
                            cuotaAct = tmpGames(partSel).MM_CuoVis
                            cuotaSig = tmpGames(j + 1).MM_CuoVis
                    End Select


                    '--Selecciona partido con menor cuota
                    If (cuotaAct <= cuotaSig) Then
                        'Marca el partido ACTUAL como seleccionado
                        partSel = partSel
                    Else
                        'Marca el partido SIGUIENTE como seleccionado
                        partSel = j + 1
                    End If
                    

                    'Avanza apuntador al siguiente partido
                    j = j + 1

                Else
                    'Si el horario de los partidos NO se SOLAPAN --> (partidos con mas de 2 Hrs de Diferencia)

                    'Agrega la informacion del partido al arreglo dinamico
                    addGameToVector(validGames, tmpGames(partSel))

                    'Avanza apuntador al siguiente partido
                    j = j + 1

                    'Marcada el siguiente partido como seleccionado
                    partSel = j

                End If

                '-- Verifica si el siguiente partido esta en la misma fecha
                'Checa si el partido actual y el siguiente estan en las mismas fechas,
                '(siempre que no sea el ultimo registro).
                If (j <= totalPart - 1) Then
                    sameDate = IIf(tmpGames(partSel).DatePlayed = tmpGames(j + 1).DatePlayed, True, False)
                Else
                    sameDate = False
                End If

                If Not sameDate Then
                    'Agrega la informacion del partido al arreglo dinamico                    
                    addGameToVector(validGames, tmpGames(partSel))

                    'Avanza apuntador al siguiente partido
                    j = j + 1
                End If

            End While

            'Counter increase
            i = j
        End While

        'Return result
        getValidGames = validGames


    End Function


    Private Function getGames() As CLS_GAME()

        'Variable declaration
        Dim adapter As SqlClient.SqlDataAdapter
        Dim sql, errorMsg As String
        Dim tblTmp As DataTable
        Dim rowRes As Integer
        Dim tmpGames() As CLS_GAME


        'Initialization variables
        errorMsg = ""
        adapter = Nothing
        tblTmp = Nothing
        rowRes = 0
        tmpGames = Nothing




        Try
            '
            '--1. Openning connection on database
            '
            If conn.State = ConnectionState.Closed Then conn.Open()


            'Getting employee information from View EmployeeInfo - Apply Filter from CLASS
            'sql = "SELECT     LeagueId, League, DatePlayed, TimePlayed, TeamIDLoc, TeamIDVis,Local, Visita, ResLoc,  " + _
            '      "           ResVis, MM_CuoLoc, MM_CuoEmp, MM_CuoVis " + _
            '      "FROM      [BEATS].[dbo].[vwGamesData] " + _
            '      "WHERE    (DatePlayed>='2/1/2015') AND (Active=1) AND (" + FilterSql + ")  " + _
            '      "ORDER BY LeagueId, DatePlayed, TimePlayed, Local, Visita "

            sql = "SELECT     LeagueId, League, DatePlayed, TimePlayed, TeamIDLoc, TeamIDVis,Local, Visita, ResLoc,  " + _
                  "           ResVis, MM_CuoLoc, MM_CuoEmp, MM_CuoVis " + _
                  "FROM      [BEATS].[dbo].[vwGamesData] " + _
                  "WHERE    (DatePlayed>='1/1/2015') AND (Active=1)  AND (" + FilterSql + ")  " + _
                  "ORDER BY LeagueId, DatePlayed, TimePlayed, Local, Visita "



            adapter = New SqlClient.SqlDataAdapter(sql, conn)
            tblTmp = New DataTable()
            adapter.Fill(tblTmp)

            For Each row As DataRow In tblTmp.Rows

                'If first element
                If rowRes = 0 Then
                    ReDim tmpGames(rowRes)
                Else
                    ReDim Preserve tmpGames(rowRes)
                End If


                'Save on vector+
                tmpGames(rowRes) = New CLS_GAME
                tmpGames(rowRes).LeagueID = row.Item("LeagueID")
                tmpGames(rowRes).League = row.Item("League")
                tmpGames(rowRes).DatePlayed = row.Item("DatePlayed")
                tmpGames(rowRes).TimePlayed = row.Item("TimePlayed")
                tmpGames(rowRes).TeamIDLoc = row.Item("TeamIDLoc")
                tmpGames(rowRes).TeamIDVis = row.Item("TeamIDVis")
                tmpGames(rowRes).Local = row.Item("Local")
                tmpGames(rowRes).Visita = row.Item("Visita")
                tmpGames(rowRes).ResLoc = row.Item("ResLoc")
                tmpGames(rowRes).ResVis = row.Item("ResVis")
                tmpGames(rowRes).MM_CuoLoc = row.Item("MM_CuoLoc")
                tmpGames(rowRes).MM_CuoEmp = row.Item("MM_CuoEmp")
                tmpGames(rowRes).MM_CuoVis = row.Item("MM_CuoVis")

                'Getting Win flags
                If tmpGames(rowRes).ResLoc <> -1 Then

                    tmpGames(rowRes).WinLoc = IIf(tmpGames(rowRes).ResLoc > tmpGames(rowRes).ResVis, 1, 0)
                    tmpGames(rowRes).WinEmp = IIf(tmpGames(rowRes).ResLoc = tmpGames(rowRes).ResVis, 1, 0)
                    tmpGames(rowRes).WinVis = IIf(tmpGames(rowRes).ResLoc < tmpGames(rowRes).ResVis, 1, 0)
                    tmpGames(rowRes).WinOvr = IIf((tmpGames(rowRes).ResLoc + tmpGames(rowRes).ResVis) > 3, 1, 0)
                    tmpGames(rowRes).WinUnd = IIf((tmpGames(rowRes).ResLoc + tmpGames(rowRes).ResVis) < 3, 1, 0)

                End If


                'Increase counter
                rowRes = rowRes + 1


            Next


        Catch ex As Exception
            'Returning a posible error conection
            errorMsg = "Error de conexion:" + ex.Message
        Finally
            'Closing connection and adapter            
            If Not adapter Is Nothing Then adapter.Dispose()
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try

        If errorMsg <> "" Then
            tmpGames = Nothing
        End If

        'Returning the result
        getGames = tmpGames




    End Function


    Private Sub addGameToVector(ByRef gamesVector() As CLS_GAME, ByRef _gameToAdd As CLS_GAME)
        'Variable declaration
        Dim i As Integer

        'Initialization
        If gamesVector Is Nothing Then
            ReDim gamesVector(i)
        Else
            i = UBound(gamesVector) + 1
            ReDim Preserve gamesVector(i)
        End If


        gamesVector(i) = New CLS_GAME
        gamesVector(i) = _gameToAdd



    End Sub


    '
    'Despliega la informacion de la simulacion, en la hoja de excel.
    '
    Public Sub displayResulSim(ByVal _filePath As String, ByRef infoSimula() As CLS_GAME, ByRef _ProgressBar As ProgressBar, ByRef _lblProgress As Label)
        'Variable declaration
        Dim i, rowRes, totalPart As Integer


        'Vars for openning excel File
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        _filePath = "E:\Projects\BEATS\ExcelFiles\ResultadoSimulaciones.xlsx"

        'Opening Excel Files
        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(_filePath)
        xlWorkSheet = xlWorkBook.Worksheets("ResulSimul")

        '~~> Display Excel
        xlApp.Visible = True




        'Variable Initialization        
        i = LBound(infoSimula)
        totalPart = UBound(infoSimula)
        rowRes = 2

        'Updating settings progressBar
        _ProgressBar.Minimum = 0
        _ProgressBar.Maximum = totalPart
        _ProgressBar.Value = 0



        '
        '2. Obtencion de partidos validos... filtrar partidos con rangos de 2 hrs de separacion
        '        
        While (i <= totalPart)

            'Displaying informacion general
            rowRes = i + 2
            xlWorkSheet.Cells(rowRes, "A").Value = infoSimula(i).SimulNo
            xlWorkSheet.Cells(rowRes, "B").Value = infoSimula(i).tipoApuesta
            xlWorkSheet.Cells(rowRes, "C").Value = infoSimula(i).Estrategia
            xlWorkSheet.Cells(rowRes, "D").Value = infoSimula(i).GameNo
            xlWorkSheet.Cells(rowRes, "E").Value = infoSimula(i).LeagueID
            xlWorkSheet.Cells(rowRes, "F").Value = CStr(infoSimula(i).DatePlayed)
            xlWorkSheet.Cells(rowRes, "G").Value = CStr(infoSimula(i).TimePlayed)
            xlWorkSheet.Cells(rowRes, "H").Value = infoSimula(i).partido
            xlWorkSheet.Cells(rowRes, "I").Value = infoSimula(i).Marcador
            xlWorkSheet.Cells(rowRes, "J").Value = infoSimula(i).cuota


            'Si el partido ya se ha jugado, despliega informacion del juego
            If infoSimula(i).ResLoc <> -1 Then

                xlWorkSheet.Cells(rowRes, "I").Value = infoSimula(i).Marcador
                xlWorkSheet.Cells(rowRes, "K").Value = infoSimula(i).ApGanada
                xlWorkSheet.Cells(rowRes, "L").Value = infoSimula(i).noFallos
                xlWorkSheet.Cells(rowRes, "M").Value = infoSimula(i).ApTotal
                xlWorkSheet.Cells(rowRes, "N").Value = infoSimula(i).GanNeta
                xlWorkSheet.Cells(rowRes, "O").Value = infoSimula(i).Impuesto
                xlWorkSheet.Cells(rowRes, "P").Value = infoSimula(i).GanTotal
                xlWorkSheet.Cells(rowRes, "Q").Value = infoSimula(i).ganAcum


            End If


            'Counter increase
            i = i + 1



            'Update progressBar
            If rowRes <= totalPart Then
                _ProgressBar.Value = i
                _lblProgress.Text = "Juegos Procesados: [" + CStr(i) + "] Juegos --  Porcentaje:[" + Format((i / totalPart), "0%") + "]"
            End If

        End While




        '~~> Save the file
        xlWorkBook.Save()


        'Closing ExcelFiles
        xlWorkBook.Close()
        xlApp.Quit()
        xlWorkSheet = Nothing
        xlWorkBook = Nothing
        xlApp = Nothing


    End Sub


    Private Sub updateSimulInfo(ByVal _SimulationID As String)
        'Variable declaration
        Dim adapter As SqlClient.SqlDataAdapter
        Dim sql, errorMsg As String
        Dim tblTmp As DataTable

        'Initialization variables
        errorMsg = ""
        adapter = Nothing
        tblTmp = Nothing

        'Getting DinnerRoom
        Try
            '
            '--1. Openning connection on database
            '
            If conn.State = ConnectionState.Closed Then conn.Open()

            'Getting Simulation information
            sql = "SELECT    SimulationID, TipoApuesta, RangeCoutes, Strategy, " + _
                "        Description, FilterSql, ApBase, LimNoFails, QtyIncrease, BankRoll, FiltroTipoApuesta " + _
                "        FROM  Catalog_Simulation " + _
                "        WHERE     SimulationID ='" + Trim(_SimulationID) + "'"


            'Creating connection to dabase for getting Simulation information 
            adapter = New SqlClient.SqlDataAdapter(sql, conn)
            tblTmp = New DataTable()
            adapter.Fill(tblTmp)

            If tblTmp.Rows.Count > 0 Then
                'Updating Simulation information on CLASS                 
                SimulationID = tblTmp.Rows(0)("SimulationID")
                TipoApuesta = tblTmp.Rows(0)("TipoApuesta")
                RangeCoutes = tblTmp.Rows(0)("RangeCoutes")
                Strategy = tblTmp.Rows(0)("Strategy")
                Description = tblTmp.Rows(0)("Description")
                FilterSql = tblTmp.Rows(0)("FilterSql")
                ApBase = tblTmp.Rows(0)("ApBase")
                LimNoFails = tblTmp.Rows(0)("LimNoFails")
                QtyIncrease = tblTmp.Rows(0)("QtyIncrease")
                BankRoll = tblTmp.Rows(0)("BankRoll")
                FiltroTipoApuesta = tblTmp.Rows(0)("FiltroTipoApuesta")

            Else
                errorMsg = "La simulacion:" + Trim(_SimulationID) + " NO existe en la base de datos!"
            End If

        Catch ex As Exception
            'Returning a posible error conection
            errorMsg = "Error de conexion, tratando de actualizar datos de la simulacion:" + ex.Message

        Finally
            'Closing connection and adapter            
            If Not adapter Is Nothing Then adapter.Dispose()
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try


    End Sub


    Public Function mttoSimulGames(ByVal _infoGame As CLS_GAME) As String
        'Variable declaration
        Dim result As String

        Try
            'Openning connection with database
            If conn.State = ConnectionState.Closed Then conn.Open()

            objCommand = New SqlClient.SqlCommand()

            'Setting Parameters for store procedure: spInsertSerialComponents
            With objCommand
                .CommandText = "spMttoSimulGames"
                .CommandType = CommandType.StoredProcedure
                .Connection = conn

                'Setting parameters
                With objCommand.Parameters
                    'Input parameters
                    .Add(New SqlClient.SqlParameter("@SimulationID", _infoGame.SimulationID))
                    .Add(New SqlClient.SqlParameter("@LeagueID", _infoGame.LeagueID))
                    .Add(New SqlClient.SqlParameter("@GameNo", _infoGame.GameNo))
                    .Add(New SqlClient.SqlParameter("@GameDate", _infoGame.DatePlayed))
                    .Add(New SqlClient.SqlParameter("@GameTime", _infoGame.TimePlayed))
                    .Add(New SqlClient.SqlParameter("@TeamIDLoc", _infoGame.TeamIDLoc))
                    .Add(New SqlClient.SqlParameter("@TeamIDVis", _infoGame.TeamIDVis))
                    .Add(New SqlClient.SqlParameter("@GameRes", _infoGame.Marcador))
                    .Add(New SqlClient.SqlParameter("@FilterCuote", _infoGame.CuoteFilter))
                    .Add(New SqlClient.SqlParameter("@BeatCuote", _infoGame.cuota))
                    .Add(New SqlClient.SqlParameter("@BeatResult", _infoGame.ApGanada))
                    .Add(New SqlClient.SqlParameter("@FailsNo", _infoGame.noFallos))
                    .Add(New SqlClient.SqlParameter("@BeatQty", _infoGame.ApTotal))
                    .Add(New SqlClient.SqlParameter("@ProfTot", _infoGame.GanTotal))
                    .Add(New SqlClient.SqlParameter("@ProfAcum", _infoGame.ganAcum))
                    .Add(New SqlClient.SqlParameter("@GamePlayed", _infoGame.GamePlayed))
                    .Add(New SqlClient.SqlParameter("@ActionType", 1))                  'ActionType = 1 ---> "Insert" Action
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
        mttoSimulGames = result

    End Function

    Public Function delSimulGames(ByVal _simulID As String) As String
        'Variable declaration
        Dim result As String
        Dim sql As String

        Try
            'Openning connection with database
            If conn.State = ConnectionState.Closed Then conn.Open()


            sql = "DELETE from SimulGames WHERE SimulationID='" + _simulID + "'"
            objCommand = New SqlClient.SqlCommand(sql, conn)

            'Executing store procedure for insertion
            objCommand.ExecuteNonQuery()

            result = ""

        Catch ex As Exception
            'If an error occurs with conection with the server database
            result = "Error al ejecutar la accion:" + ex.Message
        Finally
            'closing objects
            If Not objCommand Is Nothing Then objCommand.Dispose()
            If conn.State = ConnectionState.Open Then conn.Close()

        End Try


        'Returning the result
        delSimulGames = result

    End Function



End Class
