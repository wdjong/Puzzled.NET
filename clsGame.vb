Option Strict Off
Option Explicit On

Friend Class ClsGame
    Const maxPiece As Byte = 8 'Pieces in the game
    Const cellHMax As Byte = 8 'Horizontal cells in a piece
    Const cellVMax As Byte = 8 'Vertical cells in a piece
    Const boardHMax As Byte = 8 'Horizontal cells in the board
    Const boardVMax As Byte = 8 'Vertical cells in the board
    ReadOnly combination(Factorial(8)) As String 'array of combinations
    Dim currWord As String 'a combination
    Dim wordCount As Integer 'index into combination
    Public Property PuzzleName As String
    Public Property StartTime As Date
    Public Property UserBreak As Boolean 'cmdSearch during the searching says 'stop' and clicking on it interrupts the search

    Public Function AllOnBoard(ByVal pieces() As ClsPiece) As Boolean
        'Check that all pieces are fully on the board
        Dim p As Integer

        AllOnBoard = True 'lets assume for the moment
        For p = 1 To maxPiece
            If pieces(p).IsAPiece Then 'it needs to be on board
                If Not pieces(p).IsOnBoard Then
                    AllOnBoard = False 'one bad apple spoils the whole lot
                    Exit Function
                End If
            End If
        Next p
    End Function

    Public Sub CreatePieces(ByVal puzFile As String, ByVal pieces() As ClsPiece)
        'Read piece dimensions from .dat file PUZFILE
        Dim answer(cellHMax, cellVMax) As Integer
        Dim h As Integer 'h hori cell in piece
        Dim v As Integer 'v vertical cell in piece
        Dim p As Integer 'piece
        Dim a As String = "" 'value from file may not be a number

        'This is an array of number representing the location of different colours in the puzzle solution
        Try
            FileOpen(1, puzFile, OpenMode.Input)
            For h = 1 To cellHMax
                For v = 1 To cellVMax
                    Input(1, a)
                    If Not Integer.TryParse(a, answer(h, v)) Then
                        answer(h, v) = 0
                    End If
                Next v
            Next h
            FileClose(1)
        Catch ex As Exception
            MsgBox("createPieces: FileOpen: " & ex.Message)
        End Try
        Try
            For p = 1 To maxPiece 'each piece
                For h = 1 To boardHMax 'each cell
                    For v = 1 To boardVMax
                        If answer(h, v) = p Then 'If this piece is visible at this position
                            pieces(p).SetCell(h, v) '
                        Else
                            pieces(p).ResetCell(h, v)
                        End If
                    Next v
                Next h
                pieces(p).Colour = p
            Next p
            PositionOutsideSolution(pieces)
        Catch ex As Exception
            MsgBox("CreatePieces: " & ex.Message)
            Stop
        End Try
    End Sub

    Private Sub PositionOutsideSolution(pieces() As ClsPiece)
        Dim p As Integer 'piece
        Dim x As Integer 'position

        For p = 1 To maxPiece 'each piece
            If p < 5 Then x = p Else x = p + 1
            pieces(p).SetHPos(((x - 1) Mod 3) * boardHMax - cellHMax + 1) '0*8-8+1, 1*8-8+1, 2*8-8+1
            pieces(p).SetVPos(Int((x - 1) / 3) * boardVMax - cellVMax + 1) '0*8-8+1,x=4:1*8-8+1,x=7:2*8-8+1
            'Debug.Print("p: " & p & " bh " & pieces(p).GetHPos() & " bv " & pieces(p).GetVPos())
            pieces(p).Centre()
            If My.Settings.Level = 2 Then
                Dim d As Integer
                For d = 1 To Int(Rnd() * 4 + 1)
                    pieces(p).RotateL()
                Next d
            End If
            If My.Settings.Level = 3 Then
                Dim d As Integer
                For d = 1 To Int(Rnd() * 2 + 1)
                    pieces(p).Flip()
                Next d
            End If
        Next p
    End Sub

    Public Sub DrawIt(ByVal aBoard As ClsBoard, ByVal pieces() As ClsPiece, ByVal g As System.Drawing.Graphics)
        Dim p As Integer 'piece

        aBoard.DrawIt(g)
        Try
            For p = 1 To maxPiece
                pieces(p).DrawIt(g)
            Next p
        Catch ex As Exception
            MsgBox("drawIt: " & ex.Message & " " & ex.StackTrace)
        End Try
    End Sub

    Public Sub EndGame()
        Dim timetaken As Single
        Dim CSVText As String

        timetaken = DateDiff(Microsoft.VisualBasic.DateInterval.Second, StartTime, Now)
        MsgBox("Well done. That only took " & timetaken & " seconds.")
        CSVText = DateTime.Now.ToString & ", " & PuzzleName & ", " & timetaken & vbNewLine
        My.Computer.FileSystem.WriteAllText("PuzzleHistory.csv", CSVText, True)
    End Sub

    Public Function Factorial(ByVal Factor As Byte) As Object
        '******************************************************
        'PURPOSE:
        'SOLVE FACTORIALS: N!, or N*(N-1)*(N-2)*...*2*1
        'Parameter: N
        'Returns: Factorial

        'EXAMPLE: 
        'MsgBox Factorial(6) 
        'Will display 720 because 6 * 5 * 4 * 3 * 2 * 1 = 720

        'NOTE: Overflow will occur with any parameter over 170
        '******************************************************

        On Error GoTo ErrorHandler

        If Factor = 0 Then
            Factorial = 1
        Else
            Factorial = Factor * Factorial(Factor - 1)
        End If
        Exit Function

ErrorHandler:
        If Err.Number = 6 Then 'oveflow
            Err.Raise(Err.Number, , "Overflow: Number passed to function was too large")

        Else 'unknown reason for error, shouldn't occur
            Err.Raise(Err.Number, , Err.Description)
        End If
    End Function

    Public Sub FindCombinations()
        'Find all combinations and put them into tblCombination
        Dim n As Short 'each letter
        Dim wordLen As Short 'length of word

        Try
            UserBreak = False 'true if escape pressed
            currWord = "12345678" 'e.g. "tracehent"
            wordLen = Len(currWord)
            wordCount = 0
            My.Forms.FrmPuzzle.ProgressBar1.Visible = True
            My.Forms.FrmPuzzle.ProgressBar1.Minimum = 0
            My.Forms.FrmPuzzle.ProgressBar1.Maximum = Factorial(wordLen)
            My.Forms.FrmPuzzle.ProgressBar1.Value = wordCount
            For n = 1 To wordLen 'for each letter in the word
                currWord = ShiftRight(currWord) 'move the last letter to the beginning
                Recombine(wordLen - 1, wordLen) 'in which words are added to tblCombine
                My.Forms.FrmPuzzle.ProgressBar1.Value = wordCount
                Application.DoEvents() 'check for a break
            Next n
            My.Forms.FrmPuzzle.ProgressBar1.Visible = False
        Catch ex As Exception
            MsgBox("doStage1: " & ex.Message)
        End Try
    End Sub

    Public Sub GeneratePieces2()
        Dim answer(cellHMax, cellVMax) As Integer 'create a grid of cells
        Dim h As Integer 'h horizongal position
        Dim v As Integer 'v vertical position
        Dim aBoard As ClsBoard = New ClsBoard()
        Dim pieces(maxPiece) As ClsPiece
        Dim outputString As String = ""
        Dim p As Integer
        Dim a As String
        Dim f As Integer

        For p = 1 To maxPiece
            pieces(p) = New ClsPiece
        Next p

        SetStartingSpots(aBoard, pieces)
        SetPiecesOnBoard(aBoard, pieces)
        While Not IsBoardFull(aBoard) And f < 1000 'is occupied
            SetAdjoining(aBoard, pieces)
            SetPiecesOnBoard(aBoard, pieces)
            f += 1
        End While

        'Create output file
        For v = 1 To cellVMax
            For h = 1 To cellHMax
                a = "0"
                For p = 1 To maxPiece
                    If pieces(p).GetCell(h, v) Then
                        a = p.ToString
                    End If
                Next p
                outputString += a
                If h <> cellHMax Then
                    outputString += ","
                Else
                    If v <> cellVMax Then
                        outputString += vbNewLine
                    End If
                End If
            Next h
        Next v
        My.Computer.FileSystem.WriteAllText("walter.dat", outputString, False, System.Text.ASCIIEncoding.ASCII)
    End Sub

    Public Sub Init(ByVal aBoard As ClsBoard, ByVal pieces() As ClsPiece, ByVal puzFile As String)
        'Initialise the board and it's pieces
        Dim p As Integer

        PuzzleName = puzFile
        aBoard.Init()
        For p = 1 To maxPiece
            pieces(p).Init(p, p)
        Next
        CreatePieces(puzFile, pieces)
    End Sub

    Public Function IsBoardFull(ByVal aBoard As ClsBoard) As Boolean
        'Check if the board (solution part) is full (puzzle complete)
        'Clear board, then set occupied cells for each piece, then check if any unoccupied
        Dim h As Integer
        Dim v As Integer

        For h = 1 To cellHMax 'Scan board for unnoccupied positions
            For v = 1 To cellVMax
                If Not aBoard.GetPosition(h, v) Then 'unoccupied board space
                    IsBoardFull = False
                    Exit Function
                End If
            Next v
        Next h
        IsBoardFull = True
    End Function

    Public Function GetVacancyToPoint(ByVal aBoard As ClsBoard, h As Integer, v As Integer) As Integer
        'Find unoccupied spots
        Dim h1 As Integer
        Dim v1 As Integer

        For h1 = 1 To h 'Scan board for unnoccupied positions
            For v1 = 1 To v
                If Not aBoard.GetPosition(h1, v1) Then 'unoccupied board space
                    GetVacancyToPoint += 1
                End If
            Next v1
        Next h1
    End Function

    Private Sub LayThemOut(order As String, aBoard As ClsBoard, pieces() As ClsPiece)
        Dim i As Integer
        Dim pIndex() As Char = order.ToCharArray

        PositionOutsideSolution(pieces)
        For i = 0 To order.Length - 1
            Select Case My.Settings.Level
                Case 1
                    If Not PlaceNext(aBoard, pieces, Integer.Parse(pIndex(i))) Then
                        Exit For 'couldn't place part without overlapping. give up
                    End If
                Case 2
                    If Not PlaceNext2(aBoard, pieces, Integer.Parse(pIndex(i))) Then
                        Exit For 'couldn't place part without overlapping. give up
                    End If
                Case 3
                    If Not PlaceNext3(aBoard, pieces, Integer.Parse(pIndex(i))) Then
                        Exit For 'couldn't place part without overlapping. give up
                    End If
            End Select
            SetPiecesOnBoard(aBoard, pieces)
        Next
    End Sub

    Private Function Overlapping(aboard As ClsBoard, pieces() As ClsPiece, p As Integer) As Boolean
        'Dim op As Integer 'other piece index
        Dim cv As Integer 'cell v
        Dim ch As Integer 'cell h

        Overlapping = False
        For cv = 1 To cellVMax
            For ch = 1 To cellHMax
                If aboard.GetPosition(pieces(p).GetHPos - 1 + ch, pieces(p).GetVPos - 1 + cv) And pieces(p).GetCell(ch, cv) Then
                    Overlapping = True
                End If
            Next
        Next
        'For op = 1 To maxPiece
        '    If op <> p Then 'ignore self
        '        If pieces(op).LeftExtremity <= pieces(p).LeftExtremity And pieces(op).RightExtremity >= pieces(p).LeftExtremity _
        '            Or pieces(op).RightExtremity >= pieces(p).RightExtremity And pieces(op).LeftExtremity <= pieces(p).RightExtremity _
        '            Or pieces(op).TopExtremity <= pieces(p).TopExtremity And pieces(op).BottomExtremity >= pieces(p).TopExtremity _
        '            Or pieces(op).BottomExtremity >= pieces(p).BottomExtremity And pieces(op).TopExtremity <= pieces(p).BottomExtremity Then 'maybe overlapping (but not necessarally


        '        End If
        '    End If
        'Next op
    End Function

    Private Function PlaceNext(aBoard As ClsBoard, pieces() As ClsPiece, p As Integer) As Boolean
        'Level 1 placement no twists or turns
        Dim bh As Integer
        Dim bv As Integer
        Dim overlap As Boolean 'assume it overlaps
        Dim originalBH As Integer
        Dim originalBV As Integer 'for undo

        originalBH = pieces(p).GetHPos() 'prepare to replace piece outside solution if it ends up overlapping
        originalBV = pieces(p).GetVPos()

        bv = pieces(p).TopStart 'e.g. top left visible cell in piece = 3, 3 bh and bv should be -1, -1 (because board is from 1,1 to 8,8
        pieces(p).SetVPos(bv)
        bh = pieces(p).LeftStart
        pieces(p).SetHPos(bh)
        overlap = Overlapping(aBoard, pieces, p) 'preceding pieces have been set on the board but not this piece

        'My.Forms.FrmPuzzle.Game.Invalidate()
        'Threading.Thread.Sleep(2000)
        'Application.DoEvents()

        While bh < pieces(p).LeftEnd And overlap
            bh += 1
            pieces(p).SetHPos(bh)
            overlap = Overlapping(aBoard, pieces, p)

            'My.Forms.FrmPuzzle.Game.Invalidate()
            'Threading.Thread.Sleep(200)
            'Application.DoEvents()

        End While 'bh

        While bv < pieces(p).TopEnd And overlap '<= so it does all h pos in last row
            'If Not overlap Or bv = pieces(p).TopEnd Then Exit While
            bv += 1
            pieces(p).SetVPos(bv)
            bh = pieces(p).LeftStart
            pieces(p).SetHPos(bh)
            overlap = Overlapping(aBoard, pieces, p)

            While bh < pieces(p).LeftEnd And overlap
                bh += 1
                pieces(p).SetHPos(bh)
                overlap = Overlapping(aBoard, pieces, p)

                'My.Forms.FrmPuzzle.Game.Invalidate()
                'Threading.Thread.Sleep(200)
                'Application.DoEvents()

            End While 'bh

        End While 'bv

        If overlap Then
            PlaceNext = False
            pieces(p).SetHPos(originalBH)
            pieces(p).SetVPos(originalBV)

            'My.Forms.FrmPuzzle.Game.Invalidate()
            'Threading.Thread.Sleep(2000)
            'Application.DoEvents()
        Else
            PlaceNext = True
        End If
    End Function

    Private Function PlaceNext2(aBoard As ClsBoard, pieces() As ClsPiece, p As Integer) As Boolean
        Dim d As Integer 'rotation
        Dim bBoard As New ClsBoard()
        Dim dBest As Integer 'best rotation is one with least vacancy to left
        Dim v As Integer
        Dim dBestVacancy As Integer = boardHMax * boardVMax

        PlaceNext2 = False
        For d = 1 To 4 'For Each rotation, 
            pieces(p).RotateL()
            If PlaceNext(aBoard, pieces, p) Then
                aBoard.Copy(bBoard) 'copy board
                SetPiecesOnBoard(bBoard, pieces) 'add this piece to it
                v = GetVacancyToPoint(bBoard, pieces(p).RightExtremity + pieces(p).GetHPos - 1, pieces(p).BottomExtremity + pieces(p).GetVPos - 1)
                If v < dBestVacancy Then
                    dBestVacancy = v
                    dBest = d
                End If
            End If
            'If aBoard.OccupiedCount > 50 Then
            'My.Forms.FrmPuzzle.Game.Invalidate()
            'Threading.Thread.Sleep(1000)
            'Application.DoEvents()
            'End If
        Next d
        'Use best
        For d = 1 To dBest
            pieces(p).RotateL() 'For each flip,
            PlaceNext2 = PlaceNext(aBoard, pieces, p)
        Next
    End Function

    Private Function PlaceNext3(aboard As ClsBoard, pieces() As ClsPiece, p As Integer) As Boolean
        Dim d As Integer 'rotation
        Dim f As Integer 'flip
        Dim bBoard As New ClsBoard()
        Dim dBest As Integer 'best rotation is one with least vacancy to left
        Dim fBest As Integer '
        Dim v As Integer
        Dim dBestVacancy As Integer = boardHMax * boardVMax

        For d = 1 To 4 'For Each rotation, 
            pieces(p).RotateL()
            For f = 1 To 2 'For Each flip,
                pieces(p).Flip()
                If PlaceNext(aboard, pieces, p) Then
                    aboard.Copy(bBoard) 'copy board
                    SetPiecesOnBoard(bBoard, pieces) 'add this piece to it
                    v = GetVacancyToPoint(bBoard, pieces(p).RightExtremity + pieces(p).GetHPos - 1, pieces(p).BottomExtremity + pieces(p).GetVPos - 1)
                    If v < dBestVacancy Then
                        dBestVacancy = v
                        dBest = d
                        fBest = f
                    End If
                End If
            Next f
        Next d
        'Use best
        For d = 1 To dBest
            pieces(p).RotateL() 'For each flip,
        Next
        For f = 1 To fBest
            pieces(p).Flip() 'For each flip,
        Next
        PlaceNext3 = PlaceNext(aboard, pieces, p)
    End Function

    Private Sub NextDirection(ByRef hd As Integer, ByRef vd As Integer)
        'move to next direction clockwise
        Select Case hd
            Case -1 'westish
                Select Case vd
                    Case 0 'w -> n
                        hd = 0
                        vd = -1
                End Select
            Case 0 'up & down
                Select Case vd
                    Case -1 'n -> e
                        hd = 1
                        vd = 0
                    Case 1 's -> w
                        hd = -1
                        vd = 0
                End Select
            Case 1 'eastish
                Select Case vd
                    Case 0 'e -> s
                        hd = 0
                        vd = 1
                End Select
        End Select
    End Sub

    Private Sub NextDirection8(ByRef hd As Integer, ByRef vd As Integer)
        'move to next direction clockwise of eight directions (Not currently used in automatic piece creation because of the strangish shapes generated
        Select Case hd
            Case -1 'westish
                Select Case vd
                    Case -1 'nw -> n
                        hd = 0
                    Case 0 'w -> nw
                        vd = -1
                    Case 1 'sw -> w
                        vd = 0
                End Select
            Case 0 'up & down
                Select Case vd
                    Case -1 'n -> ne
                        hd = 1
                    Case 1 's -> sw
                        hd = -1
                End Select
            Case 1 'eastish
                Select Case vd
                    Case -1 'ne -> e
                        vd = 0
                    Case 0 'e -> se
                        vd = 1
                    Case 1 'se -> s
                        hd = 0
                End Select
        End Select
    End Sub

    Public Sub Recombine(ByRef iLevel As Short, ByRef wordLen As Short)
        'Called from dostage1 and then recursively to get each combination of word
        'currword "abc" recombine 2, 3 
        'currword=a cb recombine 1, 3 addToDB "a cb"
        'currword=a bc recombine 1, 3 addToDB "a bc"
        Dim n As Short 'change the right n letters

        Try
            If iLevel > 1 And Not UserBreak Then 'keep going deeper
                For n = 1 To iLevel 'for each of the iLevel rightmost letters
                    currWord = Left(currWord, wordLen - iLevel) & ShiftRight(Right(currWord, iLevel))
                    Recombine(iLevel - 1, wordLen)
                Next n
            Else 'iLevel = 1
                wordCount += 1
                combination(wordCount) = currWord
                'Debug.Print(currWord)
            End If
        Catch ex As Exception
            If ex.HResult <> -2147467259 Then 'ignore error relating to inserting duplicates
                MsgBox("recombine: " & ex.Message)
            End If
        End Try
    End Sub

    Private Sub SetAdjoining(aBoard As ClsBoard, pieces As ClsPiece())
        'used by GeneratePieces2() when generating a puzzle
        Dim p As Integer 'pieces
        Dim i As Integer 'positions index
        Dim h As Integer 'horiz
        Dim hd As Integer 'positions index
        Dim v As Integer 'verti
        Dim vd As Integer 'positions index
        Dim boardFull As Boolean
        Dim d As Integer 'directional possibilities (i.e. 8)

        For p = 1 To maxPiece
            i = 0
            Dim positions(cellHMax * cellVMax - 1, 1) As Integer
            For h = 1 To cellHMax ' Get a list of positions occupied by this piece
                For v = 1 To cellVMax
                    If pieces(p).GetCell(h, v) Then
                        positions(i, 0) = h
                        positions(i, 1) = v
                        i += 1
                    End If
                Next v
            Next h
            Randomize()
            i = Int(i * Rnd()) 'pick a point in the shape 0 - (n-1)
            Randomize()
            hd = CInt(Int(3 * Rnd() - 1)) 'horizontal direction -1 0 1
            Randomize()
            vd = CInt(Int(3 * Rnd() - 1)) 'vertical direction -1 0 1
            While (hd = 0 And vd = 0) Or (hd <> 0 And vd <> 0)  'pick a direction they can't both be 0 and they can't both be non-zero
                Randomize()
                hd = CInt(Int(3 * Rnd() - 1)) 'horizontal direction -1 0 1
                Randomize()
                vd = CInt(Int(3 * Rnd() - 1)) 'vertical direction -1 0 1
            End While
            For d = 1 To 4 'for each direction (starting from first
                h = positions(i, 0) + hd 'check adjoining spot
                v = positions(i, 1) + vd
                While pieces(p).GetCell(h, v) 'while in same piece
                    h += hd 'keep going
                    v += vd
                End While
                If aBoard.GetPosition(h, v) Then 'there's something else there
                    'Debug.Print("Found other piece at h" & h & " v" & v) 'next
                ElseIf Not aBoard.IsOnBoard(h, v) Then
                    'Debug.Print("Found edge of board")
                Else 'you've found somewhere to grow
                    pieces(p).SetCell(h, v)
                    pieces(p).SetHPos(1)
                    pieces(p).SetVPos(1)
                    SetPiecesOnBoard(aBoard, pieces)
                    boardFull = IsBoardFull(aBoard)
                    Exit For
                End If
                NextDirection(hd, vd) 'passed by ref so they can change
            Next d
            If boardFull Then
                Exit For
            End If
        Next p
    End Sub

    Public Sub SetPiecesOnBoard(aBoard As ClsBoard, pieces() As ClsPiece)
        'Set pieces on board for checking occupancy
        Dim p As Integer 'piece
        Dim h As Integer
        Dim v As Integer

        aBoard.Clear() 'Clear board
        For p = 1 To maxPiece 'Determine occupied positions
            For h = pieces(p).LeftExtremity To pieces(p).RightExtremity
                For v = pieces(p).TopExtremity To pieces(p).BottomExtremity
                    If pieces(p).GetCell(h, v) Then
                        aBoard.SetPosition(pieces(p).GetHPos - 1 + h, pieces(p).GetVPos - 1 + v)
                    End If
                Next v
            Next h
        Next p
    End Sub

    Private Sub SetStartingSpots(aBoard As ClsBoard, pieces As ClsPiece())
        Dim p As Integer
        Dim h As Integer
        Dim v As Integer
        Dim f As Integer
        Dim boardFull As Boolean

        For p = 1 To maxPiece
            Randomize()
            h = CInt(Int(boardHMax * Rnd() + 1))
            Randomize()
            v = CInt(Int(boardVMax * Rnd() + 1))
            While aBoard.GetPosition(h, v) And f < 1000 'is occupied
                Randomize()
                h = CInt(Int(boardHMax * Rnd() + 1))
                Randomize()
                v = CInt(Int(boardVMax * Rnd() + 1))
                f += 1
            End While
            pieces(p).Colour = p
            pieces(p).SetCell(h, v)
            pieces(p).SetHPos(1)
            pieces(p).SetVPos(1)
            SetPiecesOnBoard(aBoard, pieces)
            boardFull = IsBoardFull(aBoard) 'this is currently the only place that updates board occupancy
        Next p
    End Sub

    Public Function ShiftRight(ByVal a As String) As String
        'Called by recombine
        'move all the letters of the 'word' right
        'move the last letter to the start of the 'word'
        ShiftRight = Right(a, 1) & Left(a, Len(a) - 1)
    End Function

    Public Sub ShowSolution(ByVal puzFile As String, ByVal pieces() As ClsPiece)
        'Read piece dimensions from .dat file PUZFILE
        Dim answer(cellHMax, cellVMax) As Integer
        Dim h As Integer 'h hori cell in piece
        Dim v As Integer 'v vertical cell in piece
        Dim p As Integer 'piece
        Dim a As String = "" 'value from file may not be a number

        'This is an array of number representing the location of different colours in the puzzle solution
        Try
            FileOpen(1, puzFile, OpenMode.Input)
            For h = 1 To cellHMax
                For v = 1 To cellVMax
                    Input(1, a)
                    If Not Integer.TryParse(a, answer(h, v)) Then
                        answer(h, v) = 0
                    End If
                Next v
            Next h
            FileClose(1)
        Catch ex As Exception
            MsgBox("createPieces: FileOpen: " & ex.Message)
        End Try
        Try
            For p = 1 To maxPiece 'each piece
                For h = 1 To boardHMax 'each cell
                    For v = 1 To boardVMax
                        If answer(h, v) = p Then 'If this piece is visible at this position
                            pieces(p).SetCell(h, v) '
                        Else
                            pieces(p).ResetCell(h, v)
                        End If
                    Next v
                Next h
                pieces(p).Colour = p
                pieces(p).SetHPos(1) '0*8-8+1, 1*8-8+1, 2*8-8+1
                pieces(p).SetVPos(1) '0*8-8+1,x=4:1*8-8+1,x=7:2*8-8+1
            Next p
        Catch ex As Exception
            MsgBox("CreatePieces: " & ex.Message)
            Stop
        End Try
    End Sub

    Public Sub SolvePuzzle(aBoard As ClsBoard, pieces As ClsPiece())
        Dim p As Integer
        Dim s As Integer 'solutions
        Dim v As Integer
        Dim vTimeOut As Integer 'just by way of giving users visual feedback
        Dim dBestVacancy As Integer = boardHMax * boardVMax

        StartTime = Now
        'LayThemOut("12634587", aBoard, pieces)
        My.Forms.FrmPuzzle.StatusBar1.Text = "Find combinations"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        FindCombinations() 'fills combination array
        My.Forms.FrmPuzzle.StatusBar1.Text = "Trying combinations"
        Application.DoEvents()
        My.Forms.FrmPuzzle.ProgressBar1.Visible = True
        My.Forms.FrmPuzzle.ProgressBar1.Minimum = 0
        My.Forms.FrmPuzzle.ProgressBar1.Maximum = combination.Length - 1
        My.Forms.FrmPuzzle.ProgressBar1.Value = p
        For p = 1 To combination.Length - 1 'For each permutation of pieces
            'Try To lay them out left to right, top to bottom
            LayThemOut(combination(p), aBoard, pieces)
            SetPiecesOnBoard(aBoard, pieces)
            If IsBoardFull(aBoard) Then
                My.Forms.FrmPuzzle.Game.Invalidate()
                Application.DoEvents()
                'Debug.Print(combination(p))
                Dim timetaken As Single
                timetaken = DateDiff(Microsoft.VisualBasic.DateInterval.Second, StartTime, Now)
                My.Forms.FrmPuzzle.StatusBar1.Text = "Solution found: " & timetaken & " seconds."
                s += 1 'or count solutions (but as yet they're not really unique solutions
                Exit For
            Else
                v = GetVacancyToPoint(aBoard, 8, 8)
                'Debug.Print("v: " & v & "   dBestVacancy: " & dBestVacancy)
                If v < dBestVacancy Then
                    dBestVacancy = v
                    My.Forms.FrmPuzzle.Game.Invalidate()
                    Threading.Thread.Sleep(200)
                    Application.DoEvents()
                    vTimeOut = 0
                End If
                vTimeOut += 1
                If vTimeOut > 2000 Then
                    vTimeOut = 0
                    dBestVacancy = boardHMax * boardVMax
                End If
                My.Forms.FrmPuzzle.ProgressBar1.Value = p
            End If
        Next
        If s = 0 Then
            MsgBox("I couldn't find the solution...")
        End If
        My.Forms.FrmPuzzle.ProgressBar1.Visible = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Public Sub SolvePuzzle1(aBoard As ClsBoard, pieces As ClsPiece()) 'NOT IN USE
        'Unfortunally if each piece can be in 30ish positions -> 30^8 is too big a number to iterate through all possibilities and that's without flips and twists
        TryPieces(aBoard, pieces, 1)
    End Sub

    Public Sub TryPieces(aBoard As ClsBoard, pieces As ClsPiece(), aPiece As Integer)
        'NOT IN USE 
        'Laboriously try every piece in every position
        Dim bh As Integer 'board horizontal position
        Dim bv As Integer 'board vertical position
        Static a As Integer 'attempts
        Static s As Integer 'solutions
        'Dim d As Integer 'direction of rotation
        'Dim f As Integer 'flip sides

        If aPiece > maxPiece Then 'finished
            Exit Sub
        Else
            For bv = pieces(aPiece).TopStart To pieces(aPiece).TopEnd
                For bh = pieces(aPiece).LeftStart To pieces(aPiece).LeftEnd
                    pieces(aPiece).SetHPos(bh)
                    pieces(aPiece).SetVPos(bv)
                    Application.DoEvents()
                    If aPiece = 5 Then
                        'Debug.Print(aPiece & " " & bv & " " & bh & " " & a & " " & s)
                        My.Forms.FrmPuzzle.Game.Invalidate()
                        Threading.Thread.Sleep(20)
                    End If
                    a += 1
                    If pieces(aPiece).IsOnBoard Then
                        SetPiecesOnBoard(aBoard, pieces)
                        If IsBoardFull(aBoard) Then
                            'Debug.Print(aPiece & " " & bv & " " & bh)
                            My.Forms.FrmPuzzle.Game.Invalidate()
                            Application.DoEvents()
                            MessageBox.Show("Solution found")
                            s += 1
                        End If
                    End If

                    TryPieces(aBoard, pieces, aPiece + 1)

                Next bh
            Next bv
        End If

    End Sub

    Public Function WhichPiece(ByVal pieces() As ClsPiece, ByRef h As Integer, ByRef v As Integer) As Integer
        'Given a co-ordinate pointed to with the mouse
        'Return the piece pointed to
        Dim p As Integer 'piece

        'Because pieces are drawn from 1 to maxpiece
        'The bigger numbered piece is on top and therefore
        'we should assume the top one is what's intended first
        For p = maxPiece To 1 Step -1 'check each piece
            If pieces(p).IsClicked(h, v) Then
                WhichPiece = p 'first piece is on top
                Exit Function
            End If
        Next p
        WhichPiece = 0 'no piece clicked
    End Function

End Class