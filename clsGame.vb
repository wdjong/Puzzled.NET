Option Strict Off
Option Explicit On
Imports System.Net

Friend Class ClsGame
    Const MAXPIECE As Byte = 8 'Pieces in the game
    Const MAXHCELL As Byte = 8 'Horizontal cells in a piece
    Const MAXVCELL As Byte = 8 'Vertical cells in a piece
    Const HSize As Byte = 8 'Horizontal cells in the board
    Const VSize As Byte = 8 'Vertical cells in the board

    Public Property PuzzleName As String
    Public Property StartTime As Date

    Public Sub EndGame()
        Dim timetaken As Single
        Dim CSVText As String

        timetaken = DateDiff(Microsoft.VisualBasic.DateInterval.Second, StartTime, Now)
        MsgBox("Well done. That only took " & timetaken & " seconds.")
        CSVText = DateTime.Now.ToString & ", " & PuzzleName & ", " & timetaken & vbNewLine
        My.Computer.FileSystem.WriteAllText("PuzzleHistory.csv", CSVText, True)
    End Sub

    Public Function AllOnBoard(ByVal aPiece() As ClsPiece) As Boolean
        'Check that all pieces are fully on the board
        Dim p As Integer

        AllOnBoard = True 'lets assume for the moment
        For p = 1 To MAXPIECE
            If aPiece(p).IsAPiece Then 'it needs to be on board
                If Not aPiece(p).IsOnBoard Then
                    AllOnBoard = False 'one bad apple spoils the whole lot
                    Exit Function
                End If
            End If
        Next p
    End Function

    Public Sub CreatePieces(ByVal puzFile As String, ByVal aPiece() As ClsPiece)
        'Read piece dimensions from .dat file PUZFILE
        Dim answer(MAXHCELL, MAXVCELL) As Integer
        Dim h As Integer 'h hori cell in piece
        Dim v As Integer 'v vertical cell in piece
        Dim p As Integer 'piece
        Dim x As Integer 'position
        Dim a As String = "" 'value from file may not be a number

        'This is an array of number representing the location of different colours in the puzzle solution
        Try
            FileOpen(1, puzFile, OpenMode.Input)
            For h = 1 To MAXHCELL
                For v = 1 To MAXVCELL
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
            For p = 1 To MAXPIECE 'each piece
                For h = 1 To HSize 'each cell
                    For v = 1 To VSize
                        If answer(h, v) = p Then 'If this piece is visible at this position
                            aPiece(p).SetCell(h, v) '
                        Else
                            aPiece(p).ResetCell(h, v)
                        End If
                    Next v
                Next h
                aPiece(p).Colour = p
                'This is about positioning the pieces around the board
                'There a 9 spots to put pieces but nothing will go in the 5th (so skip it)
                If p < 5 Then x = p Else x = p + 1
                aPiece(p).SetHPos(((x - 1) Mod 3) * HSize - MAXHCELL + 1) '0*8-8+1, 1*8-8+1, 2*8-8+1
                aPiece(p).SetVPos(Int((x - 1) / 3) * VSize - MAXVCELL + 1) '0*8-8+1,x=4:1*8-8+1,x=7:2*8-8+1
                aPiece(p).Centre()
                If My.Settings.Level = 2 Then
                    Dim d As Integer
                    For d = 1 To Int(Rnd() * 4 + 1)
                        aPiece(p).RotateL()
                    Next d
                End If
                If My.Settings.Level = 3 Then
                    Dim d As Integer
                    For d = 1 To Int(Rnd() * 2 + 1)
                        aPiece(p).Flip()
                    Next d
                End If
            Next p
        Catch ex As Exception
            MsgBox("CreatePieces: " & ex.Message)
            Stop
        End Try
    End Sub

    Public Sub ShowSolution(ByVal puzFile As String, ByVal aPiece() As ClsPiece)
        'Read piece dimensions from .dat file PUZFILE
        Dim answer(MAXHCELL, MAXVCELL) As Integer
        Dim h As Integer 'h hori cell in piece
        Dim v As Integer 'v vertical cell in piece
        Dim p As Integer 'piece
        Dim a As String = "" 'value from file may not be a number

        'This is an array of number representing the location of different colours in the puzzle solution
        Try
            FileOpen(1, puzFile, OpenMode.Input)
            For h = 1 To MAXHCELL
                For v = 1 To MAXVCELL
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
            For p = 1 To MAXPIECE 'each piece
                For h = 1 To HSize 'each cell
                    For v = 1 To VSize
                        If answer(h, v) = p Then 'If this piece is visible at this position
                            aPiece(p).SetCell(h, v) '
                        Else
                            aPiece(p).ResetCell(h, v)
                        End If
                    Next v
                Next h
                aPiece(p).Colour = p
                aPiece(p).SetHPos(1) '0*8-8+1, 1*8-8+1, 2*8-8+1
                aPiece(p).SetVPos(1) '0*8-8+1,x=4:1*8-8+1,x=7:2*8-8+1
            Next p
        Catch ex As Exception
            MsgBox("CreatePieces: " & ex.Message)
            Stop
        End Try
    End Sub

    Public Sub DrawIt(ByVal aBoard As ClsBoard, ByVal aPiece() As ClsPiece, ByVal g As System.Drawing.Graphics)
        Dim p As Integer 'piece

        aBoard.DrawIt(g)
        Try
            For p = 1 To MAXPIECE
                aPiece(p).DrawIt(g)
            Next p
        Catch ex As Exception
            MsgBox("drawIt: " & ex.Message & " " & ex.StackTrace)
        End Try
    End Sub

    Public Sub Init(ByVal aBoard As ClsBoard, ByVal pieces() As ClsPiece, ByVal puzFile As String)
        'Initialise the board and it's pieces
        Dim p As Integer

        PuzzleName = puzFile
        aBoard.Init()
        For p = 1 To MAXPIECE
            pieces(p).Init(p, p)
        Next
        CreatePieces(puzFile, pieces)
    End Sub

    Public Function IsBoardFull(ByVal aBoard As ClsBoard, ByVal aPiece() As ClsPiece) As Boolean
        'Check if the board (solution part) is full (puzzle complete)
        'Clear board, then set occupied cells for each piece, then check if any unoccupied
        Dim p As Integer 'piece
        Dim h As Integer
        Dim v As Integer

        aBoard.Clear() 'Clear board
        For p = 1 To MAXPIECE 'Determine occupied positions
            For h = aPiece(p).LeftExtremity To aPiece(p).RightExtremity
                For v = aPiece(p).TopExtremity To aPiece(p).BottomExtremity
                    If aPiece(p).GetCell(h, v) Then
                        aBoard.SetPosition(aPiece(p).GetHPos - 1 + h, aPiece(p).GetVPos - 1 + v)
                    End If
                Next v
            Next h
        Next p
        For h = 1 To MAXHCELL 'Scan board for unnoccupied positions
            For v = 1 To MAXVCELL
                If Not aBoard.GetPosition(h, v) Then 'unoccupied board space
                    IsBoardFull = False
                    Exit Function
                End If
            Next v
        Next h
        IsBoardFull = True
    End Function

    Public Function WhichPiece(ByVal aPiece() As ClsPiece, ByRef h As Integer, ByRef v As Integer) As Integer
        'Given a co-ordinate pointed to with the mouse
        'Return the piece pointed to
        Dim p As Integer 'piece

        'Because pieces are drawn from 1 to maxpiece
        'The bigger numbered piece is on top and therefore
        'we should assume the top one is what's intended first
        For p = MAXPIECE To 1 Step -1 'check each piece
            If aPiece(p).IsClicked(h, v) Then
                WhichPiece = p 'first piece is on top
                Exit Function
            End If
        Next p
        WhichPiece = 0 'no piece clicked
    End Function

    Public Sub GeneratePieces()
        Dim answer(MAXHCELL, MAXVCELL) As Integer
        Dim h As Integer 'h hori cell in piece
        Dim v As Integer 'v vertical cell in piece
        Dim outputString As String = ""

        Randomize()
        Try
            For v = 1 To MAXVCELL
                For h = 1 To MAXHCELL
                    answer(h, v) = CInt(Int(MAXPIECE * Rnd() + 1))
                Next h
            Next v

            'Create output file
            For v = 1 To MAXVCELL
                For h = 1 To MAXHCELL
                    outputString += answer(h, v).ToString '
                    If h <> MAXHCELL Then
                        outputString += ","
                    Else
                        If v <> MAXVCELL Then
                            outputString += vbNewLine
                        End If
                    End If
                Next h
            Next v
            My.Computer.FileSystem.WriteAllText("walter.dat", outputString, False, System.Text.ASCIIEncoding.ASCII)
        Catch ex As Exception
            MsgBox("GeneratePieces: " & ex.Message)
        End Try
    End Sub

    Public Sub GeneratePieces2()
        Dim answer(MAXHCELL, MAXVCELL) As Integer 'create a grid of cells
        Dim h As Integer 'h horizongal position
        Dim v As Integer 'v vertical position
        Dim aBoard As ClsBoard = New ClsBoard()
        Dim pieces(MAXPIECE) As ClsPiece
        Dim outputString As String = ""
        Dim p As Integer
        Dim a As String
        Dim f As Integer

        For p = 1 To MAXPIECE
            pieces(p) = New ClsPiece
        Next p

        SetStartingSpots(aBoard, pieces)
        While Not IsBoardFull(aBoard, pieces) And f < 1000 'is occupied
            SetAdjoining(aBoard, pieces)
            f += 1
        End While

        'Create output file
        For v = 1 To MAXVCELL
            For h = 1 To MAXHCELL
                a = "0"
                For p = 1 To MAXPIECE
                    If pieces(p).GetCell(h, v) Then
                        a = p.ToString
                    End If
                Next p
                outputString += a
                If h <> MAXHCELL Then
                    outputString += ","
                Else
                    If v <> MAXVCELL Then
                        outputString += vbNewLine
                    End If
                End If
            Next h
        Next v
        My.Computer.FileSystem.WriteAllText("walter.dat", outputString, False, System.Text.ASCIIEncoding.ASCII)
    End Sub

    Private Sub SetStartingSpots(aBoard As ClsBoard, pieces As ClsPiece())
        Dim p As Integer
        Dim h As Integer
        Dim v As Integer
        Dim f As Integer
        Dim boardFull As Boolean

        For p = 1 To MAXPIECE
            Randomize()
            h = CInt(Int(HSize * Rnd() + 1))
            Randomize()
            v = CInt(Int(VSize * Rnd() + 1))
            While aBoard.GetPosition(h, v) And f < 1000 'is occupied
                Randomize()
                h = CInt(Int(HSize * Rnd() + 1))
                Randomize()
                v = CInt(Int(VSize * Rnd() + 1))
                f += 1
            End While
            pieces(p).Colour = p
            pieces(p).SetCell(h, v)
            pieces(p).SetHPos(1)
            pieces(p).SetVPos(1)
            boardFull = IsBoardFull(aBoard, pieces) 'this is currently the only place that updates board occupancy
        Next p
    End Sub

    Private Sub SetAdjoining(aBoard As ClsBoard, pieces As ClsPiece())
        Dim p As Integer 'pieces
        Dim i As Integer 'positions index
        Dim h As Integer 'horiz
        Dim hd As Integer 'positions index
        Dim v As Integer 'verti
        Dim vd As Integer 'positions index
        Dim boardFull As Boolean
        Dim d As Integer 'directional possibilities (i.e. 8)

        For p = 1 To MAXPIECE
            i = 0
            Dim positions(MAXHCELL * MAXVCELL - 1, 1) As Integer
            For h = 1 To MAXHCELL ' Get a list of positions occupied by this piece
                For v = 1 To MAXVCELL
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
                    Debug.Print("Found other piece at h" & h & " v" & v) 'next
                ElseIf Not aBoard.IsOnBoard(h, v) Then
                    Debug.Print("Found edge of board")
                Else 'you've found somewhere to grow
                    pieces(p).SetCell(h, v)
                    pieces(p).SetHPos(1)
                    pieces(p).SetVPos(1)
                    boardFull = IsBoardFull(aBoard, pieces)
                    Exit For
                End If
                NextDirection(hd, vd) 'passed by ref so they can change
            Next d
            If boardFull Then
                Exit For
            End If
            'While aBoard.GetPosition(h, v) And f < 1000 'is occupied
            '        h = CInt(Int(HSize * Rnd() + 1))
            '        v = CInt(Int(VSize * Rnd() + 1))
            '        f += 1
            '    End While
            '    pieces(p).Colour = p
            '    pieces(p).SetCell(h, v)
            '    pieces(p).SetHPos(1)
            '    pieces(p).SetVPos(1)
            '    If IsBoardFull(aBoard, pieces) Then 'This is currently the only place where piece occupancy is set
            '        MsgBox("This shouldn't happen")
            '    End If
        Next p
    End Sub

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
        'move to next direction clockwise of eight directions
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


End Class