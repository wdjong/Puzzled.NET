Option Strict Off
Option Explicit On
Friend Class clsPuzzled
    Const MAXPIECE As Byte = 8 'Pieces in the game
    Const MAXHCELL As Byte = 8 'Horizontal cells in a piece
    Const MAXVCELL As Byte = 8 'Vertical cells in a piece
    Const HSize As Byte = 8 'Horizontal cells in the board
    Const VSize As Byte = 8 'Vertical cells in the board

    Function allOnBoard(ByVal aPiece() As clsPiece) As Boolean
        'Check that all pieces are fully on the board
        Dim p As Integer

        allOnBoard = True 'lets assume for the moment
        For p = 1 To MAXPIECE
            If aPiece(p).isAPiece Then 'it needs to be on board
                If Not aPiece(p).isOnBoard Then
                    allOnBoard = False 'one bad apple spoils the whole lot
                    Exit Function
                End If
            End If
        Next p
    End Function

    Public Sub createPieces(ByVal puzFile As String, ByVal aPiece() As clsPiece)
        'Read piece dimensions from .dat file PUZFILE
        Dim answer(MAXHCELL, MAXVCELL) As Integer
        Dim h As Integer 'h hori cell in piece
        Dim v As Integer 'v vertical cell in piece
        Dim p As Integer 'piece
        Dim x As Integer 'position

        'This is an array of number representing the location of different colours in the puzzle solution
        Try
            FileOpen(1, puzFile, OpenMode.Input)
            For h = 1 To MAXHCELL
                For v = 1 To MAXVCELL
                    Input(1, answer(h, v))
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
                            aPiece(p).setCell(h, v) '
                        Else
                            aPiece(p).resetCell(h, v)
                        End If
                    Next v
                Next h
                aPiece(p).Colour = p
                'This is about positioning the pieces around the board
                'There a 9 spots to put pieces but nothing will go in the 5th (so skip it)
                If p < 5 Then x = p Else x = p + 1
                aPiece(p).setHPos(((x - 1) Mod 3) * HSize - MAXHCELL + 1) '0*8-8+1, 1*8-8+1, 2*8-8+1
                aPiece(p).setVPos(Int((x - 1) / 3) * VSize - MAXVCELL + 1) '0*8-8+1,x=4:1*8-8+1,x=7:2*8-8+1
                aPiece(p).Centre()
            Next p
        Catch ex As Exception
            MsgBox("CreatePieces: " & ex.message)
        End Try
    End Sub

    Public Sub drawIt(ByVal aBoard As clsBoard, ByVal aPiece() As clsPiece, ByVal g As System.Drawing.Graphics)
        Dim p As Integer 'piece

        aBoard.drawIt(g)
        Try
            For p = 1 To MAXPIECE
                aPiece(p).drawIt(g)
            Next p
        Catch ex As Exception
            MsgBox("drawIt: " & ex.Message & " " & ex.StackTrace)
        End Try
    End Sub

    Public Sub init(ByVal aBoard As clsBoard, ByVal aPiece() As clsPiece, ByVal puzFile As String)
        'Initialise the board and it's pieces
        Dim p As Integer

        aBoard.init()
        For p = 1 To MAXPIECE
            aPiece(p).init(p, p)
        Next
        createPieces(puzFile, aPiece)
    End Sub

    Public Function isBoardFull(ByVal aBoard As clsBoard, ByVal aPiece() As clsPiece) As Boolean
        'Clear board
        'Set occupied cells for each piece
        Dim p As Integer 'piece
        Dim h As Integer
        Dim v As Integer

        aBoard.clear() 'Clear board
        For p = 1 To MAXPIECE 'Determine occupied positions
            For h = aPiece(p).leftExtremity To aPiece(p).rightExtremity
                For v = aPiece(p).topExtremity To aPiece(p).bottomExtremity
                    If aPiece(p).getCell(h, v) Then
                        aBoard.setPosition(aPiece(p).getHPos - 1 + h, aPiece(p).getVPos - 1 + v)
                    End If
                Next v
            Next h
        Next p
        For h = 1 To MAXHCELL 'Scan board for unnoccupied positions
            For v = 1 To MAXVCELL
                If Not aBoard.getPosition(h, v) Then 'unoccupied board space
                    isBoardFull = False
                    Exit Function
                End If
            Next v
        Next h
        isBoardFull = True
    End Function

    Public Function whichPiece(ByVal aPiece() As clsPiece, ByRef h As Integer, ByRef v As Integer) As Integer
        'Given a co-ordinate pointed to with the mouse
        'Return the piece pointed to
        Dim p As Integer 'piece

        'Because pieces are drawn from 1 to maxpiece
        'The bigger numbered piece is on top and therefore
        'we should assume the top one is what's intended first
        For p = MAXPIECE To 1 Step -1 'check each piece
            If aPiece(p).isClicked(h, v) Then
                whichPiece = p 'first piece is on top
                Exit Function
            End If
        Next p
        whichPiece = 0 'no piece clicked
    End Function

    Public Sub generatePieces()
        Dim answer(MAXHCELL, MAXVCELL) As Integer
        Dim h As Integer 'h hori cell in piece
        Dim v As Integer 'v vertical cell in piece
        Dim p As Integer 'piece
        Dim count As Integer = 0 'count how many pieces are filled
        Dim outputString As String = ""
        'For each piece (colour)

        Randomize()
        'Do While count < HSize * VSize 'e.g. 64
        '    p = CInt(Rnd() * MAXPIECE)
        '    h = CInt(Rnd() * MAXHCELL)
        '    v = CInt(Rnd() * MAXVCELL)
        '    If answer(h, v) = 0 Then 'No piece is visible at this position
        '        answer(h, v) = p 'set the cell to this piece colour
        '        count = count + 1
        '    End If
        'Loop

        Try
            'FileOpen(1, "walter.dat", OpenMode.Output)
            For v = 1 To MAXVCELL
                For h = 1 To MAXHCELL
                    p = CInt(Int(MAXPIECE * Rnd() + 1))
                    outputString = outputString & p.ToString 'answer(h, v)
                    If h <> MAXHCELL Then
                        outputString = outputString & ","
                    Else
                        If v <> MAXVCELL Then
                            outputString = outputString & vbNewLine
                        End If
                    End If
                Next h
            Next v
            'FileClose(1)
            My.Computer.FileSystem.WriteAllText("walter.dat", outputString, False, System.Text.ASCIIEncoding.ASCII)
        Catch ex As Exception
            MsgBox("createPieces: FileOpen: " & ex.Message)
        End Try
    End Sub
End Class