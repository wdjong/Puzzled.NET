Option Strict Off
Option Explicit On
'Each piece can be thought of as an array of cells (squares) 
'with some visible and some not.
'In the case of a single visible cell that could mean the upper left cell of
'a piece might be in position -7, -7 through to 0,0
Friend Class clsPiece
    Private Const MAXHCELL As Byte = 8
    Private Const MAXVCELL As Byte = 8
    Private Const MAGN As Byte = 20
    Private Const HSIZE As Byte = 8 'Horizontal number of cells in piece
    Private Const VSIZE As Byte = 8 'Vertical number of cells in piece
    Private Const BOARDH As Byte = MAXHCELL * MAGN 'origin of place to put pieces 'allow for a border around board
    Private Const BOARDV As Byte = MAXVCELL * MAGN

    Private pID As Integer 'Piece number
    Private cell(HSIZE + 1, VSIZE + 1) As Boolean 'true is visible: Allow extra to simplify code when checking adjacent cells
    Private hPos As Integer 'Board pos 1,1 is top left
    Private vPos As Integer 'Hori first, Vert second
    Private rotation As Integer '0 is up, 90 is right, 180 is down, 270 is left
    Private hHome As Integer 'where it belongs
    Private vHome As Integer 'where it belongs
    Private aColor(8) As System.Drawing.Color 'Array of colours to use for pieces
    Private oColour As Integer 'RGB 0 is black use

    Public Function bottomExtremity() As Integer
        'called from centreV: Find bottomost cell
        Dim h As Integer
        Dim v As Integer

        For v = VSIZE To 1 Step -1
            For h = 1 To HSIZE
                If cell(h, v) Then
                    bottomExtremity = v
                    Exit Function
                End If
            Next h
        Next v
        bottomExtremity = 0 'None found
    End Function

    Public Sub Centre()
        'Centre the piece
        centreH() 'Centre horizontally
        centreV() 'Centre vertically
    End Sub

    Private Sub CentreV()
        'Centre the piece Vertically
        Dim vL As Integer 'top extremity
        Dim vR As Integer 'bottom extremity
        Dim vD As Integer 'distance to move
        Dim vC As Integer 'vertical centre

        vL = topExtremity()
        vR = bottomExtremity()
        vC = (vL + vR) / 2 'current centre = av of 2 points
        vD = VSIZE / 2 - vC 'find distance from centre
        If vD > 0 Then 'shift down
            shiftDown(vD)
        Else
            If vD < 0 Then 'shift up
                shiftUp(vD)
            End If
        End If
    End Sub

    Private Sub centreH()
        'Centre the piece Horizontally
        Dim hL As Integer 'left extremity
        Dim hR As Integer 'right extremity
        Dim hD As Integer 'distance to move
        Dim hC As Integer 'centre

        hL = leftExtremity()
        hR = rightExtremity()
        hC = (hL + hR) / 2 'current centre = av of 2 points
        hD = HSIZE / 2 - hC
        If hD > 0 Then 'shift right
            shiftRight(hD)
        Else
            If hD < 0 Then 'shift left
                shiftLeft(hD)
            End If
        End If
    End Sub

    Public Sub drawIt(ByVal formGraphics As System.Drawing.Graphics)
        'BoardH is an offset from window 0 to allow a border around the pieces
        'BoardV is an offset from window 0
        'MAGN determines the size of the cell currently 15
        'hPos is the position of the piece horizontally on the board: 1 is left -> 8 right
        'vPos is the position of the piece vertically on the board: 1 is top -> 8 down
        'hSize is the number of cells horizontally within the piece currently 8 same as board size
        'vSize is the number of cells vertically within the piece: currently 8 the same as board

        Dim pieceH As Integer 'Pixel position left of whole piece
        Dim pieceV As Integer 'Pixel position top of whole piece
        Dim cellH As Integer 'Pixel position left of current cell
        Dim cellV As Integer 'Pixel position top of current cell
        Dim h As Integer 'current cell hori co-ord within piece
        Dim v As Integer 'current cell vert co-ord within piece
        Dim myBrush As New System.Drawing.SolidBrush(System.Drawing.Color.Gray)
        Dim pen As New Drawing.Pen(System.Drawing.Color.Red, 1)

        Try
            pieceH = BOARDH + (hPos - 1) * MAGN 'Convert upper left position to pixel pos
            pieceV = BOARDV + (vPos - 1) * MAGN '
            For h = 1 To HSIZE 'each column of cells of the piece
                cellH = pieceH + (h - 1) * MAGN 'find the left origin in pixels of the cell
                For v = 1 To VSIZE 'each line of cells of the piece
                    cellV = pieceV + (v - 1) * MAGN 'find the top origin in pixesl of the cell
                    If cell(h, v) Then 'the cell is visible
                        myBrush.Color = aColor(oColour)
                        formGraphics.FillRectangle(myBrush, cellH, cellV, MAGN, MAGN) ', colour, BF
                        If Not cell(h - 1, v) Then 'nothing to left so light line on left
                            pen.Color = System.Drawing.Color.White
                            formGraphics.DrawLine(pen, cellH, cellV, cellH, cellV + MAGN - 1)
                        End If
                        If Not cell(h, v - 1) Then 'nothing above so light line on top
                            pen.Color = System.Drawing.Color.White
                            formGraphics.DrawLine(pen, cellH, cellV, cellH + MAGN - 1, cellV)
                        End If
                        If Not cell(h + 1, v) Then 'nothing to right so shadow on right
                            pen.Color = System.Drawing.Color.Black
                            formGraphics.DrawLine(pen, cellH + MAGN - 1, cellV, cellH + MAGN - 1, cellV + MAGN - 1)
                        End If
                        If Not cell(h, v + 1) Then 'nothing below
                            pen.Color = System.Drawing.Color.Black
                            formGraphics.DrawLine(pen, cellH, cellV + MAGN - 1, cellH + MAGN - 1, cellV + MAGN - 1)
                        End If
                    End If
                Next v
            Next h
            myBrush.Dispose()
            pen.Dispose()
        Catch ex As Exception
            MsgBox("clsPiece.drawIt: Drawing a piece: " & ex.Message)
        End Try
    End Sub

    Public Sub flip()
        'Turn piece upside down
        Dim cell2(HSIZE, VSIZE) As Boolean 'create a copy of the piece
        Dim h As Integer 'cell reference h
        Dim v As Integer 'cell reference v

        For h = 1 To HSIZE 'Copy piece
            For v = 1 To VSIZE
                cell2(h, v) = cell(h, v)
            Next v
        Next h
        For h = 1 To HSIZE 'Flip vertically
            For v = 1 To VSIZE
                cell(h, VSIZE + 1 - v) = cell2(h, v)
            Next v
        Next h

    End Sub

    Public Function getCell(ByRef h As Integer, ByRef v As Integer) As Boolean
        'If the cell is a visible section then this is true
        getCell = cell(h, v)
    End Function

    Public Property Colour() As Integer
        Get
            Colour = oColour
        End Get
        Set(ByVal Value As Integer)
            oColour = Value
        End Set
    End Property

    Public Function getHPos() As Integer
        'The cells horizontal position 
        getHPos = hPos
    End Function

    Public Function getPID() As Integer
        'The cells id
        getPID = pID
    End Function

    Public Function getRotation() As Integer

        getRotation = rotation
    End Function

    Public Function getVPos() As Integer
        getVPos = vPos
    End Function

    Public Sub init(ByVal aPID As Integer, ByVal aColour As Integer)
        'Initialize Piece object
        Dim h As Integer ' Hori cell
        Dim v As Integer ' Vert cell

        pID = aPID
        oColour = aColour
        For h = 1 To HSIZE 'left to right
            For v = 1 To VSIZE 'top to bottom
                cell(h, v) = True 'visible
            Next v
        Next h
        rotation = 0 'up
        hPos = 1 'left
        vPos = 1 'top
        aColor(0) = System.Drawing.Color.Transparent
        aColor(1) = System.Drawing.Color.Red
        aColor(2) = System.Drawing.Color.Orange
        aColor(3) = System.Drawing.Color.Yellow
        aColor(4) = System.Drawing.Color.Green
        aColor(5) = System.Drawing.Color.Blue
        aColor(6) = System.Drawing.Color.Indigo
        aColor(7) = System.Drawing.Color.Violet
        aColor(8) = System.Drawing.Color.Beige
    End Sub

    Public Function isAPiece() As Boolean
        'some piece objects may have no visible parts
        'therefore they for instance don't need to be on the board
        isAPiece = (leftExtremity() <> 0)
    End Function

    Public Function isClicked(ByRef h As Integer, ByRef v As Integer) As Boolean
        'Given the position of the piece relative to the board find where on the piece you clicked
        Dim xOnBoard As Integer 'position of click on board
        Dim yOnBoard As Integer
        Dim xOnPiece As Integer 'Position of click on piece
        Dim yOnPiece As Integer

        xOnBoard = Int(h / MAGN) - (MAXHCELL - 1) '-7 -> 0 off board then board = 1 to 8
        yOnBoard = Int(v / MAGN) - (MAXVCELL - 1)
        xOnPiece = (xOnBoard - hPos) + 1
        yOnPiece = (yOnBoard - vPos) + 1
        If xOnPiece < 1 Or xOnPiece > HSIZE Or yOnPiece < 1 Or yOnPiece > VSIZE Then
            isClicked = False 'not clicked on piece
        Else 'clicked on piece but may or may not be visible bit
            isClicked = cell(xOnPiece, yOnPiece)
        End If
        'e.g. If -3,-3 of board clicked
        'and if upper left of piece (hpos, vpos) is at -4,-4
        'then the clicked cell of the piece is 2,2
        'or, in the simplest case, if 2,2 is clicked on the board
        'and the piece is at 1,1
        'then cell 2,2 of piece was clicked
    End Function

    Public Function isOnBoard() As Boolean
        'Determine if the piece is on the board
        Dim hL As Integer
        Dim vT As Integer
        Dim hR As Integer
        Dim vB As Integer
        Dim aBoard As New clsBoard
        isOnBoard = False
        hL = leftExtremity()
        vT = topExtremity()
        hR = rightExtremity()
        vB = bottomExtremity()
        If aBoard.isOnBoard(hPos - 1 + hL, vPos - 1 + vT) Then
            If aBoard.isOnBoard(hPos - 1 + hR, vPos - 1 + vB) Then
                isOnBoard = True
            End If
        End If
    End Function

    Public Function leftExtremity() As Integer
        'Finding the left most visible cell is used by other routines for centring and so on
        'also for check if there's any visible cell at all i.e. if the piece 'exists'
        Dim h As Integer
        Dim v As Integer

        For h = 1 To HSIZE
            For v = 1 To VSIZE
                If cell(h, v) Then 'found a cell
                    leftExtremity = h
                    Exit Function
                End If
            Next v
        Next h
        leftExtremity = 0 'Empty
    End Function

    Public Function rightExtremity() As Integer
        Dim h As Integer
        Dim v As Integer

        For h = HSIZE To 1 Step -1
            For v = 1 To VSIZE
                If cell(h, v) Then 'found a cell
                    rightExtremity = h
                    Exit Function
                End If
            Next v
        Next h
        rightExtremity = 0 'Empty
    End Function

    Public Sub resetCell(ByRef h As Integer, ByRef v As Integer)
        'make this cell invisible
        cell(h, v) = False
    End Sub

    Public Sub rotateL()
        'Anticlockwise 90 degree rotation
        Dim cell2(HSIZE, VSIZE) As Boolean
        Dim h As Integer
        Dim v As Integer

        For h = 1 To HSIZE
            For v = 1 To VSIZE
                cell2(h, v) = cell(h, v)
            Next v
        Next h
        For h = 1 To HSIZE
            For v = 1 To VSIZE
                cell(v, HSIZE + 1 - h) = cell2(h, v)
            Next v
        Next h
    End Sub

    Public Sub setCell(ByRef h As Integer, ByRef v As Integer)
        cell(h, v) = True
    End Sub

    Public Sub setHPos(ByRef h As Integer)
        hPos = h
    End Sub

    Public Sub setPID(ByRef p As Integer)
        pID = p
    End Sub

    Public Sub setRotation(ByRef a As Integer)
        rotation = a
    End Sub

    Public Sub setVPos(ByRef v As Integer)
        vPos = v
    End Sub

    Private Sub shiftDown(ByRef vD As Integer)
        Dim h As Integer
        Dim v As Integer

        For v = VSIZE To 1 Step -1 'bottom to top
            For h = 1 To HSIZE 'left to right
                If v > vD Then 'apart from the last one or two
                    cell(h, v) = cell(h, v - vD) 'shift it down to current from above
                Else
                    cell(h, v) = False 'v - vd would be illegal (0 or -ve) if v <= vd
                End If
            Next h
        Next v
    End Sub

    Private Sub shiftLeft(ByRef hD As Integer)
        Dim h As Integer
        Dim v As Integer

        For h = 1 To HSIZE 'each destination
            For v = 1 To VSIZE 'each row
                If h <= HSIZE + hD Then
                    cell(h, v) = cell(h - hD, v)
                Else 'e.g. column 8 when hd = -1
                    cell(h, v) = False
                End If
            Next v
        Next h
    End Sub

    Private Sub shiftRight(ByRef hD As Integer)
        Dim h As Integer 'hori cell
        Dim v As Integer 'vert cell

        For h = HSIZE To 1 Step -1 'go right to left
            For v = 1 To VSIZE 'for each row
                If h > hD Then 'is a destination column
                    cell(h, v) = cell(h - hD, v)
                Else 'not a destination so clear
                    cell(h, v) = False
                End If
            Next v
        Next h
    End Sub

    Private Sub shiftUp(ByRef vD As Integer)
        Dim h As Integer
        Dim v As Integer

        For v = 1 To VSIZE 'each row
            For h = 1 To HSIZE 'each column
                If v <= VSIZE + vD Then '
                    cell(h, v) = cell(h, v - vD)
                Else
                    cell(h, v) = False
                End If
            Next h
        Next v
    End Sub

    Public Function topExtremity() As Integer
        Dim h As Integer
        Dim v As Integer

        For v = 1 To VSIZE 'find vertical extremities
            For h = 1 To HSIZE
                If cell(h, v) Then
                    topExtremity = v
                    Exit Function
                End If
            Next h
        Next v
        topExtremity = 0 'Empty
    End Function
End Class