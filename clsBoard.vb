Option Strict Off
Option Explicit On

Friend Class ClsBoard
    'Each cell (position) on the board is occupied or not.
    Private Const MAXHCELL As Byte = 8 'Number of horizontal Cells
    Private Const MAXVCELL As Byte = 8 'Number of vertical Cells
    Private Const MAGN As Byte = 20 'Magnification... number of pixels per cell
    Private Const BOARDH As Byte = MAXHCELL * MAGN 'origin of place to put pieces
    Private Const BOARDV As Byte = MAXVCELL * MAGN
    Const aBorderWidth As Byte = 1
    Private ReadOnly position(MAXHCELL, MAXVCELL) As Boolean 'each cell when true is occupied

    Public Sub Clear()
        'Remove all pieces from the board
        Dim h As Integer
        Dim v As Integer

        For h = 1 To MAXHCELL
            For v = 1 To MAXVCELL
                position(h, v) = False
            Next v
        Next h
    End Sub

    Public Sub DrawIt(ByVal formGraphics As System.Drawing.Graphics)
        'Draw the board
        Dim myBrush As New System.Drawing.SolidBrush(System.Drawing.Color.Gray)
        Dim myPen As New System.Drawing.Pen(System.Drawing.Color.Blue, aBorderWidth)
        Const BOARDH2 As Short = MAXHCELL * 2 * MAGN
        Const BOARDV2 As Short = MAXVCELL * 2 * MAGN

        'The board is 3 times the size of the puzzle to allow the pieces to be arranged around the solution area
        formGraphics.FillRectangle(myBrush, 0, 0, 3 * MAXHCELL * MAGN + aBorderWidth * 2, 3 * MAXVCELL * MAGN + aBorderWidth * 2) ', RGB(&H40s, &H40s, &H40s), BF
        formGraphics.DrawRectangle(myPen, 0, 0, 3 * MAXHCELL * MAGN + aBorderWidth * 2, 3 * MAXVCELL * MAGN + aBorderWidth * 2) ', QBColor(11), B
        myPen.Color = System.Drawing.Color.Green
        myPen.DashStyle = Drawing2D.DashStyle.Dash
        formGraphics.DrawRectangle(myPen, BOARDH, BOARDV, BOARDH2 - BOARDH, BOARDV2 - BOARDV) ', QBColor(10), B 'dotted
        myPen.Dispose()
        myBrush.Dispose()
    End Sub

    Public Function IsOnBoard(ByRef h As Integer, ByRef v As Integer) As Boolean
        'Given a cell reference identify if it's on the board
        IsOnBoard = False
        If h > 0 And h < MAXHCELL + 1 Then
            If v > 0 And v < MAXVCELL + 1 Then
                IsOnBoard = True
            End If
        End If
    End Function

    Public Function PixToCoordH(ByRef h As Integer) As Integer
        'return the board co-ordinate given a h pixel
        PixToCoordH = (h / MAGN) + 1 - MAXHCELL
    End Function

    Public Function PixToCoordV(ByRef v As Integer) As Integer
        'return the board co-ordinate given a v pixel
        PixToCoordV = (v / MAGN) + 1 - MAXVCELL
    End Function

    Public Function GetPosition(ByRef h As Integer, ByRef v As Integer) As Boolean
        'is a position 'occupied'. If position(h,v) answer is yes
        GetPosition = False ' the default is to return false i.e. 
        If h > 0 And h <= MAXHCELL And v > 0 And v <= MAXVCELL Then
            GetPosition = position(h, v)
        End If
    End Function

    Public Sub Init()
        Clear()
    End Sub

    Public Sub ResetPosition(ByRef h As Integer, ByRef v As Integer)
        'clear cell position -> unoccupied
        position(h, v) = False
    End Sub

    Public Sub SetPosition(ByRef h As Integer, ByRef v As Integer)
        'indicate position on board is occupied
        position(h, v) = True
    End Sub
End Class