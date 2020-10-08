Public Class FrmPuzzle
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Game As System.Windows.Forms.PictureBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents MnuFileExit As System.Windows.Forms.MenuItem
    Friend WithEvents MnuPuzzle As System.Windows.Forms.MenuItem
    Friend WithEvents MnuPuzzleOriginal As System.Windows.Forms.MenuItem
    Friend WithEvents MnuHelpAbout As System.Windows.Forms.MenuItem
    Friend WithEvents MnuPuzzleFreya As System.Windows.Forms.MenuItem
    Friend WithEvents MnuPuzzleAlex As System.Windows.Forms.MenuItem
    Friend WithEvents MnuPuzzleSarah As System.Windows.Forms.MenuItem
    Friend WithEvents MnuPuzzleNicholas As System.Windows.Forms.MenuItem
    Friend WithEvents MnuPuzzleWalter As System.Windows.Forms.MenuItem
    Friend WithEvents MnuHelp As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmPuzzle))
        Me.Game = New System.Windows.Forms.PictureBox
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MnuFile = New System.Windows.Forms.MenuItem
        Me.MnuFileExit = New System.Windows.Forms.MenuItem
        Me.MnuPuzzle = New System.Windows.Forms.MenuItem
        Me.MnuPuzzleOriginal = New System.Windows.Forms.MenuItem
        Me.MnuPuzzleFreya = New System.Windows.Forms.MenuItem
        Me.MnuPuzzleAlex = New System.Windows.Forms.MenuItem
        Me.MnuPuzzleSarah = New System.Windows.Forms.MenuItem
        Me.MnuPuzzleNicholas = New System.Windows.Forms.MenuItem
        Me.MnuPuzzleWalter = New System.Windows.Forms.MenuItem
        Me.MnuHelp = New System.Windows.Forms.MenuItem
        Me.MnuHelpAbout = New System.Windows.Forms.MenuItem
        CType(Me.Game, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Game
        '
        Me.Game.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Game.Location = New System.Drawing.Point(0, 0)
        Me.Game.Name = "Game"
        Me.Game.Size = New System.Drawing.Size(480, 446)
        Me.Game.TabIndex = 0
        Me.Game.TabStop = False
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuFile, Me.MnuPuzzle, Me.MnuHelp})
        '
        'mnuFile
        '
        Me.MnuFile.Index = 0
        Me.MnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuFileExit})
        Me.MnuFile.Text = "&File"
        '
        'mnuFileExit
        '
        Me.MnuFileExit.Index = 0
        Me.MnuFileExit.Text = "E&xit"
        '
        'mnuPuzzle
        '
        Me.MnuPuzzle.Index = 1
        Me.MnuPuzzle.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuPuzzleOriginal, Me.MnuPuzzleFreya, Me.MnuPuzzleAlex, Me.MnuPuzzleSarah, Me.MnuPuzzleNicholas, Me.MnuPuzzleWalter})
        Me.MnuPuzzle.Text = "&Puzzle"
        '
        'mnuPuzzleOriginal
        '
        Me.MnuPuzzleOriginal.Index = 0
        Me.MnuPuzzleOriginal.Text = "&Original"
        '
        'mnuPuzzleFreya
        '
        Me.MnuPuzzleFreya.Index = 1
        Me.MnuPuzzleFreya.Text = "&Freya"
        '
        'mnuPuzzleAlex
        '
        Me.MnuPuzzleAlex.Index = 2
        Me.MnuPuzzleAlex.Text = "&Alex"
        '
        'mnuPuzzleSarah
        '
        Me.MnuPuzzleSarah.Index = 3
        Me.MnuPuzzleSarah.Text = "&Sarah"
        '
        'mnuPuzzleNicholas
        '
        Me.MnuPuzzleNicholas.Index = 4
        Me.MnuPuzzleNicholas.Text = "&Nicholas"
        '
        'mnuPuzzleWalter
        '
        Me.MnuPuzzleWalter.Index = 5
        Me.MnuPuzzleWalter.Text = "&Walter"
        '
        'mnuHelp
        '
        Me.MnuHelp.Index = 2
        Me.MnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuHelpAbout})
        Me.MnuHelp.Text = "&Help"
        '
        'mnuHelpAbout
        '
        Me.MnuHelpAbout.Index = 0
        Me.MnuHelpAbout.Text = "&About"
        '
        'frmPuzzle
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(480, 446)
        Me.Controls.Add(Me.Game)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmPuzzle"
        Me.Text = "Puzzled?"
        CType(Me.Game, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Const MAXHCELL As Byte = 8 'Number of horizontal Cells on board
    Const MAXVCELL As Byte = 8 'Number of vertical Cells on board
    Const MAGN As Byte = 20 'Magnification... number of pixels per cell
    Const HSIZE As Byte = 8 'Horizontal number of cells in piece
    Const aBorderWidth As Byte = 2 'Add a Line around the Field of play
    ReadOnly aPuzzled As New clsPuzzled
    ReadOnly aBoard As New clsBoard
    ReadOnly aPiece(8) As clsPiece
    Shared draggingPiece As Integer
    Shared startTime As Date
    Const numPieces As Byte = 8

    Private Sub FrmPuzzle_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim p As Integer
        Dim sz As System.Drawing.Size

        sz.Width = MAXHCELL * MAGN * 3 + aBorderWidth * 2
        sz.Height = MAXVCELL * MAGN * 3 + aBorderWidth * 2
        Me.ClientSize = sz
        For p = 1 To numPieces
            aPiece(p) = New clsPiece
        Next p
        startTime = Now()
        aPuzzled.init(aBoard, aPiece, "original.dat")
    End Sub

    Private Sub Game_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Game.Paint
        Dim p As Integer

        Me.BackColor = System.Drawing.Color.Black
        aBoard.drawIt(e.Graphics)
        For p = 1 To numPieces
            aPiece(p).drawIt(e.Graphics)
        Next p
    End Sub

    Private Sub Game_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Game.MouseDown
        'Initiate dragging with mouse and check for rotate (Right button) and flip (R-Button + Shift)
        Dim Button As Integer = e.Button \ &H100000
        Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        draggingPiece = aPuzzled.whichPiece(aPiece, e.X, e.Y)
        If draggingPiece > 0 Then
            If Shift = 0 And Button = System.Windows.Forms.Keys.RButton Then 'Right click to rotate
                aPiece(draggingPiece).rotateL()
                Game.Invalidate()
            Else
                If Shift = 1 And Button = System.Windows.Forms.Keys.RButton Then 'Shift Right button to flip
                    aPiece(draggingPiece).flip()
                    Game.Invalidate()
                End If
            End If
        End If
    End Sub

    Private Sub Game_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Game.MouseMove
        Dim Button As Integer = e.Button \ &H100000
        Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim dH As Integer
        Dim dV As Integer

        If Button = System.Windows.Forms.Keys.LButton Then 'dragging
            If draggingPiece > 0 Then
                dH = Int(e.X / MAGN) - MAXHCELL - HSIZE / 2 + 1
                dV = Int(e.Y / MAGN) - MAXVCELL - HSIZE / 2 + 1
                If dH <> aPiece(draggingPiece).getHPos Then
                    If dH > (-MAXHCELL + 1 - aPiece(draggingPiece).leftExtremity) And dH < (2 * MAXHCELL + 2 - aPiece(draggingPiece).rightExtremity) Then
                        aPiece(draggingPiece).setHPos(dH)
                        Game.Invalidate()
                    End If
                End If
                If dV <> aPiece(draggingPiece).getVPos Then 'has moved
                    If dV > (-MAXVCELL + 1 - aPiece(draggingPiece).topExtremity) And dV < (2 * MAXVCELL + 2 - aPiece(draggingPiece).bottomExtremity) Then
                        aPiece(draggingPiece).setVPos(dV)
                        Game.Invalidate()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Game_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Game.MouseUp
        Dim Button As Integer = e.Button \ &H100000
        Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim timetaken As Single

        If draggingPiece > 0 Then
            If aPiece(draggingPiece).isOnBoard Then
                'Perhaps they're all on board
                If aPuzzled.allOnBoard(aPiece) Then
                    If aPuzzled.isBoardFull(aBoard, aPiece) Then
                        timetaken = DateDiff(Microsoft.VisualBasic.DateInterval.Second, startTime, Now)
                        MsgBox("Well done. That only took " & timetaken & " seconds.")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub MnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuFileExit.Click
        End
    End Sub

    Private Sub FrmPuzzle_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        End
    End Sub

    Private Sub MnuPuzzleOriginal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuPuzzleOriginal.Click
        Const aPuzzle As String = "Original"
        Dim SoundInst As New SoundClass
        SoundInst.PlaySoundFile(aPuzzle & ".wav")
        aPuzzled.createPieces(aPuzzle & ".dat", aPiece)
        startTime = Now
        Game.Invalidate()
    End Sub

    Private Sub MnuPuzzleFreya_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuPuzzleFreya.Click
        Const aPuzzle As String = "Freya"
        Dim SoundInst As New SoundClass
        SoundInst.PlaySoundFile(aPuzzle & ".wav")
        aPuzzled.createPieces(aPuzzle & ".dat", aPiece)
        startTime = Now
        Game.Invalidate()
    End Sub

    Private Sub MnuPuzzleAlex_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuPuzzleAlex.Click
        Const aPuzzle As String = "Alex"
        Dim SoundInst As New SoundClass
        SoundInst.PlaySoundFile(aPuzzle & ".wav")
        aPuzzled.createPieces(aPuzzle & ".dat", aPiece)
        startTime = Now
        Game.Invalidate()
    End Sub

    Private Sub MnuPuzzleSarah_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuPuzzleSarah.Click
        Const aPuzzle As String = "Sarah"
        Dim SoundInst As New SoundClass
        SoundInst.PlaySoundFile(aPuzzle & ".wav")
        aPuzzled.createPieces(aPuzzle & ".dat", aPiece)
        startTime = Now
        Game.Invalidate()
    End Sub

    Private Sub MnuPuzzleNicholas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuPuzzleNicholas.Click
        Const aPuzzle As String = "Nichol"
        Dim SoundInst As New SoundClass
        SoundInst.PlaySoundFile(aPuzzle & ".wav")
        aPuzzled.createPieces(aPuzzle & ".dat", aPiece)
        startTime = Now
        Game.Invalidate()
    End Sub

    Private Sub MnuPuzzleWalter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuPuzzleWalter.Click
        Const aPuzzle As String = "Walter"
        Dim SoundInst As New SoundClass
        SoundInst.PlaySoundFile(aPuzzle & ".wav")
        aPuzzled.generatePieces()
        aPuzzled.createPieces(aPuzzle & ".dat", aPiece)
        startTime = Now
        Game.Invalidate()
    End Sub

    Private Sub MnuHelpAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuHelpAbout.Click
        frmAbout.Show()
    End Sub

    Public Class SoundClass
        'Declare Auto Function PlaySound Lib "winmm.dll" (ByVal name _
        '   As String, ByVal hmod As Integer, ByVal flags As Integer) As Integer
        '' name specifies the sound file when the SND_FILENAME flag is set.
        '' hmod specifies an executable file handle.
        '' hmod must be Nothing if the SND_RESOURCE flag is not set.
        '' flags specifies which flags are set. 

        '' The PlaySound documentation lists all valid flags.
        'Public Const SND_SYNC = &H0          ' play synchronously
        'Public Const SND_ASYNC = &H1         ' play asynchronously
        'Public Const SND_FILENAME = &H20000  ' name is file name
        'Public Const SND_RESOURCE = &H40004  ' name is resource name or atom

        Public Sub PlaySoundFile(ByVal filename As String)
            ' Plays a sound from filename.
            'PlaySound(filename, Nothing, SND_FILENAME Or SND_ASYNC)
            My.Computer.Audio.Play(filename)
        End Sub
    End Class

End Class
