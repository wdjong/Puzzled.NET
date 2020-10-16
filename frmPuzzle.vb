Imports System.IO
Imports Puzzled.My

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
    Friend WithEvents MnuPuzzleWalter As System.Windows.Forms.MenuItem
    Friend WithEvents MnuHelpSolve As MenuItem
    Friend WithEvents MnuFileLoad As MenuItem
    Friend WithEvents MnuHelpInstructions As MenuItem
    Friend WithEvents MnuPuzzleL1 As MenuItem
    Friend WithEvents MnuPuzzleL2 As MenuItem
    Friend WithEvents MnuPuzzleL3 As MenuItem
    Friend WithEvents MnuPuzzleSound As MenuItem
    Friend WithEvents MnuPuzzleSolve As MenuItem
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents StatusBar1 As ToolStripStatusLabel
    Friend WithEvents ProgressBar1 As ToolStripProgressBar
    Friend WithEvents MnuHelp As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmPuzzle))
        Me.Game = New System.Windows.Forms.PictureBox()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MnuFile = New System.Windows.Forms.MenuItem()
        Me.MnuFileLoad = New System.Windows.Forms.MenuItem()
        Me.MnuFileExit = New System.Windows.Forms.MenuItem()
        Me.MnuPuzzle = New System.Windows.Forms.MenuItem()
        Me.MnuPuzzleOriginal = New System.Windows.Forms.MenuItem()
        Me.MnuPuzzleWalter = New System.Windows.Forms.MenuItem()
        Me.MnuPuzzleL1 = New System.Windows.Forms.MenuItem()
        Me.MnuPuzzleL2 = New System.Windows.Forms.MenuItem()
        Me.MnuPuzzleL3 = New System.Windows.Forms.MenuItem()
        Me.MnuPuzzleSound = New System.Windows.Forms.MenuItem()
        Me.MnuPuzzleSolve = New System.Windows.Forms.MenuItem()
        Me.MnuHelp = New System.Windows.Forms.MenuItem()
        Me.MnuHelpInstructions = New System.Windows.Forms.MenuItem()
        Me.MnuHelpSolve = New System.Windows.Forms.MenuItem()
        Me.MnuHelpAbout = New System.Windows.Forms.MenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.StatusBar1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        CType(Me.Game, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Game
        '
        Me.Game.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Game.Location = New System.Drawing.Point(0, 0)
        Me.Game.Name = "Game"
        Me.Game.Size = New System.Drawing.Size(480, 491)
        Me.Game.TabIndex = 0
        Me.Game.TabStop = False
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuFile, Me.MnuPuzzle, Me.MnuHelp})
        '
        'MnuFile
        '
        Me.MnuFile.Index = 0
        Me.MnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuFileLoad, Me.MnuFileExit})
        Me.MnuFile.Text = "&File"
        '
        'MnuFileLoad
        '
        Me.MnuFileLoad.Index = 0
        Me.MnuFileLoad.Text = "&Load"
        '
        'MnuFileExit
        '
        Me.MnuFileExit.Index = 1
        Me.MnuFileExit.Text = "E&xit"
        '
        'MnuPuzzle
        '
        Me.MnuPuzzle.Index = 1
        Me.MnuPuzzle.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuPuzzleOriginal, Me.MnuPuzzleWalter, Me.MnuPuzzleL1, Me.MnuPuzzleL2, Me.MnuPuzzleL3, Me.MnuPuzzleSound, Me.MnuPuzzleSolve})
        Me.MnuPuzzle.Text = "&Puzzle"
        '
        'MnuPuzzleOriginal
        '
        Me.MnuPuzzleOriginal.Index = 0
        Me.MnuPuzzleOriginal.Text = "&Original"
        '
        'MnuPuzzleWalter
        '
        Me.MnuPuzzleWalter.Index = 1
        Me.MnuPuzzleWalter.Text = "&Generate"
        '
        'MnuPuzzleL1
        '
        Me.MnuPuzzleL1.Index = 2
        Me.MnuPuzzleL1.Text = "Level 1"
        '
        'MnuPuzzleL2
        '
        Me.MnuPuzzleL2.Index = 3
        Me.MnuPuzzleL2.Text = "Level 2"
        '
        'MnuPuzzleL3
        '
        Me.MnuPuzzleL3.Index = 4
        Me.MnuPuzzleL3.Text = "Level 3"
        '
        'MnuPuzzleSound
        '
        Me.MnuPuzzleSound.Index = 5
        Me.MnuPuzzleSound.Text = "Sound"
        '
        'MnuPuzzleSolve
        '
        Me.MnuPuzzleSolve.Index = 6
        Me.MnuPuzzleSolve.Text = "Solve"
        '
        'MnuHelp
        '
        Me.MnuHelp.Index = 2
        Me.MnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuHelpInstructions, Me.MnuHelpSolve, Me.MnuHelpAbout})
        Me.MnuHelp.Text = "&Help"
        '
        'MnuHelpInstructions
        '
        Me.MnuHelpInstructions.Index = 0
        Me.MnuHelpInstructions.Text = "&Instructions"
        '
        'MnuHelpSolve
        '
        Me.MnuHelpSolve.Index = 1
        Me.MnuHelpSolve.Text = "&Show answer"
        '
        'MnuHelpAbout
        '
        Me.MnuHelpAbout.Index = 2
        Me.MnuHelpAbout.Text = "&About"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusBar1, Me.ProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 469)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(480, 22)
        Me.StatusStrip1.TabIndex = 3
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'StatusBar1
        '
        Me.StatusBar1.BackColor = System.Drawing.SystemColors.Menu
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Size = New System.Drawing.Size(62, 17)
        Me.StatusBar1.Text = "StatusBar1"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(100, 16)
        '
        'FrmPuzzle
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(480, 491)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Game)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "FrmPuzzle"
        Me.Text = "Puzzled?"
        CType(Me.Game, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Const MAXHCELL As Byte = 8 'Number of horizontal Cells on board
    Const MAXVCELL As Byte = 8 'Number of vertical Cells on board
    Const MAGN As Byte = 20 'Magnification... number of pixels per cell
    Const HSIZE As Byte = 8 'Horizontal number of cells in piece
    Const aBorderWidth As Byte = 2 'Add a Line around the Field of play
    ReadOnly aGame As New ClsGame
    ReadOnly aBoard As New ClsBoard
    ReadOnly pieces(8) As ClsPiece
    Shared draggingPiece As Integer
    Shared ReadOnly draggingPieceZ As Integer
    Const MAXPIECE As Byte = 8

    Private Sub FrmPuzzle_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim p As Integer
        Dim sz As System.Drawing.Size

        sz.Width = MAXHCELL * MAGN * 3 + aBorderWidth * 2
        sz.Height = MAXVCELL * MAGN * 3 + aBorderWidth * 2 + Me.ProgressBar1.Height + Me.StatusStrip1.Height
        Me.ClientSize = sz
        For p = 1 To MAXPIECE
            pieces(p) = New ClsPiece
        Next p
        aGame.StartTime = Now()
        aGame.Init(aBoard, pieces, "original.dat")
        MnuPuzzleL1.Checked = True
    End Sub

    Private Sub Game_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Game.Paint
        Dim p As Integer

        Me.BackColor = System.Drawing.Color.Black
        aBoard.DrawIt(e.Graphics)
        For p = 1 To MAXPIECE
            pieces(p).DrawIt(e.Graphics)
        Next p
    End Sub

    Private Sub Game_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Game.MouseDown
        'Initiate dragging with mouse and check for rotate (Right button) and flip (R-Button + Shift)
        Dim Button As Integer = e.Button \ &H100000
        Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        draggingPiece = aGame.WhichPiece(pieces, e.X, e.Y)
        'draggingPieceZ = pieces(draggingPiece).GetZ()
        If draggingPiece > 0 Then
            If Shift = 0 And Button = System.Windows.Forms.Keys.RButton Then 'Right click to rotate
                pieces(draggingPiece).RotateL()
                Game.Invalidate()
            ElseIf Shift = 1 And Button = System.Windows.Forms.Keys.RButton Then 'Shift Right button to flip
                pieces(draggingPiece).Flip()
                Game.Invalidate()
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
                If dH <> pieces(draggingPiece).GetHPos Then
                    If dH > (-MAXHCELL + 1 - pieces(draggingPiece).LeftExtremity) And dH < (2 * MAXHCELL + 2 - pieces(draggingPiece).RightExtremity) Then
                        pieces(draggingPiece).SetHPos(dH)
                        Game.Invalidate()
                    End If
                End If
                If dV <> pieces(draggingPiece).GetVPos Then 'has moved
                    If dV > (-MAXVCELL + 1 - pieces(draggingPiece).TopExtremity) And dV < (2 * MAXVCELL + 2 - pieces(draggingPiece).BottomExtremity) Then
                        pieces(draggingPiece).SetVPos(dV)
                        Game.Invalidate()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Game_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Game.MouseUp
        Dim Button As Integer = e.Button \ &H100000
        Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000

        If draggingPiece > 0 Then
            If pieces(draggingPiece).IsOnBoard Then
                'Perhaps they're all on board
                'If aGame.AllOnBoard(pieces) Then 'this prevents errors when pieces are off board
                aGame.SetPiecesOnBoard(aBoard, pieces)
                If aGame.IsBoardFull(aBoard) Then
                    aGame.EndGame()
                End If
                'End If
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
        aGame.PuzzleName = aPuzzle & ".dat"
        aGame.CreatePieces(aGame.PuzzleName, pieces)
        aGame.StartTime = Now
        Game.Invalidate()
    End Sub

    Private Sub MnuPuzzleWalter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuPuzzleWalter.Click
        Const aPuzzle As String = "Walter"

        Dim SoundInst As New SoundClass
        SoundInst.PlaySoundFile(aPuzzle & ".wav")
        aGame.PuzzleName = aPuzzle & ".dat"
        aGame.GeneratePieces2()
        aGame.CreatePieces(aGame.PuzzleName, pieces)
        aGame.StartTime = Now
        Game.Invalidate()
    End Sub

    Private Sub MnuHelpAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuHelpAbout.Click
        FrmAbout.Show()
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
            If My.Settings.Sound Then
                My.Computer.Audio.Play(filename)
            End If
        End Sub
    End Class

    Private Sub MnuHelpSolve_Click(sender As Object, e As EventArgs) Handles MnuHelpSolve.Click
        aGame.ShowSolution(aGame.PuzzleName, pieces)
        Game.Invalidate()
    End Sub

    Private Sub MnuFileLoad_Click(sender As Object, e As EventArgs) Handles MnuFileLoad.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim aPuzzle As String

        fd.Title = "Choose a puzzle"
        'fd.InitialDirectory = "C:\"
        fd.Filter = "Puzzle files (*.dat)|*.dat|puzzle files (*.dat)|*.dat"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            aPuzzle = Path.GetFileName(fd.FileName).Replace(".dat", "")
            Dim SoundInst As New SoundClass
            SoundInst.PlaySoundFile(aPuzzle & ".wav")
            aGame.PuzzleName = aPuzzle & ".dat"
            aGame.CreatePieces(aGame.PuzzleName, pieces)
            aGame.StartTime = Now
            Game.Invalidate()
        End If
    End Sub

    Private Sub MnuHelpInstructions_Click(sender As Object, e As EventArgs) Handles MnuHelpInstructions.Click
        Try
            Dim AppPath As String = System.AppDomain.CurrentDomain.BaseDirectory
            System.Diagnostics.Process.Start(AppPath + "PuzzledHelp.html")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub MnuPuzzleL1_Click(sender As Object, e As EventArgs) Handles MnuPuzzleL1.Click
        My.Settings.Level = 1
        MnuPuzzleL1.Checked = True
        MnuPuzzleL2.Checked = False
        MnuPuzzleL3.Checked = False
    End Sub

    Private Sub MnuPuzzleL2_Click(sender As Object, e As EventArgs) Handles MnuPuzzleL2.Click
        My.Settings.Level = 2
        MnuPuzzleL1.Checked = False
        MnuPuzzleL2.Checked = True
        MnuPuzzleL3.Checked = False
    End Sub

    Private Sub MnuPuzzleL3_Click(sender As Object, e As EventArgs) Handles MnuPuzzleL3.Click
        My.Settings.Level = 3
        MnuPuzzleL1.Checked = False
        MnuPuzzleL2.Checked = False
        MnuPuzzleL3.Checked = True
    End Sub

    Private Sub MnuPuzzleSound_Click(sender As Object, e As EventArgs) Handles MnuPuzzleSound.Click
        If My.Settings.Sound Then
            My.Settings.Sound = False
        Else
            My.Settings.Sound = True
        End If
        MnuPuzzleSound.Checked = My.Settings.Sound
    End Sub

    Private Sub MnuPuzzleSolve_Click(sender As Object, e As EventArgs) Handles MnuPuzzleSolve.Click
        aGame.SolvePuzzle(aBoard, pieces)
    End Sub
End Class
