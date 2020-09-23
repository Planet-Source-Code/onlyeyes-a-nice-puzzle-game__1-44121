VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Puzzle Game"
   ClientHeight    =   6480
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   5055
      Left            =   5760
      ScaleHeight     =   4995
      ScaleWidth      =   4995
      TabIndex        =   4
      Top             =   480
      Width           =   5055
      Begin VB.Shape Image3 
         BorderColor     =   &H000000FF&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Height          =   975
         Left            =   240
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSForms.Image Image2 
         Height          =   5000
         Left            =   0
         Top             =   0
         Width           =   5000
         AutoSize        =   -1  'True
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "8819;8819"
         PictureTiling   =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5060
      Left            =   480
      ScaleHeight     =   4995
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   480
      Width           =   5060
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      Top             =   5880
      Width           =   1335
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   600
      Top             =   5760
      _ExtentX        =   7938
      _ExtentY        =   7938
      _Version        =   393216
      Picture         =   "Puzzle1.frx":0000
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   1
      Left            =   5640
      Picture         =   "Puzzle1.frx":41F02
      Top             =   6000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   0
      Left            =   5400
      Picture         =   "Puzzle1.frx":4228C
      Top             =   6000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5160
      Picture         =   "Puzzle1.frx":4259E
      ToolTipText     =   "More Details"
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rezolve Puzzle"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Picture Preview"
      Height          =   195
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   1110
   End
   Begin VB.Menu mnuFile 
      Caption         =   "F&ile"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnushuffle 
         Caption         =   "Shuffle"
         Shortcut        =   {F2}
      End
      Begin VB.Menu line 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "E&dit"
      Begin VB.Menu mnubeginer 
         Caption         =   "Beginner"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAdvansed 
         Caption         =   "Advanced"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "H&elp"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim uniqueValue As Byte

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Dim i, j As Integer
    
    'setarea dimensiunilor de pornire.
    Form1.Height = 7140
    Form1.Width = 6135
    Form1.Left = Screen.Width / 2 - Form1.Width / 2
    Form1.Top = Screen.Height / 2 - Form1.Height / 2
    Label1.Visible = False
    Image4.Tag = "1"
    Picture2.Visible = False
    Command1.Left = Form1.Width - Command1.Width - 250
    Form1.Show
    
    'creaza un director Images in directorul in care se afla aplicatia.
    MkDir App.path & "\Images"
    
    refreshMatrix
    'determina in ce mod se afla utilizatorul si calculeaza dimensiunea
    'patratelelor de puzzle.
    If mnubeginer.Checked = True Then
        Image3.Height = 1246
        Image3.Width = 1246
        For i = 0 To 15
            If i <> 0 Then
                Load Image1(i)
            End If
            Image1(i).Height = 1246
            Image1(i).Width = 1246
            Image1(i).Left = (i Mod 4) * 1246
            Image1(i).Top = (i \ 4) * 1246
            Image1(i).Visible = True
        Next i
    End If
    Form1.PictureClip1.Cols = 4
    Form1.PictureClip1.Rows = 4
    initMatrix (1)
    shuffleMatrix (1)
    'in functie de valorile care au rezultat in urma actiunii procedurii
    'shuffleMatrix se vor initializa imaginile.
    For i = 1 To 4
        For j = 1 To 4
            If theMatrix(i, j) <> 15 Then
                Form1.Image1(CreateIndex(i, j, 1)).Picture = Form1.PictureClip1 _
                .GraphicCell(theMatrix(i, j))
            Else
                Form1.Image1(CreateIndex(i, j, 1)).Picture = Nothing
            End If
        Next j
    Next i
    'setarea valorilor pentru aspect.
    Form1.Picture1.Height = Form1.Image1(15).Top + Form1.Image1(15).Height + 80
    Form1.Picture1.Width = Form1.Image1(3).Left + Form1.Image1(7).Width + 80
    Image2.Picture = PictureClip1.Picture
    playerName = ""
    pathDir = ""
    uniqueValue = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'daca indicatorul de ajutor este inca vizibil cand s-a iesit din suprafata de
    'jocatunci se va inactiva.
    If Image3.Visible = True Then
        Image3.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Image1_Click(Index As Integer)
    Dim tempVal As Byte
    
    'verifica modudul in care se afla jucatorul si calculeaza pozitiile pozX si
    'pozY ale matricii theMatrix in functie de parametrul Index al imaginii pe
    'care se afla mouse-ul.
    If mnubeginer.Checked = True Then
        MatrixPos Index + 1, 1
    ElseIf mnuAdvansed.Checked = True Then
        MatrixPos Index + 1, 2
    End If
    If mnubeginer.Checked = True Then
        'daca valoarea matricii in punctele pozX ,pozY calculate anterior este
        'diferita de valoarea 15 pentru modul1 (16 casute) atunci se trece la
        'determinarea mutarii posibile.Daca aceasta nu este posibila nu se va
        'semnala in nici un fel.
        If theMatrix(pozX, pozY) <> 15 Then
            If pozX - 1 > 0 Then
                If theMatrix(pozX - 1, pozY) = 15 Then
                    Image1(CreateIndex(pozX - 1, pozY, 1)).Picture = _
                        Image1(Index).Picture
                    Image1(Index).Picture = Nothing
                    
                    tempVal = theMatrix(pozX, pozY)
                    theMatrix(pozX, pozY) = 15
                    theMatrix(pozX - 1, pozY) = tempVal
                End If
            End If
            If pozX + 1 < 5 Then
                If theMatrix(pozX + 1, pozY) = 15 Then
                    Image1(CreateIndex(pozX + 1, pozY, 1)).Picture = _
                        Image1(Index).Picture
                    Image1(Index).Picture = Nothing
                    
                    tempVal = theMatrix(pozX, pozY)
                    theMatrix(pozX, pozY) = 15
                    theMatrix(pozX + 1, pozY) = tempVal
                End If
            End If
            If pozY - 1 > 0 Then
                If theMatrix(pozX, pozY - 1) = 15 Then
                    Image1(CreateIndex(pozX, pozY - 1, 1)).Picture = _
                        Image1(Index).Picture
                    Image1(Index).Picture = Nothing
                    
                    tempVal = theMatrix(pozX, pozY)
                    theMatrix(pozX, pozY) = 15
                    theMatrix(pozX, pozY - 1) = tempVal
                End If
            End If
            If pozY + 1 < 5 Then
                If theMatrix(pozX, pozY + 1) = 15 Then
                    Image1(CreateIndex(pozX, pozY + 1, 1)).Picture = _
                        Image1(Index).Picture
                    Image1(Index).Picture = Nothing

                    tempVal = theMatrix(pozX, pozY)
                    theMatrix(pozX, pozY) = 15
                    theMatrix(pozX, pozY + 1) = tempVal
                End If
            End If
            End If
            'la fiecare mutare este verificat daca puzzle-ul a fost terminat
            'sau nu
            whenWin (1)
    ElseIf mnuAdvansed.Checked = True Then
        'daca valoarea matricii in punctele pozX ,pozY calculate anterior este
        'diferita de valoarea 63 pentru modul1 (64 casute) atunci se trece la
        'determinarea mutarii posibile.Daca aceasta nu este posibila nu se va
        'semnala in nici un fel.
        If theMatrix(pozX, pozY) <> 63 Then
            If pozX - 1 > 0 Then
                If theMatrix(pozX - 1, pozY) = 63 Then
                    Image1(CreateIndex(pozX - 1, pozY, 2)).Picture = _
                        Image1(Index).Picture
                    Image1(Index).Picture = Nothing
                    
                    tempVal = theMatrix(pozX, pozY)
                    theMatrix(pozX, pozY) = 63
                    theMatrix(pozX - 1, pozY) = tempVal
                End If
            End If
            If pozX + 1 < 9 Then
                If theMatrix(pozX + 1, pozY) = 63 Then
                    Image1(CreateIndex(pozX + 1, pozY, 2)).Picture = _
                        Image1(Index).Picture
                    Image1(Index).Picture = Nothing
                                
                    tempVal = theMatrix(pozX, pozY)
                    theMatrix(pozX, pozY) = 63
                    theMatrix(pozX + 1, pozY) = tempVal
                End If
            End If
            If pozY - 1 > 0 Then
                If theMatrix(pozX, pozY - 1) = 63 Then
                    Image1(CreateIndex(pozX, pozY - 1, 2)).Picture = _
                        Image1(Index).Picture
                    Image1(Index).Picture = Nothing

                    tempVal = theMatrix(pozX, pozY)
                    theMatrix(pozX, pozY) = 63
                    theMatrix(pozX, pozY - 1) = tempVal
                End If
            End If
            If pozY + 1 < 9 Then
                If theMatrix(pozX, pozY + 1) = 63 Then
                    Image1(CreateIndex(pozX, pozY + 1, 2)).Picture = _
                        Image1(Index).Picture
                    Image1(Index).Picture = Nothing
                    
                    tempVal = theMatrix(pozX, pozY)
                    theMatrix(pozX, pozY) = 63
                    theMatrix(pozX, pozY + 1) = tempVal
                End If
            End If
            End If
            'la fiecare mutare este verificat daca puzzle-ul a fost terminat
            'sau nu
            whenWin (2)
    End If
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'se verifica modul in care se afla jucatorul si in functie de mod se
    'calcueaza valorile pozX ,pozY.
    If mnubeginer.Checked = True Then
        MatrixPos Index + 1, 1
    ElseIf mnuAdvansed.Checked = True Then
        MatrixPos Index + 1, 2
    End If
    If mnubeginer.Checked = True Then
        If uniqueValue <> theMatrix(pozX, pozY) Then
            'se determina valoarea indicatorului pe imaginea ajutatoare
            'in junctie de imaginea pe care se afla mouse-ul
            Image3.Visible = True
            uniqueValue = theMatrix(pozX, pozY)
            If theMatrix(pozX, pozY) <> 15 Then
                Image3.Left = ((theMatrix(pozX, pozY) Mod 4) * 1246)
                Image3.Top = ((theMatrix(pozX, pozY) \ 4) * 1246)
            Else
                Image3.Visible = False
            End If
        End If
    ElseIf mnuAdvansed.Checked = True Then
        If uniqueValue <> theMatrix(pozX, pozY) Then
            'se determina valoarea indicatorului pe imaginea ajutatoare
            'in junctie de imaginea pe care se afla mouse-ul
            Image3.Visible = True
            uniqueValue = theMatrix(pozX, pozY)
            If theMatrix(pozX, pozY) <> 63 Then
                Image3.Left = ((theMatrix(pozX, pozY) Mod 8) * 624)
                Image3.Top = ((theMatrix(pozX, pozY) \ 8) * 624)
            Else
                Image3.Visible = False
            End If
        End If
    End If
End Sub

Private Sub Image4_Click()
    
    'daca jucatorul s-a hotarat sa primeasca ajutor se va afisa imaginea
    'ajutatoare si indicatorul.
    If Image4.Tag = "1" Then
        Form1.Width = 11280
        Form1.Left = Screen.Width / 2 - Form1.Width / 2
        Form1.Top = Screen.Height / 2 - Form1.Height / 2
        Command1.Left = Form1.Width - Command1.Width - 250
        Label1.Visible = True
        Picture2.Visible = True
        Image3.Visible = False
        Image4.Tag = "2"
        Image4.Picture = Image5(0).Picture
        Image4.ToolTipText = "Low Details"
    'daca jucatorul nu mai are nevoie de ajutor poate renunta la modul
    'ajutator.
    ElseIf Image4.Tag = "2" Then
        Form1.Width = 6135
        Form1.Left = Screen.Width / 2 - Form1.Width / 2
        Form1.Top = Screen.Height / 2 - Form1.Height / 2
        Command1.Left = Form1.Width - Command1.Width - 250
        Label1.Visible = False
        Picture2.Visible = False
        Image3.Visible = False
        Image4.Tag = "1"
        Image4.Picture = Image5(1)
        Image4.ToolTipText = "More Details"
    End If
    
End Sub

Private Sub mnuAbout_Click()
    MsgBox "           Puzzle Game           " + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Created by Micu Dan-Cristian                 ", vbInformation + vbOKOnly
End Sub

'*****************************************************************************
' Procedura realzizeaza trecerea din modul beginer in modul advanced.
'*****************************************************************************
Private Sub mnuAdvansed_Click()
    Dim i, j, k, l As Integer
    initMatrix (2)
    If Form1.mnuAdvansed.Checked = False Then
        Form1.mnuAdvansed.Checked = True
        Form1.mnubeginer.Checked = False
        For i = 0 To 15
            If (i <> 0) Then
                Unload Form1.Image1(i)
            Else
                Form1.Image1(0).Visible = False
            End If
        Next i
        Form1.Image3.Height = 624
        Form1.Image3.Width = 624
        For i = 0 To 63
            If (i <> 0) Then
                Load Form1.Image1(i)
            End If
            Form1.Image1(i).Height = 624
            Form1.Image1(i).Width = 624
            Form1.Image1(i).Left = (i Mod 8) * 624
            Form1.Image1(i).Top = (i \ 8) * 624
            Form1.Image1(i).Visible = True
        Next i
        Form1.Picture1.Height = Form1.Image1(63).Top + Form1.Image1(63).Height + 80
        Form1.Picture1.Width = Form1.Image1(7).Left + Form1.Image1(7).Width + 80
        If Form1.mnuAdvansed.Checked = True Then
            Form1.PictureClip1.Cols = 8
            Form1.PictureClip1.Rows = 8
            shuffleMatrix (2)
            For k = 1 To 8
                For l = 1 To 8
                    If theMatrix(k, l) <> 63 Then
                        Form1.Image1(CreateIndex(k, l, 2)).Picture = Form1.PictureClip1.GraphicCell(theMatrix(k, l))
                    Else
                        Form1.Image1(CreateIndex(k, l, 2)).Picture = Nothing
                    End If
                Next l
            Next k
        End If
    End If
End Sub


'*****************************************************************************
' Procedura realizeaza trecerea din modul advanced in modul beginer.
'*****************************************************************************
Private Sub mnuBeginer_Click()
    Dim i, j, k, l As Integer
    initMatrix (1)
    If Form1.mnubeginer.Checked = False Then
        Form1.mnubeginer.Checked = True
        Form1.mnuAdvansed.Checked = False
        Form1.Image1(15).Picture = Nothing
        For i = 0 To 63
            If (i <> 0) Then
                Unload Form1.Image1(i)
            Else
                Form1.Image1(0).Visible = False
            End If
        Next i
            Form1.PictureClip1.Cols = 4
            Form1.PictureClip1.Rows = 4
            Form1.Image3.Height = 1246
            Form1.Image3.Width = 1246
            For i = 0 To 15
                If (i <> 0) Then
                    Load Form1.Image1(i)
                End If
                Form1.Image1(i).Height = 1246
                Form1.Image1(i).Width = 1246
                Form1.Image1(i).Left = (i Mod 4) * 1246
                Form1.Image1(i).Top = (i \ 4) * 1246
                Form1.Image1(i).Visible = True
            Next i
            Form1.Picture1.Height = Form1.Image1(15).Top + Form1.Image1(15).Height + 80
            Form1.Picture1.Width = Form1.Image1(3).Left + Form1.Image1(7).Width + 80
            If Form1.mnubeginer.Checked = True Then
                shuffleMatrix (1)
                For k = 1 To 4
                    For l = 1 To 4
                        If theMatrix(k, l) <> 15 Then
                            Form1.Image1(CreateIndex(k, l, 1)).Picture = Form1.PictureClip1.GraphicCell(theMatrix(k, l))
                        Else
                            Form1.Image1(CreateIndex(k, l, 1)).Picture = Nothing
                        End If
                    Next l
                Next k
            End If
        End If
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuNewGame_Click()
    
    'seteaza valorile de design pentru afisarea ferestrei.
    Form2.Left = Screen.Width / 2 - Form2.Width / 2
    Form2.Top = Screen.Height / 2 - Form2.Height / 2
    Form2.Show 1
    
End Sub

'*****************************************************************************
' Proceduta realizeaza amestecul patratelelor chiar si in timpul jocului
'  apasand F2 sau Shuffle din meniul File in functie de modul in care se
'  afla jucatorul.
'*****************************************************************************
Private Sub mnuShuffle_Click()
    Dim i, j, k, l As Byte
    If Picture2.Visible = True Then
        Image3.Visible = False
    End If
    If mnubeginer.Checked = True Then
        shuffleMatrix (1)
            For i = 1 To 4
                For j = 1 To 4
                    If theMatrix(i, j) <> 15 Then
                        Image1(CreateIndex(i, j, 1)).Picture = PictureClip1.GraphicCell(theMatrix(i, j))
                    Else
                        Image1(CreateIndex(i, j, 1)).Picture = Nothing
                    End If
                Next j
            Next i
    ElseIf mnuAdvansed.Checked = True Then
        shuffleMatrix (2)
        For i = 1 To 8
            For j = 1 To 8
                If theMatrix(i, j) <> 63 Then
                    Image1(CreateIndex(i, j, 2)).Picture = PictureClip1.GraphicCell(theMatrix(i, j))
                Else
                    Image1(CreateIndex(i, j, 2)).Picture = Nothing
                End If
            Next j
        Next i
    End If
End Sub
