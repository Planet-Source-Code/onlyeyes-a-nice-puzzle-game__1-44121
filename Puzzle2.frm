VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Puzzle Game"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   5760
      Pattern         =   "*.jpg"
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Advanced"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Beginer"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   6000
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Image"
      FileName        =   "*.jpg"
      Filter          =   "JPEG Pictures(*.jpg)"
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   1110
      Width           =   350
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Text            =   "Player Name"
      Top             =   480
      Width           =   3975
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   960
      Picture         =   "Puzzle2.frx":0000
      ToolTipText     =   "Previews Picture"
      Top             =   5640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   4200
      Picture         =   "Puzzle2.frx":0312
      ToolTipText     =   "Next Picture"
      Top             =   5640
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Width x Height:"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Dificulty"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   1050
   End
   Begin MSForms.Image Image1 
      Height          =   3015
      Left            =   960
      Top             =   2520
      Width           =   3495
      AutoSize        =   -1  'True
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6165;5318"
      PictureTiling   =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image Preview"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Image"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player Name :"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fileIndex As Integer

'se va incarca imaginea aleasa din lista de imagini.
Private Sub Combo1_Click()
    Image1.Picture = LoadPicture(Combo1.List(Combo1.ListIndex))
End Sub

'*****************************************************************************
' Daca se evita cautarea si se stie calea se poate tasta urmata de Enter.
'  In cazul in care calea este invalida se va semnala acest lucrul vizual.
'*****************************************************************************
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim fileName As String
    If (KeyAscii = 13 And Combo1.Text <> "") Then
        Image1.Picture = LoadPicture(Combo1.Text)
        Form1.PictureClip1.Picture = Image1.Picture
        Label5.Caption = "Height x Width:" & Form1.PictureClip1.Height & " x " & _
            Form1.PictureClip1.Width
        Combo1.AddItem Combo1.Text
        Command2.Enabled = True
        playerName = Trim(Text1.Text)
        File1.path = obtainPath(Combo1.Text, fileName)
        Image2.Visible = False
        Image3.Visible = False
        If File1.ListCount > 1 Then
            For i = 0 To File1.ListCount
                If LCase(File1.List(i)) = LCase(Combo1.Text) Then
                    fileIndex = i
                    Exit For
                End If
            Next i
            If (fileIndex > 0 And fileIndex < File1.ListCount - 1) Then
                Image2.Visible = True
                Image3.Visible = True
            ElseIf fileIndex = 0 Then
                Image2.Visible = True
                Image3.Visible = False
            ElseIf fileIndex = File1.ListCount - 1 Then
                Image2.Visible = False
                Image3.Visible = True
            End If
        End If
    End If
    Exit Sub
    
ErrorHandle:
    MsgBox "Invalid path. File " & Combo1.Text & " was not found", vbCritical + vbOKOnly
    Combo1.Text = ""
    
End Sub

Private Sub Command1_Click()
    On Error GoTo ErrorHandle
    
    Dim fileName As String
    Dim i As Integer
    'seteaza filtrul pentru fisiere.
    CommonDialog1.fileName = "*.jpg"
    CommonDialog1.ShowOpen
    'se afiseaza calea si numele fisierului.
    Combo1.Text = Trim(CommonDialog1.fileName)
    File1.path = obtainPath(CommonDialog1.fileName, fileName)
    'daca s-a introdus o cale atunci se incearca incarcarea imaginii.
    If Combo1.Text <> "" Then
        Image1.Picture = LoadPicture(Combo1.Text)
        Form1.PictureClip1.Picture = Image1.Picture
        Label5.Caption = "Width x Height:" & Form1.PictureClip1.Width & " x " & _
            Form1.PictureClip1.Height
        pathDir = Combo1.Text
        Combo1.AddItem Combo1.Text
        Command2.Enabled = True
        playerName = Trim(Text1.Text)
        Image2.Visible = False
        Image3.Visible = False
        If File1.ListCount > 1 Then
            For i = 0 To File1.ListCount
                If LCase(File1.List(i)) = LCase(fileName) Then
                    fileIndex = i
                    Exit For
                End If
            Next i
            If (fileIndex > 0 And fileIndex < File1.ListCount - 1) Then
                Image2.Visible = True
                Image3.Visible = True
            ElseIf fileIndex = 0 Then
                Image2.Visible = True
                Image3.Visible = False
            ElseIf fileIndex = File1.ListCount - 1 Then
                Image2.Visible = False
                Image3.Visible = True
            End If
        End If
    End If
    Exit Sub
    
ErrorHandle:
    Combo1.Text = ""
    
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim i, j As Integer
    'daca toate campurile necesare au fost completate atunci se incearca
    'deschiderea unui joc nou.
    If (Text1.Text <> "") And (Combo1.Text <> "") Then
        Form1.PictureClip1.Picture = Nothing
        Form1.PictureClip1.Picture = Image1.Picture
        Form1.Image2.Picture = Image1.Picture
        If Option1.Value = True Then
            Form1.mnubeginer.Checked = True
            Form1.mnuAdvansed.Checked = False
        ElseIf Option2.Value = True Then
            Form1.mnuAdvansed.Checked = True
            Form1.mnubeginer.Checked = False
        End If
        If Form1.mnubeginer.Checked = True Then
           Form1.Image1(15).Picture = Nothing
            For i = 0 To 63
                If (i <> 0) Then
                    Unload Form1.Image1(i)
                Else
                    Form1.Image1(0).Visible = False
                End If
            Next i
            For i = 0 To 15
                If (i <> 0) Then
                    Load Form1.Image1(i)
                End If
                Form1.Image1(i).Height = 1246
                Form1.Image1(i).Width = 1246
                Form1.Image1(i).Left = ((i Mod 4) * 1246)
                Form1.Image1(i).Top = ((i \ 4) * 1246)
                Form1.Image1(i).Visible = True
            Next i
            Form1.PictureClip1.Cols = 4
            Form1.PictureClip1.Rows = 4
            initMatrix (1)
            shuffleMatrix (1)
            For i = 1 To 4
                For j = 1 To 4
                    If theMatrix(i, j) <> 15 Then
                        Form1.Image1(CreateIndex(i, j, 1)).Picture = Form1.PictureClip1.GraphicCell(theMatrix(i, j))
                   Else
                        Form1.Image1(CreateIndex(i, j, 1)).Picture = Nothing
                    End If
                Next j
            Next i
            Form1.Picture1.Height = Form1.Image1(15).Top + Form1.Image1(15).Height + 80
            Form1.Picture1.Width = Form1.Image1(3).Left + Form1.Image1(7).Width + 80
        ElseIf Form1.mnuAdvansed.Checked = True Then
            For i = 0 To 63
                If (i <> 0) Then
                    Unload Form1.Image1(i)
                Else
                    Form1.Image1(0).Visible = False
                End If
            Next i
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
            Form1.PictureClip1.Cols = 8
            Form1.PictureClip1.Rows = 8
            initMatrix (2)
            shuffleMatrix (2)
            For i = 1 To 8
                For j = 1 To 8
                    If theMatrix(i, j) <> 63 Then
                        Form1.Image1(CreateIndex(i, j, 2)).Picture = Form1.PictureClip1.GraphicCell(theMatrix(i, j))
                    Else
                        Form1.Image1(CreateIndex(i, j, 2)).Picture = Nothing
                    End If
                Next j
            Next i
            Form1.Picture1.Height = Form1.Image1(63).Top + Form1.Image1(63).Height + 80
            Form1.Picture1.Width = Form1.Image1(7).Left + Form1.Image1(7).Width + 80
        End If
        
        Form1.mnubeginer.Enabled = True
        Form1.mnuAdvansed.Enabled = True
        Form1.mnushuffle.Enabled = True
        Form1.Caption = ""
        Form1.Caption = "Puzzle Games" & " [ " & Trim(Text1.Text) & " ] "
        Unload Me
    ElseIf (Text1.Text = "") Then
        MsgBox "Enter player name please", vbInformation + vbOKOnly
        Text1.SetFocus
    ElseIf Combo1.Text = "" Then
        MsgBox "Enter a path for image you whant to play with", vbInformation + vbOKOnly
        Combo1.SetFocus
    End If
End Sub

Private Sub Form_Load()
    'initializarea directorului de pornire si completarea datelor pentru prima
    'afisare si nu numai.
    CommonDialog1.InitDir = App.path & "\Images"
    If playerName <> "" Then
        Text1.Text = playerName
    End If
    If pathDir <> "" Then
        Combo1.Text = pathDir
        Image1.Picture = Form1.PictureClip1.Picture
        Label5.Caption = "Width x Height:" & Form1.PictureClip1.Width & " x " & _
            Form1.PictureClip1.Height
        Text1.Text = playerName
        Command2.Enabled = True
    End If
End Sub

Private Sub Image2_Click()
    
    'realizeaza derularea imaginilor inainte fara a mai fi nevoie sa scrieti
    'calea pentru fiecare.
    If fileIndex + 1 < File1.ListCount Then
        Image3.Visible = True
        fileIndex = fileIndex + 1
        Combo1.Text = File1.path & "\" & File1.List(fileIndex)
        Image1.Picture = LoadPicture(File1.path & "\" & File1.List(fileIndex))
        Form1.PictureClip1.Picture = Image1.Picture
        Label5.Caption = "Width x Height:" & Form1.PictureClip1.Width & " x " & _
            Form1.PictureClip1.Height
        If fileIndex = File1.ListCount - 1 Then
            Image2.Visible = False
        End If
    End If
End Sub

Private Sub Image3_Click()
    
    'realizeaza derularea imaginilor inapoi fara a mai fi nevoie sa scrieti
    'calea pentru fiecare.
    If fileIndex - 1 >= 0 Then
        Image2.Visible = True
        fileIndex = fileIndex - 1
        Combo1.Text = File1.path & "\" & File1.List(fileIndex)
        Image1.Picture = LoadPicture(File1.path & "\" & File1.List(fileIndex))
        Form1.PictureClip1.Picture = Image1.Picture
        Label5.Caption = "Width x Height:" & Form1.PictureClip1.Width & " x " & _
            Form1.PictureClip1.Height
        If fileIndex = 0 Then
            Image3.Visible = False
        End If
    End If
End Sub

Private Sub Text1_Click()
    'faciliteaza scrierea numelui jucatorului.
    If Text1.Text = "Player Name" Then
        Text1.Text = ""
    End If
End Sub
