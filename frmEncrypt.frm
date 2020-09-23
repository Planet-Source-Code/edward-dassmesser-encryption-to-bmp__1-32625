VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Picture Encryption"
   ClientHeight    =   6120
   ClientLeft      =   1035
   ClientTop       =   1320
   ClientWidth     =   13740
   Icon            =   "frmEncrypt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   13740
   Begin VB.CommandButton cmdLoadBMP 
      Caption         =   "Load .BMP"
      Height          =   195
      Left            =   12360
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveBMP 
      Caption         =   "Save .BMP"
      Height          =   195
      Left            =   11040
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save to Text"
      Height          =   195
      Left            =   9480
      TabIndex        =   5
      Top             =   5880
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog loadfile 
      Left            =   4080
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load File"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt Text"
      Height          =   195
      Left            =   8280
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt Text"
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.PictureBox picEncrypted 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   7080
      ScaleHeight     =   383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   439
      TabIndex        =   2
      Top             =   0
      Width           =   6615
   End
   Begin VB.TextBox txtEncrypt 
      Height          =   5775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Menu mnuStuff 
      Caption         =   "&Stuff you might want to do"
      Begin VB.Menu mnuVote 
         Caption         =   "&Vote at Planet Source Code"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "E-mail Author (durnurd@hotmail.com)"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEncrypt_Click()
    Dim P As PictureBox, T As TextBox, L As Long, A As Integer, B As Integer
    Dim C As Integer, D As Integer, E As Integer, Red As Integer
    Dim Green As Integer, Blue As Integer
    Set P = picEncrypted
    Set T = txtEncrypt
    L = Len(T.Text)
    C = Int(Sqr(L)): D = C
    E = 0
    P.Cls
    P.Picture = LoadPicture("")
    P.Width = C * Screen.TwipsPerPixelX + 200
    P.Height = D * Screen.TwipsPerPixelY + 200
    For A = 1 To C
        For B = 1 To D
            E = E + 1
            Red = 0
            Green = 0
            Blue = 0
            If E > L Then Exit For
            Red = Asc(Mid(T.Text, E, 1))
            E = E + 1
            If E > L Then GoTo MidColor
            Green = Asc(Mid(T.Text, E, 1))
            E = E + 1
            If E > L Then GoTo MidColor
            Blue = Asc(Mid(T.Text, E, 1))
MidColor:   P.PSet (A, B), RGB(Red, Green, Blue)
        Next B
    Next A
    SavePicture picEncrypted.Image, "C:\Encrypted File.bmp"
End Sub

Private Sub cmdDecrypt_Click()
    Dim P As PictureBox, T As TextBox, L As Integer, A As Integer, B As Integer
    Dim C As Integer, D As Integer, E As Long, Red As Integer
    Dim Green As Integer, Blue As Integer
    Set P = picEncrypted
    Set T = txtEncrypt
    On Error Resume Next
    P.Cls
    P.Picture = LoadPicture("C:\Encrypted File.bmp")
    C = P.ScaleWidth
    D = C
    T.Text = ""
    For A = 1 To C
        For B = 1 To D
            E = P.Point(A, B)
            Red = 0
            Blue = 0
            Green = 0
            If E = -1 Then Exit For
            Blue = Int(E / 65536)
            If Blue = 255 Then Exit For
            Green = Int((E - (Blue * 65536)) / 256)
            If Green = 255 Then Exit For
            Red = Int(E - Blue * 65536 - Green * 256)
            If Red = 255 Then Exit For
            T.Text = T.Text & Chr$(Red) & Chr$(Green) & Chr$(Blue)
        Next B
    Next A
End Sub

Private Sub cmdLoad_Click()
    On Error Resume Next
    loadfile.ShowOpen
    Dim P As PictureBox, T As TextBox, L As Long, A As Integer, B As Integer
    Dim C As Integer, D As Integer, E As Integer, Red As Integer
    Dim Green As Integer, Blue As Integer, TxtLoaded As String, TempTxt As String
    Open loadfile.FileName For Input As #1
        Do While Not EOF(1)
            Line Input #1, TempTxt
            TxtLoaded = TxtLoaded & vbCrLf & TempTxt
        Loop
    Close #1
    TxtLoaded = Right(TxtLoaded, Len(TxtLoaded) - 2)
    Set P = picEncrypted
    L = Len(TxtLoaded)
    C = Int(Sqr(L)): D = C
    E = 0
    P.Cls
    P.Picture = LoadPicture("")
    P.Width = C * Screen.TwipsPerPixelX + 200
    P.Height = D * Screen.TwipsPerPixelY + 200
    For A = 1 To C
        For B = 1 To D
            E = E + 1
            Red = 0
            Green = 0
            Blue = 0
            If E > L Then Exit For
            Red = Asc(Mid(TxtLoaded, E, 1))
            E = E + 1
            If E > L Then GoTo MidColor
            Green = Asc(Mid(TxtLoaded, E, 1))
            E = E + 1
            If E > L Then GoTo MidColor
            Blue = Asc(Mid(TxtLoaded, E, 1))
MidColor:   P.PSet (A, B), RGB(Red, Green, Blue)
        Next B
    Next A
    SavePicture picEncrypted.Image, "C:\Encrypted File.bmp"
End Sub

Private Sub cmdLoadBMP_Click()
    Dim P As PictureBox, T As TextBox, L As Integer, A As Integer, B As Integer
    Dim C As Integer, D As Integer, E As Long, Red As Integer
    Dim Green As Integer, Blue As Integer
    Set P = picEncrypted
    Set T = txtEncrypt
    On Error Resume Next
    P.Cls
    loadfile.Filter = "*.bmp|*.bmp"
    loadfile.ShowOpen
    If loadfile.FileName = "" Then Exit Sub
    picEncrypted.Picture = LoadPicture(loadfile.FileName)
    loadfile.Filter = ""
    C = P.ScaleWidth
    D = C
    T.Text = ""
    For A = 1 To C
        For B = 1 To D
            E = P.Point(A, B)
            Red = 0
            Blue = 0
            Green = 0
            If E = -1 Then Exit For
            Blue = Int(E / 65536)
            If Blue = 255 Then Exit For
            Green = Int((E - (Blue * 65536)) / 256)
            If Green = 255 Then Exit For
            Red = Int(E - Blue * 65536 - Green * 256)
            If Red = 255 Then Exit For
            T.Text = T.Text & Chr$(Red) & Chr$(Green) & Chr$(Blue)
        Next B
    Next A
End Sub

Private Sub cmdSave_Click()
Dim P As PictureBox, T As TextBox, L As Integer, A As Integer, B As Integer
    Dim C As Integer, D As Integer, E As Long, Red As Integer
    Dim Green As Integer, Blue As Integer, txtSaved As String
    loadfile.ShowSave
    Set P = picEncrypted
    On Error Resume Next
    P.Cls
    P.Picture = LoadPicture("C:\Encrypted File.bmp")
    C = P.ScaleWidth
    D = C
    Open loadfile.FileName For Output As #1
    For A = 1 To C
        For B = 1 To D
            E = P.Point(A, B)
            Red = 0
            Blue = 0
            Green = 0
            If E = -1 Then Exit For
            Blue = Int(E / 65536)
            If Blue = 255 Then Exit For
            Green = Int((E - (Blue * 65536)) / 256)
            If Green = 255 Then Exit For
            Red = Int(E - Blue * 65536 - Green * 256)
            If Red = 255 Then Exit For
            txtSaved = txtSaved & Chr$(Red) & Chr$(Green) & Chr$(Blue)
        Next B
    Next A
    Print #1, txtSaved
    Close #1
End Sub

Private Sub cmdSaveBMP_Click()
    loadfile.Filter = "*.bmp|*.bmp"
    loadfile.ShowSave
    If loadfile.FileName = "" Then Exit Sub
    SavePicture picEncrypted.Image, loadfile.FileName
    loadfile.Filter = ""
End Sub
Private Sub mnuEmail_Click()
    StartURL "mailto:durnurd@hotmail.com"
End Sub
Private Sub mnuVote_Click()
    StartURL "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=32601&lngWId=1"
End Sub
Private Sub StartURL(strURL As String)
    On Error Resume Next
    Shell "Explorer """ & strURL & """"
    If Err.Number <> 0 Then
        Err.Clear
        Shell "Start """ & strURL & """"
    End If
    If Err.Number <> 0 Then
        If MsgBox("Can't figure out how to navigate on this OS.  Copy the URL to the clipboard?", vbExclamation + vbYesNo) = vbYes Then
            Clipboard.SetText strURL
        End If
    End If
End Sub
