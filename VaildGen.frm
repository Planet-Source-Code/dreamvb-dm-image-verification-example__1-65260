VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Verification"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   420
      Left            =   2805
      TabIndex        =   13
      Top             =   3435
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   420
      Left            =   1410
      TabIndex        =   12
      Top             =   3435
      Width           =   1290
   End
   Begin VB.TextBox txtLen 
      Height          =   285
      Left            =   1665
      TabIndex        =   11
      Text            =   "8"
      Top             =   2235
      Width           =   750
   End
   Begin VB.TextBox txtCode 
      Height          =   330
      Left            =   4065
      TabIndex        =   9
      Top             =   150
      Width           =   1695
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check Code"
      Height          =   420
      Left            =   4035
      TabIndex        =   8
      Top             =   525
      Width           =   1755
   End
   Begin VB.CheckBox chkJump 
      Caption         =   "Show Jumpy Text"
      Height          =   270
      Left            =   3450
      TabIndex        =   6
      Top             =   1830
      Width           =   1815
   End
   Begin VB.CommandButton cmdGenCode 
      Caption         =   "Generate Code"
      Height          =   420
      Left            =   30
      TabIndex        =   5
      Top             =   3435
      Width           =   1290
   End
   Begin VB.ComboBox CboBack 
      Height          =   315
      Left            =   2025
      TabIndex        =   4
      Top             =   1830
      Width           =   1290
   End
   Begin VB.ComboBox CboTypeA 
      Height          =   315
      Left            =   285
      TabIndex        =   2
      Top             =   1830
      Width           =   1650
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   210
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   6030
      Picture         =   "VaildGen.frx":0000
      Top             =   3405
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image ImgPatten 
      Height          =   450
      Index           =   5
      Left            =   6030
      Picture         =   "VaildGen.frx":155C
      Top             =   3405
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Verification Length"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2295
      Width           =   1320
   End
   Begin VB.Image ImgPatten 
      Height          =   450
      Index           =   4
      Left            =   6030
      Picture         =   "VaildGen.frx":2AB8
      Top             =   2850
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Patten Styles"
      Height          =   195
      Left            =   6015
      TabIndex        =   7
      Top             =   180
      Width           =   930
   End
   Begin VB.Image ImgPatten 
      Height          =   450
      Index           =   3
      Left            =   6030
      Picture         =   "VaildGen.frx":4014
      Top             =   2310
      Width           =   900
   End
   Begin VB.Image ImgPatten 
      Height          =   450
      Index           =   2
      Left            =   6030
      Picture         =   "VaildGen.frx":556E
      Top             =   1665
      Width           =   900
   End
   Begin VB.Image ImgPatten 
      Height          =   450
      Index           =   1
      Left            =   6030
      Picture         =   "VaildGen.frx":6AC8
      Top             =   1035
      Width           =   900
   End
   Begin VB.Image ImgPatten 
      Height          =   450
      Index           =   0
      Left            =   6030
      Picture         =   "VaildGen.frx":8022
      Top             =   435
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Back Style"
      Height          =   195
      Left            =   2025
      TabIndex        =   3
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Verification Type"
      Height          =   195
      Left            =   285
      TabIndex        =   1
      Top             =   1560
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CodeCheck As New ClsCodeCheck
Dim VType As Integer, JumpyText As Boolean

Function FixPath(lPath As String) As String
    If Right(lPath, 1) <> "\" Then
        FixPath = lPath & "\"
    Else
        FixPath = lPath
    End If
End Function

Sub GenVaildID()

    With CodeCheck
        .Patten = Image1.Picture
        .VerificationType = VType
        .BorderColor = vbYellow
        .ForeColor = vbBlack
        .JumbleText = chkJump
        .VerificationLength = Val(txtLen.Text)
        .UsePatten = CBool(CboBack.ListIndex)
    End With
    
    CodeCheck.GenVerification Picture1
End Sub

Private Sub Command2_Click()
    Unload Form1
End Sub

Private Sub CboBack_Click()
    cmdGenCode_Click
End Sub

Private Sub CboTypeA_Click()
    VType = CboTypeA.ListIndex
    cmdGenCode_Click
End Sub

Private Sub Command1_Click()
    MsgBox "Verification Code Checker" _
    & vbCrLf & vbTab & " By Ben Jones", vbInformation, "About"
End Sub

Private Sub chkJump_Click()
    cmdGenCode_Click
End Sub

Private Sub cmdCheck_Click()
    If Not CodeCheck.VerificationGood(txtCode.Text) Then
        MsgBox "The code you entered does not match the Verification Image." _
        & vbCrLf & "Please try agian.", vbCritical Or vbExclamation, "Verification Faild"
        Exit Sub
    Else
        MsgBox "That was the correct verification Code.", vbInformation, "Verification Good"
    End If
    
End Sub

Private Sub cmdGenCode_Click()
    GenVaildID
End Sub

Private Sub Form_Load()
    
    'Add some random words
    CodeCheck.AddRandomWord "Piggy"
    CodeCheck.AddRandomWord "GoGoTa Go"
    CodeCheck.AddRandomWord "YeYa He"
    CodeCheck.AddRandomWord "Wlolla"
    CodeCheck.AddRandomWord "IcyTicky"
    CodeCheck.AddRandomWord "HonkeyTonkey"
    CodeCheck.AddRandomWord "Wonkey Donkey"
    CodeCheck.AddRandomWord "AveIT"
    CodeCheck.AddRandomWord "Borland"
    CodeCheck.AddRandomWord "VerificationCode"
    CodeCheck.AddRandomWord "HickyDickyDot"
    CodeCheck.AddRandomWord "AveIT"
    CodeCheck.AddRandomWord "NumSkull"
    CodeCheck.AddRandomWord "Idiot2009"
    CodeCheck.AddRandomWord "Yo,xyz"
    CodeCheck.AddRandomWord "Qwerty"
    CodeCheck.AddRandomWord "£%^&*()5465"
    '
    
    '
    CboTypeA.AddItem "LettersUpperCase"
    CboTypeA.AddItem "LettersLowerCase"
    CboTypeA.AddItem "DigitsOnly"
    CboTypeA.AddItem "RandomWords"
    CboTypeA.ListIndex = 0
    '
    CboBack.AddItem "BackColor"
    CboBack.AddItem "Patten"
    CboBack.ListIndex = 1
    
End Sub

Private Sub ImgPatten_Click(Index As Integer)
Dim sFile As String
    sFile = FixPath(App.Path) & "Textures\" & Index + 1 & ".bmp"
    Image1.Picture = LoadPicture(sFile)
    sFile = ""
    Call cmdGenCode_Click
End Sub

Private Sub Picture1_Click()
    MsgBox Picture1.Tag
End Sub
