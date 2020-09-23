VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCodeCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Verification Code Example
'Make this becuase I made a small Messanger program, And wanted somesort of
'Verification Like you see been used most of the time on the net so I came up with this little idea
'anyway hope you like it
'By Dreamvb

Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum Verification
    LettersUpperCase = 0
    LettersLowerCase
    DigitsOnly
    RandomWords
End Enum

Private m_UsePatten As Boolean
Private m_BorderColor As OLE_COLOR
Private m_Backcolor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_VerificationLen As Integer
Private m_JumbleText As Boolean
Private m_Verification As Verification
Private m_PattenImage As IPictureDisp
Private m_PattenBCreated As Boolean
Private hPatten_Brush As Long

Private m_VerificationCode As String

Private m_RandWords As New Collection

Public Function RandWordCount() As Integer
    RandWordCount = m_RandWords.Count
End Function
Public Sub ClearRandomWords()
    Set m_RandWords = Nothing
End Sub

Public Sub AddRandomWord(sWord As String)
    m_RandWords.Add sWord
End Sub

Public Sub RemoveRandomWord(Index As Integer)
    m_RandWords.Remove Index
End Sub

Function RandomWord(Index As Integer) As String
    If (Index > m_RandWords.Count) Then Index = m_RandWords.Count
    RandomWord = m_RandWords.Item(Index)
End Function

Private Function GenPassword(hi As Integer, lo As Integer, Length As Integer) As String
Dim x As Integer, s As String
    'Password generator 1
    Const Chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    
    For x = 1 To Length
        Randomize
        s = s & Mid(Chars, (hi + Int(Rnd * lo) + 1), 1)
    Next x
    
    x = 0
    GenPassword = s
    
End Function

Function GenVerification(iPicBox As PictureBox)
Static OnOff As Boolean, x As Integer, rc As RECT
Dim hi As Integer, lo As Integer

    With iPicBox
        .Cls
        .AutoRedraw = True
        .ScaleMode = 3
        .BackColor = m_Backcolor
        .ForeColor = m_ForeColor
        .FontBold = True
        .BorderStyle = 0

        Select Case m_Verification
            Case LettersUpperCase
                hi = 26: lo = 26
            Case DigitsOnly
                hi = 52: lo = 10
            Case LettersLowerCase
                hi = 0: lo = 26
            Case RandomWords
                hi = 1: lo = m_RandWords.Count
        End Select
        
        If m_VerificationLen = 0 Then m_VerificationLen = 10
        
        If (m_Verification <> RandomWords) Then
            m_VerificationCode = GenPassword(hi, lo, m_VerificationLen)
        Else
            m_VerificationCode = RandomWord(hi + Int(Rnd * lo) + 1)
        End If
        
        .CurrentX = 2
        .Width = (.TextWidth(m_VerificationCode) * Screen.TwipsPerPixelX) + (.CurrentX * 30) * .CurrentX
        .Height = 2 * (.TextHeight(m_VerificationCode) * Screen.TwipsPerPixelX) + 30
        .CurrentY = (.ScaleHeight - .TextHeight(m_VerificationCode)) \ 2
        
        If (Not m_PattenBCreated) And UsePatten Then
            hPatten_Brush = CreatePatternBrush(m_PattenImage)
            m_PattenBCreated = True
        End If
        
        If (UsePatten) Then
            rc.Top = 0: rc.Bottom = .ScaleHeight
            rc.Left = 0: rc.Right = .ScaleWidth
            FillRect .hdc, rc, hPatten_Brush
        End If
        
        For x = 1 To Len(m_VerificationCode)
            If m_JumbleText Then
                c = Sin(x * Rnd(.CurrentY)) + 8
            Else
                c = 0
            End If
            
            If (OnOff) Then
                .CurrentY = .CurrentY + c
            Else
                .CurrentY = .CurrentY - c
            End If
            
            iPicBox.Print Mid$(m_VerificationCode, x, 1);
            OnOff = (Not OnOff)
        Next x
        
        iPicBox.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), m_BorderColor, B
        .Tag = m_VerificationCode
        
    End With
    
End Function

Function VerificationGood(iCheck As String) As Boolean
    VerificationGood = (m_VerificationCode = iCheck)
End Function

Public Property Get Patten() As IPictureDisp
    Patten = m_PattenImage
End Property

Public Property Let Patten(ByVal vNewValue As IPictureDisp)
    Set m_PattenImage = vNewValue
    m_PattenBCreated = False
End Property

Public Property Get VerificationType() As Verification
    VerificationType = m_Verification
End Property

Public Property Let VerificationType(ByVal vNewValue As Verification)
    m_Verification = vNewValue
End Property

Public Property Get JumbleText() As Boolean
    JumbleText = m_JumbleText
End Property

Public Property Let JumbleText(ByVal vNewValue As Boolean)
    m_JumbleText = vNewValue
End Property

Public Property Get VerificationLength() As Integer
    VerificationLength = m_VerificationLen
End Property

Public Property Let VerificationLength(ByVal vNewValue As Integer)
    m_VerificationLen = vNewValue
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    m_ForeColor = vNewValue
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_Backcolor
End Property

Public Property Let BackColor(ByVal vNewValue As Variant)
    m_Backcolor = vNewValue
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal vNewValue As OLE_COLOR)
    m_BorderColor = vNewValue
End Property

Public Property Get UsePatten() As Boolean
    UsePatten = m_UsePatten
End Property

Public Property Let UsePatten(ByVal vNewValue As Boolean)
    m_UsePatten = vNewValue
End Property

Private Sub Class_Terminate()
    Set m_RandWords = Nothing
    DeleteObject hPatten_Brush
End Sub

