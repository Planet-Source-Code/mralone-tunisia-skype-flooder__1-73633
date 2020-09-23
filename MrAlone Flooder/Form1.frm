VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MrAlone Flooder v1.1"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":29C12
   ScaleHeight     =   3720
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox max 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Text            =   "5"
      Top             =   1680
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   60
      Left            =   19080
      TabIndex        =   3
      Top             =   11520
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   10080
      Top             =   7200
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start / Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton btConvertir 
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   4560
      Picture         =   "Form1.frx":4EC94
      Top             =   240
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   4560
      Picture         =   "Form1.frx":50106
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Per Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ye5dem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''OMl''''':MO''''''''''OMl'''''M:''''''''''''''''''''''''''''''''''''
''''''''''''''''''OMM'''''OMM''''''''''MOM'''''Ml''''''''''''''''''''''''''''''''''''
''''''''''''''''''OOM:''''MOM'''''''''lM:M:''''M:''''''''''''''''''''''''''''''''''''
''''''''''''''''''OOOl''':MOM''MlMM'''OO'OO''''Ml'':MMMl'''MlMMM:''':MMMl''''''''''''
''''''''''''''''''OO:M'''OlOM''MO'''':M:'lM:'''Ml':Ml''MO''MM:'lM:'lM'''Ml'''''''''''
''''''''''''''''''OO'M'''O:lM''Ml''''OM'''Ml'''M:'Ol''':M:'Ml'''M:'Ml''':M'''''''''''
''''''''''''''''''OO'Ol':M'OM''M:''':Ml'''OM'''Ml'Ml''''M:'Ml'''Ml:Ml''''Ml''''''''''
''''''''''''''''''OO'lO'll'OM''M:''':MMMMMMM:''M:'M:''''Ml'Ml'''M::MMMMMMM:''''''''''
''''''''''''''''''OO''M'O:'OM''M:'''OO'''''MO''M::M:''''M:'M:'''Ml'M'''''''''''''''''
''''''''''''''''''OO''OlM''OM''M:'':M:''''':M:'M:'Ol''':M''M:'''M:'MO''''''''''''''''
''''''''''''''''''OO''OMO''OM''Ml''OM'''''''MO'Ml':M:''MO''M:'''M:'lM:''lM:''''''''''
''''''''''''''''''OO'':M:''OM''M:''Ml'''''''lM:M:'':MMMl'''M:'''M:'':MMMM''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub btConvertir_Click()

    Dim a As Long
    Dim tmp As String
    
    tmp = ""
    
    If List1.ListCount <= 0 Then
        Exit Sub
    End If
    
    For a = 1 To Val(max.text)
        tmp = tmp & List1.List(Round(Rnd * List1.ListCount, 0))
    Next
    
    Clipboard.Clear
    Clipboard.SetText tmp

End Sub


Private Sub Check1_Click()
    Timer1.Enabled = Check1.Value
    Label6.Visible = Check1.Value
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Form_Activate()

    Dim fso As New FileSystemObject
    Dim text As TextStream
    Dim a As Integer
    
    On Error GoTo erreur
    Set text = fso.OpenTextFile("AloneSmiley.mp3", ForReading, False)
    
    While Not text.AtEndOfStream
        List1.AddItem text.ReadLine
    Wend
    
    Randomize
    
    Exit Sub
erreur:
    MsgBox "File ""AloneSmiley.mp3"" not Found !"
End Sub

Private Sub Form_Load()
    Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "You can change it as you like but Please Vote For me, Thanks..." & vbCrLf & "by MrAlone -"
    
End Sub

Private Sub Timer1_Timer()
    btConvertir_Click
    
    SendKeys "^V"
    SendKeys "{ENTER}"
End Sub

Private Sub topbuttons1_Click()
windowstat = 2
End Sub

Private Sub topbuttons2_Click()
Unload Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''OMl''''':MO''''''''''OMl'''''M:''''''''''''''''''''''''''''''''''''
''''''''''''''''''OMM'''''OMM''''''''''MOM'''''Ml''''''''''''''''''''''''''''''''''''
''''''''''''''''''OOM:''''MOM'''''''''lM:M:''''M:''''''''''''''''''''''''''''''''''''
''''''''''''''''''OOOl''':MOM''MlMM'''OO'OO''''Ml'':MMMl'''MlMMM:''':MMMl''''''''''''
''''''''''''''''''OO:M'''OlOM''MO'''':M:'lM:'''Ml':Ml''MO''MM:'lM:'lM'''Ml'''''''''''
''''''''''''''''''OO'M'''O:lM''Ml''''OM'''Ml'''M:'Ol''':M:'Ml'''M:'Ml''':M'''''''''''
''''''''''''''''''OO'Ol':M'OM''M:''':Ml'''OM'''Ml'Ml''''M:'Ml'''Ml:Ml''''Ml''''''''''
''''''''''''''''''OO'lO'll'OM''M:''':MMMMMMM:''M:'M:''''Ml'Ml'''M::MMMMMMM:''''''''''
''''''''''''''''''OO''M'O:'OM''M:'''OO'''''MO''M::M:''''M:'M:'''Ml'M'''''''''''''''''
''''''''''''''''''OO''OlM''OM''M:'':M:''''':M:'M:'Ol''':M''M:'''M:'MO''''''''''''''''
''''''''''''''''''OO''OMO''OM''Ml''OM'''''''MO'Ml':M:''MO''M:'''M:'lM:''lM:''''''''''
''''''''''''''''''OO'':M:''OM''M:''Ml'''''''lM:M:'':MMMl'''M:'''M:'':MMMM''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
