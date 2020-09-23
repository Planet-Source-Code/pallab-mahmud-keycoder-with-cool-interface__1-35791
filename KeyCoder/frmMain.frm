VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2070
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3060
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   300
      Width           =   2505
   End
   Begin VB.PictureBox piBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   -15
      ScaleHeight     =   255
      ScaleWidth      =   6990
      TabIndex        =   0
      Top             =   -15
      Width           =   7020
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "KeyCoder....[Generate Keycode]"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   15
         Width           =   2400
      End
   End
   Begin VB.Line lnDesign 
      BorderColor     =   &H80000010&
      Index           =   6
      X1              =   3075
      X2              =   3075
      Y1              =   315
      Y2              =   2040
   End
   Begin VB.Label cmdButton 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   5
      Left            =   5055
      MouseIcon       =   "frmMain.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1845
      Width           =   375
   End
   Begin VB.Label cmdButton 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copy"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   4
      Left            =   3165
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   1845
      Width           =   375
   End
   Begin VB.Label cmdButton 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   120
      Index           =   3
      Left            =   5640
      MouseIcon       =   "frmMain.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   975
      Width           =   180
   End
   Begin VB.Line lnDesign 
      BorderColor     =   &H80000010&
      Index           =   5
      X1              =   2235
      X2              =   2235
      Y1              =   1350
      Y2              =   1515
   End
   Begin VB.Line lnDesign 
      BorderColor     =   &H80000010&
      Index           =   4
      X1              =   2265
      X2              =   2265
      Y1              =   1350
      Y2              =   1515
   End
   Begin VB.Line lnDesign 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   1335
      X2              =   1335
      Y1              =   1350
      Y2              =   1515
   End
   Begin VB.Line lnDesign 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   1305
      X2              =   1305
      Y1              =   1350
      Y2              =   1515
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pallmahmud@yahoo.com"
      ForeColor       =   &H80000001&
      Height          =   195
      Index           =   4
      Left            =   570
      TabIndex        =   11
      Top             =   1635
      Width           =   1785
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2002 by Pallab Mahmud"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   10
      Top             =   1830
      Width           =   2640
   End
   Begin VB.Line lnDesign 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   165
      X2              =   3000
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   1410
      TabIndex        =   9
      Top             =   780
      Width           =   705
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   1410
      TabIndex        =   8
      Top             =   1035
      Width           =   705
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KeyCode :"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   7
      Top             =   795
      Width           =   750
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pressed   :"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   6
      Top             =   1035
      Width           =   765
   End
   Begin VB.Label lblCap 
      BackStyle       =   0  'Transparent
      Caption         =   "Press a Key to know the KeyCode of the key you entered."
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   330
      Width           =   2775
   End
   Begin VB.Line lnDesign 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   90
      X2              =   2925
      Y1              =   1545
      Y2              =   1545
   End
   Begin VB.Label cmdButton 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Key Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   135
      MouseIcon       =   "frmMain.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1320
      Width           =   1110
   End
   Begin VB.Label cmdButton 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code ?"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   1
      Left            =   1545
      MouseIcon       =   "frmMain.frx":0C28
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label cmdButton 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   2
      Left            =   2505
      MouseIcon       =   "frmMain.frx":0F32
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1320
      Width           =   270
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'|------------------------------------------------|'
'|KeyCoder[Generate KeyCode !]                    |'
'|------------------------------------------------|'
'|Written by Pallab Mahmud                        |'
'|© Copyright 2001 by Pallab Mahmud               |'
'|email: pallmahmud@yahoo.com                     |'
'|                                                |'
'|This sample code is a FREEWARE. Use it in your  |'
'|own project as it fits You but do not re-sale   |'
'|this code or destroy the original authors name. |'
'|                                                |'
'|Warning: No warranty is provided with this set  |'
'|of code so use it in your own risk. The author  |'
'|is not responsible for the Damage caused by     |'
'|this code.                                      |'
'|------------------------------------------------|'
'|------------------------------------------------|'
'|Comments:This little piece of code returns      |'
'|a keycode from the input.And it can handle      |'
'|a borderless form,I mean drag.                  |'
'|It has a cool sliding side bar that shows       |'
'|code how to use keyboard key to establish       |'
'|a command. Check it out....                     |'
'|------------------------------------------------|'
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub Form_KeyPress(KeyAscii As Integer)
lblInfo(0).Caption = KeyAscii
lblInfo(1).Caption = Chr$(KeyAscii)
End Sub

Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
mkDrag Me
End Sub
Private Sub cmdButton_Click(Index As Integer)
Select Case Index
Case 0
    Clipboard.SetText lblInfo(0).Caption
Case 1
    If lblInfo(0) = "" Then
        MsgBox "Hey,you didn't press any button yet", vbExclamation, "Null KeyCode !"
    Else
        shCode
        exSize
    End If
Case 2
    Unload Me
Case 3
    comSize
Case 4
    Clipboard.SetText txtCode
Case 5
    txtCode = ""
    Clipboard.Clear
End Select
End Sub
Private Sub piBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mkDrag Me
End Sub

'|-----------------------------------------------|'
'|Code for make the form Dragable                |'
'|(First release it and then tell the system     |'
'|that it is moving)                             |'
'|-----------------------------------------------|'
    
Public Sub mkDrag(ByVal frmObject As Form)
    ReleaseCapture
    SendMessage frmObject.hWnd, &HA1, 2, 0
End Sub
Private Sub shCode()
txtCode = "Private Sub Form_KeyPress(KeyAscii As Integer)"
txtCode = txtCode & vbNewLine & "'First,Set the keypreview true from the form properties then use this code"
txtCode = txtCode & vbNewLine & "If KeyAscii = " & lblInfo(0).Caption & " Then '" & lblInfo(1).Caption
txtCode = txtCode & vbNewLine & "'Your command here"
txtCode = txtCode & vbNewLine & "End If"
txtCode = txtCode & vbNewLine & "End Sub"
End Sub
Private Sub exSize()
On Error Resume Next
Dim i
For i = 3090 To 5880 Step 0.8
Me.Width = i
DoEvents
Next
End Sub
Private Sub comSize()
On Error Resume Next
Dim j
For j = 5800 To 3090 Step -1
Me.Width = j
DoEvents
Next
End Sub
