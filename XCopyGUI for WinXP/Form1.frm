VERSION 5.00
Begin VB.Form frmXCopyGUI 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5925
   ClientLeft      =   2625
   ClientTop       =   1170
   ClientWidth     =   5730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   " Attributes "
      Height          =   1515
      Left            =   180
      TabIndex        =   5
      Top             =   960
      Width           =   5415
      Begin VB.CheckBox chkArchReset 
         Caption         =   "Copy only files with the Archive attribute set, then turn it off"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   4995
      End
      Begin VB.CheckBox chkArch 
         Caption         =   "Copy only files with the Archive attribute set"
         Height          =   200
         Left            =   120
         TabIndex        =   8
         Top             =   900
         Width           =   4035
      End
      Begin VB.CheckBox chkHiddenSystem 
         Caption         =   "Copy Hidden and System files also"
         Height          =   200
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   3135
      End
      Begin VB.CheckBox chkOverwriteReadOnly 
         Caption         =   "Overwrite Read-only files"
         Height          =   200
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2235
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Behaviour "
      Height          =   975
      Left            =   180
      TabIndex        =   10
      Top             =   2520
      Width           =   5415
      Begin VB.CheckBox chkReplace 
         Caption         =   "Replace (only copy files that already exist in destination)"
         Height          =   200
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   4275
      End
      Begin VB.CheckBox chkUpdate 
         Caption         =   "Update (only copy newer or non-existing files)"
         Height          =   200
         Left            =   120
         TabIndex        =   12
         Top             =   660
         Width           =   4155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Options "
      Height          =   1335
      Left            =   180
      TabIndex        =   16
      Top             =   4500
      Width           =   4035
      Begin VB.CheckBox chkOverwrite 
         Caption         =   "Prompt before overwriting existing files"
         Height          =   200
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox chkBypassErrors 
         Caption         =   "Continue copying even if errors occur"
         Height          =   200
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   3075
      End
      Begin VB.CheckBox chkTest 
         Caption         =   "Test only (Display files that would be copied)"
         Height          =   200
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   3555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Folders "
      Height          =   915
      Left            =   180
      TabIndex        =   13
      Top             =   3540
      Width           =   5415
      Begin VB.CheckBox chkEmpty 
         Caption         =   "Copy empty folders"
         Height          =   200
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   3075
      End
      Begin VB.CheckBox chkDirCreateOnly 
         Caption         =   "Create folder structure, but don't copy files"
         Height          =   200
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   540
      Width           =   315
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5340
      Width           =   1150
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   390
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4800
      Width           =   1150
   End
   Begin VB.TextBox txtDestination 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   540
      Width           =   4035
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "*.*"
      Top             =   120
      Width           =   4515
   End
   Begin VB.Label lblDestination 
      AutoSize        =   -1  'True
      Caption         =   "Destination"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   795
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      Caption         =   "File Name"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   705
   End
End
Attribute VB_Name = "frmXCopyGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Public Sub CButtons(frmX As Form, Optional Identifier As String)
' Button.Style must be GRAPHICAL

Dim ctl As Control

For Each ctl In frmX      'loop trough all the controls on the form
    
    '3 Methods of doing it
    'If LCase(Left(Control.Name, Len(Identifier))) = LCase(Identifier) Then
    'If TypeName(Control) = "CommandButton" Then
    If TypeOf ctl Is CommandButton Then
                SendMessage ctl.hWnd, &HF4&, &H0&, 0&
    End If

Next ctl

End Sub


Function UnQuote(ByVal Text As String) As String
Const Quote = """"
Dim sTemp As String

sTemp = Text

If Left$(Text, 1) = Quote And Right$(Text, 1) = Quote Then
'    sTemp = Right$(sTemp, Len(sTemp) - 1)
'    sTemp = Left$(sTemp, Len(sTemp) - 1)
    sTemp = Mid$(sTemp, 2, Len(sTemp) - 2)
End If


UnQuote = sTemp

End Function

Private Sub cmdOk_Click()
Dim Source As String
Dim Destination As String
Dim cmd As String
Const q = """"

If txtFileName.Text = "" Or txtDestination.Text = "" Then
    Beep
    Exit Sub
End If

Source = q & UnQuote(Command$) & "\" & txtFileName.Text & q
Source = Replace(Source, "\\", "\") 'in case!

Destination = q & txtDestination.Text & q

cmd = "xcopy " & Source & " " & Destination & " /S /I"

If chkEmpty.Value = vbChecked Then
    cmd = cmd & " /E"
End If
If chkBypassErrors.Value = vbChecked Then
    cmd = cmd & " /C"
End If
If chkHiddenSystem.Value = vbChecked Then
    cmd = cmd & " /H"
End If
If chkDirCreateOnly.Value = vbChecked Then
    cmd = cmd & " /T"
End If
If chkOverwrite.Value = vbChecked Then
    cmd = cmd & " /-Y"
Else
    cmd = cmd & " /Y"
End If
If chkOverwriteReadOnly.Value = vbChecked Then
    cmd = cmd & " /R"
End If
If chkUpdate.Value = vbChecked Then
    cmd = cmd & " /D"
End If
If chkReplace.Value = vbChecked Then
    cmd = cmd & " /U"
End If

If chkTest.Value = vbChecked Then
    cmd = cmd & " /L /Y"
End If
If chkArch.Value = vbChecked Then
    cmd = cmd & " /A"
End If
If chkArchReset.Value = vbChecked Then
    cmd = cmd & " /M"
End If


Shell cmd, vbNormalFocus

End Sub
Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()
Dim folder As String
folder = BrowseForFolder(Me, "Select Output Folder", CStr(Me.txtDestination.Text))
If Len(folder) <> 0 Then
   Me.txtDestination.Text = folder
End If

End Sub

Private Sub Form_Activate()

Me.txtFileName.SelStart = 0
Me.txtFileName.SelLength = Len(Me.txtFileName.Text)
Me.txtFileName.SetFocus



End Sub

Private Sub Form_Initialize()
Dim x As Long
x = InitCommonControls

End Sub


Private Sub Form_Load()

If Command$ = "" Then End

CButtons Me
Caption = "Copy from:  " & Command$

'''''''''''''''''''''''''''''''''''''''''''''
Dim State As String

State = GetSetting("XCopyGUI", "Options", "State", "00001000000")
txtDestination.Text = GetSetting("XCopyGUI", "Options", "LastDestination", "")

If Mid$(State, 1, 1) = "1" Then chkEmpty.Value = vbChecked Else chkEmpty.Value = vbUnchecked
If Mid$(State, 2, 1) = "1" Then chkBypassErrors.Value = vbChecked Else chkBypassErrors.Value = vbUnchecked
If Mid$(State, 3, 1) = "1" Then chkHiddenSystem.Value = vbChecked Else chkHiddenSystem.Value = vbUnchecked
If Mid$(State, 4, 1) = "1" Then chkDirCreateOnly.Value = vbChecked Else chkDirCreateOnly.Value = vbUnchecked
If Mid$(State, 5, 1) = "1" Then chkOverwrite.Value = vbChecked Else chkOverwrite.Value = vbUnchecked
If Mid$(State, 6, 1) = "1" Then chkOverwriteReadOnly.Value = vbChecked Else chkOverwriteReadOnly.Value = vbUnchecked
If Mid$(State, 7, 1) = "1" Then chkUpdate.Value = vbChecked Else chkUpdate.Value = vbUnchecked
If Mid$(State, 8, 1) = "1" Then chkReplace.Value = vbChecked Else chkReplace.Value = vbUnchecked
If Mid$(State, 9, 1) = "1" Then chkTest.Value = vbChecked Else chkTest.Value = vbUnchecked

If Mid$(State, 10, 1) = "1" Then chkArch.Value = vbChecked Else chkArch.Value = vbUnchecked
If Mid$(State, 11, 1) = "1" Then chkArchReset.Value = vbChecked Else chkArchReset.Value = vbUnchecked

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim State As String

If chkEmpty.Value = vbChecked Then State = State & "1" Else State = State & "0"
If chkBypassErrors.Value = vbChecked Then State = State & "1" Else State = State & "0"
If chkHiddenSystem.Value = vbChecked Then State = State & "1" Else State = State & "0"
If chkDirCreateOnly.Value = vbChecked Then State = State & "1" Else State = State & "0"
If chkOverwrite.Value = vbChecked Then State = State & "1" Else State = State & "0"
If chkOverwriteReadOnly.Value = vbChecked Then State = State & "1" Else State = State & "0"
If chkUpdate.Value = vbChecked Then State = State & "1" Else State = State & "0"
If chkReplace.Value = vbChecked Then State = State & "1" Else State = State & "0"
If chkTest.Value = vbChecked Then State = State & "1" Else State = State & "0"

If chkArch.Value = vbChecked Then State = State & "1" Else State = State & "0"
If chkArchReset.Value = vbChecked Then State = State & "1" Else State = State & "0"

SaveSetting "XCopyGUI", "Options", "State", State

SaveSetting "XCopyGUI", "Options", "LastDestination", txtDestination.Text

End Sub

Private Sub txtDestination_GotFocus()

Me.txtDestination.SelStart = 0
Me.txtDestination.SelLength = Len(Me.txtDestination.Text)

End Sub


