VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Windows AGENT"
   ClientHeight    =   2100
   ClientLeft      =   -45
   ClientTop       =   -330
   ClientWidth     =   12000
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&GO"
      Height          =   285
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   600
      Top             =   840
   End
   Begin VB.Timer TPause 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   840
   End
   Begin VB.Label LbLInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Index           =   5
      Left            =   10560
      TabIndex        =   17
      Top             =   300
      Width           =   1455
   End
   Begin VB.Label LbLInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   4
      Left            =   11040
      TabIndex        =   16
      Top             =   150
      Width           =   975
   End
   Begin VB.Label LbLInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   3
      Left            =   11040
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
   Begin VB.Label LbLInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Index           =   2
      Left            =   8280
      TabIndex        =   14
      Top             =   300
      Width           =   1335
   End
   Begin VB.Label LbLInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   13
      Top             =   145
      Width           =   1215
   End
   Begin VB.Label LbLInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   12
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Virtual Memory:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   9720
      TabIndex        =   11
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Page File:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   9720
      TabIndex        =   10
      Top             =   300
      Width           =   2220
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Page File:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   9720
      TabIndex        =   9
      Top             =   150
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Virtual Memory:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   6960
      TabIndex        =   8
      Top             =   300
      Width           =   2265
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Physical Memory:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   6960
      TabIndex        =   7
      Top             =   150
      Width           =   2325
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Physical Memory:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   6960
      TabIndex        =   6
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registry"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Internet"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Windows"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1185
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   11
      Left            =   2760
      Picture         =   "Form1.frx":0442
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   10
      Left            =   2520
      Picture         =   "Form1.frx":14CA4
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   9
      Left            =   2280
      Picture         =   "Form1.frx":29506
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   8
      Left            =   2040
      Picture         =   "Form1.frx":3DD68
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   7
      Left            =   1800
      Picture         =   "Form1.frx":525CA
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   6
      Left            =   1560
      Picture         =   "Form1.frx":66E2C
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   5
      Left            =   1320
      Picture         =   "Form1.frx":7B68E
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   4
      Left            =   1080
      Picture         =   "Form1.frx":8FEF0
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   3
      Left            =   840
      Picture         =   "Form1.frx":A4752
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   2
      Left            =   600
      Picture         =   "Form1.frx":B8FB4
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   1
      Left            =   360
      Picture         =   "Form1.frx":CD816
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image BG 
      Height          =   525
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":E2078
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   15
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FHeight
Dim MenuOpen As Boolean
Dim amR
Dim amG
Dim amB
Dim MOver As New ClsMouseover
Dim alert As Boolean
Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const SR = 0
Private Const GDI = 1
Private Const USR = 2

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Type MEMORYSTATUS
    dwLength        As Long
    dwMemoryLoad    As Long
    dwTotalPhys     As Long
    dwAvailPhys     As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual  As Long
    dwAvailVirtual  As Long
End Type


Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function pBGetFreeSystemResources Lib "rsrc32.dll" Alias "_MyGetFreeSystemResources32@4" (ByVal iResType As Integer) As Integer
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private mIsWin32 As Boolean

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Private Sub Command1_Click()
Text1_KeyPress (13)
End Sub

Private Sub Form_Load()
If App.PrevInstance Then alert = True
Menus.CC1.FormOnTop Me, 0, 0, Me.Width * 16, Me.Height * 16
'Variables:
FHeight = 500
'
Me.Top = 0: Me.Left = 0: Me.Width = Screen.Width: Me.Height = FHeight
MOver.SetBorderStyle (Me.BorderStyle)
Label1.Width = Form1.Width
Label1.Height = Form1.Height
Form1.Picture = BG(0).Picture
Label2.Top = (FHeight / 2) - (Label2.Height / 2)
Label3.Top = (FHeight / 2) - (Label3.Height / 2)
Label4.Top = (FHeight / 2) - (Label4.Height / 2)
Timer1_Timer

    Dim OSInfo As OSVERSIONINFO

    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    
    Call GetVersionEx(OSInfo)
    mIsWin32 = (OSInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS)

End Sub

Private Sub Form_Terminate()
DestroyAppBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("A program atempted to close me! Would you like to deny the request?", vbYesNo + vbExclamation, "Alert!") = vbYes Then
Cancel = 1
On Error Resume Next
worked = ShellExecute(hWnd, "open", App.Path & "\" & App.EXEName, vbNullString, vbNullString, conSwNormal)
DoEvents
If worked <> False Or worked <> 0 Then End
Else
DestroyAppBar
Cancel = 0
Unload Me
End
End If

End Sub

Private Sub Label2_Click()
MenuOpen = True
PopupMenu Menus.Windows, 0, Label2.Left, Label2.Height + Label2.Top
MenuOpen = False
Timer1_Timer
End Sub

Private Sub Label3_Click()
MenuOpen = True
PopupMenu Menus.Internet, 0, Label3.Left, Label3.Height + Label3.Top
MenuOpen = False
Timer1_Timer
End Sub

Private Sub Label4_Click()
MenuOpen = True
PopupMenu Menus.Registry, 0, Label4.Left, Label4.Height + Label4.Top
MenuOpen = False
Timer1_Timer
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
ShellExecute hWnd, "open", Text1.Text, vbNullString, vbNullString, conSwNormal
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If MenuOpen = False Then
If Menus.AllOnTop.Checked = True Then
Menus.CC1.FormOnTop Me, 0, 0, Me.Width / 15, Me.Height / 15
Me.Top = 0: Me.Left = 0: Me.Width = Screen.Width
Else
 Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS1)
End If
DoEvents
End If
If MOver.IsMouseOver(Me, Label1) Then
TPause.Enabled = True
Else
If MenuOpen = False And Menus.AutoHide.Checked = True Then
DestroyAppBar
tsec = Second(Time)
Do Until Form1.Height <= 0
tsec2 = Second(Time)
If tsec2 = tsec + 2 Then Exit Do
Form1.Height = Form1.Height - 1
DoEvents
Loop
Form1.Height = 10
End If
End If

If Menus.CoolE.Checked = True Then

If Form1.Height > FHeight / 2 Then
Static v
Static direct As Boolean
If MenuOpen = False Then
If direct = False Then
v = v + 1
If v = 11 Then direct = True
Else
v = v - 1
If v = 0 Then direct = False
End If

Form1.Picture = BG(v).Picture

glevel = (21.3 * v)
Text1.BackColor = RGB(glevel, glevel, glevel)
Text1.ForeColor = RGB(256 - glevel, 256 - glevel, 256 - glevel)
Randomize
col1 = RGB(glevel, glevel, glevel)
End If
Else
col1 = vbWhite

End If

Else
col1 = vbWhite
Text1.BackColor = vbBlack
Text1.ForeColor = vbWhite
Form1.Picture = Nothing






End If



DoEvents


If MenuOpen = False Then
If MOver.IsMouseOver(Me, Label2) Then
Label2.ForeColor = vbRed
Label2.FontBold = True
Else
Label2.ForeColor = col1
Label2.FontBold = False
End If
If MOver.IsMouseOver(Me, Label3) Then
Label3.ForeColor = vbRed
Label3.FontBold = True
Else
Label3.ForeColor = col1
Label3.FontBold = False
End If
If MOver.IsMouseOver(Me, Label4) Then
Label4.ForeColor = vbRed
Label4.FontBold = True
Else
Label4.ForeColor = col1
Label4.FontBold = False
End If
Else
End If
If alert = True Then App.Title = "AGENT" & Int(Rnd * 9999999)
If Menus.CoolE.Checked = True Then
Command1.BackColor = Text1.BackColor
Command1.MaskColor = Text1.ForeColor
Else
Command1.BackColor = vbWhite
Command1.MaskColor = vbWhite
End If
DoEvents
If Form1.Height > 250 Then
UPDATEInfo
End If
If Menus.AutoHide.Checked = False Then

CreateAppBar Me, jtop

End If
End Sub

Private Sub Timer2_Timer()
If Getasynckeystate(vbKeyShift) = &H1 And Getasynckeystate(vbKeyF4) = &H1 Then
Beep
WinControl.Show
MsgBox "Shift + F4 was pressed!", vbExclamation + vbOKOnly, "Shift + F4"
End If
End Sub

Private Sub TPause_Timer()
tsec = Second(Time)
TPause.Enabled = False
If MOver.IsMouseOver(Me, Label1) Then
Do Until Form1.Height >= FHeight
tsec2 = Second(Time)
If tsec2 = tsec + 2 Then Exit Do
Form1.Height = Form1.Height + 5
DoEvents
Loop
Form1.Height = FHeight
End If
End Sub
Private Sub UPDATEInfo()
Dim MS As MEMORYSTATUS

    On Local Error Resume Next


MS.dwLength = Len(MS)
Call GlobalMemoryStatus(MS)
        With MS
            LbLInfo(0) = Format$(.dwTotalPhys / 1024, "#,###") & " Kb"
            LbLInfo(1) = Format$(.dwAvailPhys / 1024, "#,###") & " Kb"
            LbLInfo(2) = Format$(.dwTotalVirtual / 1024, "#,###") & " Kb"
            LbLInfo(3) = Format$(.dwAvailVirtual / 1024, "#,###") & " Kb"
            LbLInfo(4) = Format$(.dwTotalPageFile / 1024, "#,###") & " Kb"
            LbLInfo(5) = Format$(.dwAvailPageFile / 1024, "#,###") & " Kb"
        End With


End Sub
