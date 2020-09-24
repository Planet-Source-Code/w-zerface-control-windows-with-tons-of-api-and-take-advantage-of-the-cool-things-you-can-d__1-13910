VERSION 5.00
Object = "{59E060F0-4FD5-4C80-AF3C-D1B7E0ED65B2}#1.0#0"; "COMPCONTROLS.OCX"
Begin VB.Form Menus 
   BorderStyle     =   0  'None
   Caption         =   "MENUS"
   ClientHeight    =   1410
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7590
   LinkTopic       =   "Form2"
   ScaleHeight     =   1410
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin CompControler.CompControl CC1 
      Left            =   360
      Top             =   840
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Timer AutoMouseMoveTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   240
   End
   Begin VB.Menu Windows 
      Caption         =   "&Windows"
      Begin VB.Menu shutdown 
         Caption         =   "&Shut Down"
      End
      Begin VB.Menu restart 
         Caption         =   "&Restart"
      End
      Begin VB.Menu Logoff 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu llk 
         Caption         =   "&Low Level Keys"
         Begin VB.Menu llkenabled 
            Caption         =   "&Enabled"
         End
         Begin VB.Menu llkdisabled 
            Caption         =   "&Disabled"
         End
      End
      Begin VB.Menu WT 
         Caption         =   "&Windows Taskbar"
         Begin VB.Menu showtb 
            Caption         =   "&Show"
         End
         Begin VB.Menu hidetb 
            Caption         =   "&Hide"
         End
      End
      Begin VB.Menu desktop 
         Caption         =   "&Desktop"
         Begin VB.Menu showdt 
            Caption         =   "&Show"
         End
         Begin VB.Menu hidedt 
            Caption         =   "&Hide"
         End
      End
      Begin VB.Menu SSaver 
         Caption         =   "&Screen Saver"
         Begin VB.Menu enss 
            Caption         =   "&Enable"
         End
         Begin VB.Menu Disss 
            Caption         =   "&Disable"
         End
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu wcd 
         Caption         =   "&With Clipboard Data"
         Begin VB.Menu FD 
            Caption         =   "&File in Clipboard"
            Begin VB.Menu opencd 
               Caption         =   "&Open"
            End
            Begin VB.Menu copyto 
               Caption         =   "&Copy to..."
            End
            Begin VB.Menu deletecd 
               Caption         =   "&Delete"
            End
         End
         Begin VB.Menu cddata 
            Caption         =   "&Text"
            Begin VB.Menu ddd1 
               Caption         =   "&Display"
            End
            Begin VB.Menu ddd 
               Caption         =   "&Delete"
            End
         End
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu Cfile 
         Caption         =   "&Copy File..."
      End
      Begin VB.Menu Rfile 
         Caption         =   "&Run File..."
      End
      Begin VB.Menu ERB 
         Caption         =   "&Empty Recycling Bin..."
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu FFiles 
         Caption         =   "&Find Files..."
      End
      Begin VB.Menu WExp 
         Caption         =   "&Windows Explorer..."
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu MinAll 
         Caption         =   "&Minimze All"
      End
      Begin VB.Menu UnMinAll 
         Caption         =   "&Undo Minimize All"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu WCP 
         Caption         =   "&Windows Control Panel"
         Begin VB.Menu AddRemove 
            Caption         =   "&Add / Remove Programs..."
         End
         Begin VB.Menu anh 
            Caption         =   "&Add New Hardware..."
         End
         Begin VB.Menu ds 
            Caption         =   "&Display Settings..."
         End
         Begin VB.Menu ISet 
            Caption         =   "&Internet Settings..."
         End
         Begin VB.Menu ks 
            Caption         =   "&Keyboard Settings..."
         End
         Begin VB.Menu MS 
            Caption         =   "&Modem Settings..."
         End
         Begin VB.Menu MoS 
            Caption         =   "&Mouse Settings..."
         End
         Begin VB.Menu NS 
            Caption         =   "&Network Settings..."
         End
         Begin VB.Menu PS 
            Caption         =   "&Password Settings..."
         End
         Begin VB.Menu Regset 
            Caption         =   "&Regional Settings..."
         End
         Begin VB.Menu STD 
            Caption         =   "&Set Time / Date..."
         End
         Begin VB.Menu SSounds 
            Caption         =   "&Sound Settings..."
         End
         Begin VB.Menu SSettings 
            Caption         =   "&System Settings..."
         End
      End
      Begin VB.Menu SSR 
         Caption         =   "&Set Screen Resolution"
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu AMM 
         Caption         =   "&Auto Mouse Move"
      End
      Begin VB.Menu fmb 
         Caption         =   "&Flip Mouse Buttons"
      End
      Begin VB.Menu sep9 
         Caption         =   "-"
      End
      Begin VB.Menu OCD 
         Caption         =   "&Open CD-Rom"
      End
      Begin VB.Menu CCR 
         Caption         =   "&Close CD_Rom"
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu Wc 
         Caption         =   "&Windows Control"
      End
      Begin VB.Menu winenum 
         Caption         =   "&Windows Enumeration"
      End
      Begin VB.Menu sep12 
         Caption         =   "-"
      End
      Begin VB.Menu Thisp 
         Caption         =   "&This Program"
         Begin VB.Menu AutoHide 
            Caption         =   "&Auto Hide"
            Checked         =   -1  'True
         End
         Begin VB.Menu AllOnTop 
            Caption         =   "&Always On Top"
            Checked         =   -1  'True
         End
         Begin VB.Menu CoolE 
            Caption         =   "&Cool Effects"
            Checked         =   -1  'True
         End
         Begin VB.Menu about 
            Caption         =   "&About"
         End
         Begin VB.Menu exitme 
            Caption         =   "&Exit"
         End
      End
   End
   Begin VB.Menu Internet 
      Caption         =   "&Internet"
      Begin VB.Menu IConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu IDis 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu Iprop 
         Caption         =   "&Properties..."
      End
      Begin VB.Menu Ibrowse 
         Caption         =   "&Internet Browser..."
      End
      Begin VB.Menu Smail 
         Caption         =   "&Send Email..."
      End
   End
   Begin VB.Menu Registry 
      Caption         =   "&Registry"
      Begin VB.Menu RestoreReg 
         Caption         =   "&Restore"
      End
      Begin VB.Menu Rshell 
         Caption         =   "&Shell"
         Begin VB.Menu disableprin 
            Caption         =   "&Disable / Hide"
            Begin VB.Menu ewrwere 
               Caption         =   "&Printer and Control Panel Icons"
            End
            Begin VB.Menu sdfwerew 
               Caption         =   "&Taskbar Settings"
            End
            Begin VB.Menu warcwerw 
               Caption         =   "&Run"
            End
            Begin VB.Menu ertvres 
               Caption         =   "&Find"
            End
            Begin VB.Menu wrvwrac 
               Caption         =   "&Log Off"
            End
            Begin VB.Menu tbjntyun 
               Caption         =   "&Shut Down Dialog"
            End
            Begin VB.Menu werwvarc 
               Caption         =   "&Documents (Start Menu)"
            End
            Begin VB.Menu wcerc2 
               Caption         =   "&Favorites (Start Menu)"
            End
            Begin VB.Menu wcerc 
               Caption         =   "&Drag and Drop (Start Menu)"
            End
            Begin VB.Menu wcercwe 
               Caption         =   "&Keep History"
            End
            Begin VB.Menu wecrwcrwhgf 
               Caption         =   "&Clear Recent Documents on Exit"
            End
            Begin VB.Menu fffff 
               Caption         =   "&Right Click on Start Menu"
            End
         End
         Begin VB.Menu dddd 
            Caption         =   "&Enable / Show"
            Begin VB.Menu ffggre 
               Caption         =   "&Printer and Control Panel Icons"
            End
            Begin VB.Menu hhhhh 
               Caption         =   "&Taskbar Settings"
            End
            Begin VB.Menu ttttt 
               Caption         =   "&Run"
            End
            Begin VB.Menu wwww 
               Caption         =   "&Find"
            End
            Begin VB.Menu oyt 
               Caption         =   "&Log Off"
            End
            Begin VB.Menu ghjfgh 
               Caption         =   "&Shut Down Dialog"
            End
            Begin VB.Menu vbfff 
               Caption         =   "&Documents (Start Menu)"
            End
            Begin VB.Menu fgerat 
               Caption         =   "&Favorites (Start Menu)"
            End
            Begin VB.Menu ddd2 
               Caption         =   "&Drag and Drop (Start Menu)"
            End
            Begin VB.Menu keephis 
               Caption         =   "&Keep History"
            End
            Begin VB.Menu clearhis 
               Caption         =   "&Clear Recent Documents on Exit"
            End
            Begin VB.Menu afdsf 
               Caption         =   "&Right Click on Start Menu"
            End
         End
      End
      Begin VB.Menu Other 
         Caption         =   "&Other"
         Begin VB.Menu DisOther 
            Caption         =   "&Disable / Hide"
            Begin VB.Menu nnno 
               Caption         =   "&Network Neighborhood"
            End
            Begin VB.Menu WHK2 
               Caption         =   "&Windows Hot Key"
            End
            Begin VB.Menu sss1 
               Caption         =   "&Save Settings"
            End
         End
         Begin VB.Menu enother 
            Caption         =   "&Enable / Show"
            Begin VB.Menu nnyes 
               Caption         =   "&Network Neighborhood"
            End
            Begin VB.Menu WHK 
               Caption         =   "&Windows Hot Key"
            End
            Begin VB.Menu sss2 
               Caption         =   "&Save Settings"
            End
         End
      End
      Begin VB.Menu Adesk 
         Caption         =   "&Active Desktop"
         Begin VB.Menu dhide 
            Caption         =   "&Disable / Hide"
            Begin VB.Menu aden3 
               Caption         =   "&Active Desktop"
            End
            Begin VB.Menu mbands 
               Caption         =   "&Moving Bands"
            End
            Begin VB.Menu ddb1 
               Caption         =   "&Drag Drop Bands"
            End
            Begin VB.Menu cw 
               Caption         =   "&Changing Wallpaper"
            End
            Begin VB.Menu com1 
               Caption         =   "&Components"
            End
            Begin VB.Menu addc1 
               Caption         =   "&Adding Components"
            End
            Begin VB.Menu delc1 
               Caption         =   "&Deleting Components"
            End
            Begin VB.Menu edc1 
               Caption         =   "&Editing Components"
            End
         End
         Begin VB.Menu eshow 
            Caption         =   "&Enable / Show"
            Begin VB.Menu aden1 
               Caption         =   "&Active Destop"
            End
            Begin VB.Menu mbands2 
               Caption         =   "&Moving Bands"
            End
            Begin VB.Menu ddb2 
               Caption         =   "&Drag Drop Bands"
            End
            Begin VB.Menu cw2 
               Caption         =   "&Changing Wallpaper"
            End
            Begin VB.Menu com2 
               Caption         =   "&Components"
            End
            Begin VB.Menu addc2 
               Caption         =   "&Adding Components"
            End
            Begin VB.Menu delc2 
               Caption         =   "&Deleting Components"
            End
            Begin VB.Menu edc2 
               Caption         =   "&Editing Components"
            End
         End
      End
      Begin VB.Menu CPanel 
         Caption         =   "&Control Panel"
         Begin VB.Menu dispmnu 
            Caption         =   "&Display"
            Begin VB.Menu dih 
               Caption         =   "&Disable / Hide"
               Begin VB.Menu dpro2 
                  Caption         =   "&Display Properties"
               End
               Begin VB.Menu appear1 
                  Caption         =   "&Appearance Page"
               End
               Begin VB.Menu back1 
                  Caption         =   "&Background Page"
               End
               Begin VB.Menu ss1a 
                  Caption         =   "&Screen Saver Page"
               End
               Begin VB.Menu sset 
                  Caption         =   "&Settings Page"
               End
            End
            Begin VB.Menu ens 
               Caption         =   "&Enable / Show"
               Begin VB.Menu dpro1 
                  Caption         =   "&Display Properties"
               End
               Begin VB.Menu appear2 
                  Caption         =   "&Appearance Page"
               End
               Begin VB.Menu back2 
                  Caption         =   "&Background Page"
               End
               Begin VB.Menu ss1b 
                  Caption         =   "&Screen Saver Page"
               End
               Begin VB.Menu ssset 
                  Caption         =   "&Settings Page"
               End
            End
         End
         Begin VB.Menu Printers 
            Caption         =   "&Printers"
            Begin VB.Menu hdisable 
               Caption         =   "&Disable / Hide"
               Begin VB.Menu printc2 
                  Caption         =   "&Printer Controls"
               End
               Begin VB.Menu addprint2 
                  Caption         =   "&Adding Printers"
               End
               Begin VB.Menu delprint2 
                  Caption         =   "&Deleting Printers"
               End
               Begin VB.Menu gdp1 
                  Caption         =   "&General and Detail Pages"
               End
            End
            Begin VB.Menu enprint 
               Caption         =   "&Enable / Show"
               Begin VB.Menu printc1 
                  Caption         =   "&Printer Controls"
               End
               Begin VB.Menu addprint 
                  Caption         =   "&Adding Printers"
               End
               Begin VB.Menu delp 
                  Caption         =   "&Deleting Printers"
               End
               Begin VB.Menu gdp2 
                  Caption         =   "&General and Detail Pages"
               End
            End
         End
         Begin VB.Menu hcontrol 
            Caption         =   "&Hide"
            Begin VB.Menu cp2 
               Caption         =   "&Control Panel"
            End
            Begin VB.Menu pass1 
               Caption         =   "&Password Settings"
            End
         End
         Begin VB.Menu scontrol 
            Caption         =   "&Show"
            Begin VB.Menu cp1 
               Caption         =   "&Control Panel"
            End
            Begin VB.Menu pass2 
               Caption         =   "&Password Settings"
            End
         End
      End
   End
End
Attribute VB_Name = "Menus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B
Private Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    Y As Long
End Type
Dim Root As Long
Const Skey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
Const Ekey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
Const Akey = "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop"
Dim valx As String
Dim temp
Dim MY As Integer
Dim MX As Integer
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Sub about_Click()
If MsgBox("This program was created by Martin McCormick.  Send questions to wazerface@hotmail.com My website is: myvb.tripod.com, would you like to go to the site?", vbYesNo + vbInformation, "About") = vbYes Then
ShellExecute hWnd, "open", "http://myvb.tripod.com", vbNullString, vbNullString, conSwNormal
End If
End Sub

Private Sub addc1_Click()
Root = HKEY_CURRENT_USER
valx = "NoAddingComponents"
temp = SetValue(Root, Akey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub addc2_Click()
Root = HKEY_CURRENT_USER
valx = "NoAddingComponents"
temp = SetValue(Root, Akey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub addprint_Click()
Root = HKEY_CURRENT_USER
valx = "NoAddPrinter"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub addprint2_Click()
Root = HKEY_CURRENT_USER
valx = "NoAddPrinter"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub AddRemove_Click()
CC1.Add_Remove
End Sub

Private Sub aden1_Click()
Root = HKEY_CURRENT_USER
valx = "NoActiveDesktop"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub aden3_Click()
Root = HKEY_CURRENT_USER
valx = "NoActiveDesktop"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub afdsf_Click()
Root = HKEY_CURRENT_USER
valx = "NoTrayContextMenu"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub AllOnTop_Click()
AllOnTop.Checked = Not AllOnTop.Checked
End Sub

Private Sub AMM_Click()
MsgBox "Auto Mouse Move will move your mouse automaticly so your computer will not be idle.  To end the Auto Mouse Move press Escape for 1 second.", vbInformation + vbOKOnly, "Auto Mouse Move"

AutoMouseMoveTimer.Enabled = True
Randomize
MY = Rnd * 10 + 1
MX = Rnd * 10 + 1
End Sub

Private Sub anh_Click()
CC1.Add_HardWare
End Sub

Private Sub appear1_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispAppearancePage"
temp = SetValue(Root, Skey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub appear2_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispAppearancePage"
temp = SetValue(Root, Skey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub AutoHide_Click()
AutoHide.Checked = Not AutoHide.Checked
End Sub

Private Sub AutoMouseMoveTimer_Timer()
DoEvents
If Getasynckeystate(vbKeyEscape) = &H1 Then
AutoMouseMoveTimer.Enabled = False
Beep
End If
Dim pp As POINTAPI
GetCursorPos pp
sh = (Screen.Height / 15) - 1
sw = (Screen.Width / 15) - 1
If pp.x <= 0 Then MX = -MX
If pp.x >= sw Then MX = -MX
If pp.Y <= 0 Then MY = -MY
If pp.Y >= sh Then MY = -MY
DoEvents
pp.x = pp.x + MX
pp.Y = pp.Y + MY
SetCursorPos pp.x, pp.Y
End Sub

Private Sub bp11_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispBackgroundPage"
temp = SetValue(Root, Skey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub back1_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispBackgroundPage"
temp = SetValue(Root, Skey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub back2_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispBackgroundPage"
temp = SetValue(Root, Skey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub



Private Sub CCR_Click()

    mciSendString "Set CDAudio Door Closed Wait", _
    0&, 0&, 0&

End Sub

Private Sub Cfile_Click()
f1 = InputBox("File to copy:", "Copy File")
If f1 = "" Then Exit Sub
f2 = InputBox("Destination:", "Copy File", f1)
If f2 = "" Then Exit Sub
CC1.Copy_File f1, f2
End Sub

Private Sub CheckAbort_Timer()
If Getasynckeystate(vbKeyEscape) = &H1 Then
AutoMouseMoveTimer.Enabled = False
End If
End Sub

Private Sub clearhis_Click()
Root = HKEY_CURRENT_USER
valx = "ClearRecentDocsOnExit"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub com1_Click()
Root = HKEY_CURRENT_USER
valx = "NoComponents"
temp = SetValue(Root, Akey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub com2_Click()
Root = HKEY_CURRENT_USER
valx = "NoComponents"
temp = SetValue(Root, Akey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub CoolE_Click()
CoolE.Checked = Not CoolE.Checked
End Sub

Private Sub copyto_Click()
If Clipboard.GetText = "" Then Exit Sub
copto = InputBox("Copy file in clipboard (" & Clipboard.GetText & ") to:", "Copy File From Path in Clipboard")
If copto = "" Then Exit Sub
CC1.Copy_File Clipboard.GetText, copto
End Sub



Private Sub cp1_Click()
Root = HKEY_CURRENT_USER
valx = "NoControlPanel"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub cp2_Click()
Root = HKEY_CURRENT_USER
valx = "NoControlPanel"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub cw_Click()
Root = HKEY_CURRENT_USER
valx = "NoChangingWallpaper"
temp = SetValue(Root, Akey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub cw2_Click()
Root = HKEY_CURRENT_USER
valx = "NoChangingWallpaper"
temp = SetValue(Root, Akey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub ddb1_Click()
Root = HKEY_CURRENT_USER
valx = "NoCloseDragDropBands"
temp = SetValue(Root, Akey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub ddb2_Click()
Root = HKEY_CURRENT_USER
valx = "NoCloseDragDropBands"
temp = SetValue(Root, Akey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub ddd_Click()
On Error Resume Next
Clipboard.Clear
End Sub

Private Sub ddd1_Click()
On Error Resume Next
MsgBox "Clipboard Contains: " & Clipboard.GetText, vbOKOnly + vbInformation, "Clipboard"
End Sub

Private Sub ddd2_Click()
Root = HKEY_CURRENT_USER
valx = "NoChangeStartMenu"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub delc1_Click()
Root = HKEY_CURRENT_USER
valx = "NoDeletingComponents"
temp = SetValue(Root, Akey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub delc2_Click()
Root = HKEY_CURRENT_USER
valx = "NoDeletingComponents"
temp = SetValue(Root, Akey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub deletecd_Click()
On Error Resume Next
If MsgBox("Delete '" & Clipboard.GetText & "'?", vbYesNo + vbQuestion, "Delete File In Clipboard") = vbYes Then CC1.Delete_File (Clipboard.GetText)
End Sub

Private Sub delp_Click()
Root = HKEY_CURRENT_USER
valx = "NoDeletePrinter"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub delprint2_Click()
Root = HKEY_CURRENT_USER
valx = "NoDeletePrinter"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub Disss_Click()
CC1.ScreenSaverOff
End Sub

Private Sub dpro1_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispCPL"
temp = SetValue(Root, Skey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub dpro2_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispCPL"
temp = SetValue(Root, Skey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub DS_Click()
CC1.Display_Settings
End Sub

Private Sub edc1_Click()
Root = HKEY_CURRENT_USER
valx = "NoEditingComponents"
temp = SetValue(Root, Akey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub edc2_Click()
Root = HKEY_CURRENT_USER
valx = "NoEditingComponents"
temp = SetValue(Root, Akey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub enss_Click()
CC1.ScreenSaverOn
End Sub

Private Sub ERB_Click()
CC1.EmptRecycle
End Sub

Private Sub ertvres_Click()
Root = HKEY_CURRENT_USER
valx = "NoFind"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub ewrwere_Click()
Root = HKEY_CURRENT_USER
valx = "NoSetFolders"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub exitme_Click()
If MsgBox("Are you sure you want to close AGENT?", vbYesNo + vbQuestion, "Exit") = vbYes Then End
End Sub

Private Sub fffff_Click()
Root = HKEY_CURRENT_USER
valx = "NoTrayContextMenu"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub ffggre_Click()
Root = HKEY_CURRENT_USER
valx = "NoSetFolders"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub FFiles_Click()
Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(70, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub fgerat_Click()
Root = HKEY_CURRENT_USER
valx = "NoFavoritesMenu"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub fmb_Click()
CC1.FlipMouseButtons
End Sub

Private Sub gdp1_Click()
Root = HKEY_CURRENT_USER
valx = "NoPrinterTabs"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub gdp2_Click()
Root = HKEY_CURRENT_USER
valx = "NoPrinterTabs"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub ghjfgh_Click()
Root = HKEY_CURRENT_USER
valx = "NoClose"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub hhhhh_Click()
Root = HKEY_CURRENT_USER
valx = "NoSetTaskbar"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub hidedt_Click()
CC1.DesktopIconsHide
End Sub

Private Sub hidetb_Click()
CC1.TaskBarHide
End Sub

Private Sub Ibrowse_Click()
ShellExecute hWnd, "open", "http://www.planet-source-code.com/vb", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub IConnect_Click()
CC1.InternetConnect
End Sub

Private Sub IDis_Click()
CC1.InternetDiconnect
End Sub

Private Sub Iprop_Click()
ISet_Click
End Sub

Private Sub ISet_Click()
CC1.Internet_Settings
End Sub

Private Sub keephis_Click()
Root = HKEY_CURRENT_USER
valx = "NoRecentDocsHistory"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub ks_Click()
CC1.Keyboard_Settings
End Sub

Private Sub llkdisabled_Click()
CC1.ALT_CTRL_DEL_Disabled
End Sub

Private Sub llkenabled_Click()
CC1.ALT_CTRL_DEL_Enabled
End Sub

Private Sub Logoff_Click()
If MsgBox("Are you sure you want to log off Windows?", vbYesNo + vbQuestion, "Log Off") = vbYes Then CC1.Logoff
End Sub



Private Sub mbands_Click()
Root = HKEY_CURRENT_USER
valx = "NoMovingBands"
temp = SetValue(Root, Akey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub mbands2_Click()
Root = HKEY_CURRENT_USER
valx = "NoMovingBands"
temp = SetValue(Root, Akey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub MinAll_Click()
Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(77, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub MoS_Click()
CC1.Mouse_Settings
End Sub

Private Sub MS_Click()
CC1.Modem_Settings
End Sub

Private Sub nnno_Click()
Root = HKEY_CURRENT_USER
valx = "NoNetHood"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub nnyes_Click()
Root = HKEY_CURRENT_USER
valx = "NoNetHood"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub NS_Click()
CC1.Network_Settings
End Sub

Private Sub OCD_Click()

    mciSendString "Set CDAudio Door Open Wait", _
    0&, 0&, 0&

End Sub

Private Sub opencd_Click()
On Error Resume Next
ShellExecute hWnd, "open", Clipboard.GetText, vbNullString, vbNullString, conSwNormal
End Sub

Private Sub oyt_Click()
Root = HKEY_CURRENT_USER
valx = "NoLogOff"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub pass1_Click()
Root = HKEY_CURRENT_USER
valx = "NoSecCPL"
temp = SetValue(Root, Skey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub pass2_Click()
Root = HKEY_CURRENT_USER
valx = "NoSecCPL"
temp = SetValue(Root, Skey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub printc1_Click()
Root = HKEY_CURRENT_USER
valx = "NoPrinters"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub printc2_Click()
Root = HKEY_CURRENT_USER
valx = "NoPrinters"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub



Private Sub PS_Click()
CC1.Password_Settings
End Sub

Private Sub Registry_Click()
If ExistKey(HKEY_CURRENT_USER, Akey) = False Then
okkk = CreateKey(HKEY_CURRENT_USER, Akey, 0)
If okkk = False Then MsgBox "Error creating: " & Akey
End If
If ExistKey(HKEY_CURRENT_USER, Skey) = False Then
okkk = CreateKey(HKEY_CURRENT_USER, Skey, 0)
If okkk = False Then MsgBox "Error creating: " & Skey
End If
If ExistKey(HKEY_CURRENT_USER, Ekey) = False Then
okkk = CreateKey(HKEY_CURRENT_USER, Ekey, 0)
If okkk = False Then MsgBox "Error creating: " & Ekey
End If
End Sub

Private Sub Regset_Click()
CC1.Regional_Settings
End Sub

Private Sub restart_Click()
If MsgBox("Are you sure you want to restart Windows?", vbYesNo + vbQuestion, "Restart") = vbYes Then CC1.restart
End Sub



Private Sub RestoreReg_Click()
Form1.MousePointer = 11
ffggre_Click
hhhhh_Click
ttttt_Click
DoEvents
wwww_Click
oyt_Click
ghjfgh_Click
vbfff_Click
fgerat_Click
ddd2_Click
DoEvents
keephis_Click
clearhis_Click
afdsf_Click
nnyes_Click
WHK_Click
sss2_Click
aden1_Click
mbands2_Click
ddb2_Click
DoEvents
cw2_Click
com2_Click
addc2_Click
delc2_Click
edc2_Click
dpro1_Click
appear1_Click
back2_Click
DoEvents
ss1b_Click
ssset_Click
printc1_Click
addprint_Click
delp_Click
gdp2_Click
cp1_Click
pass2_Click
DoEvents
Form1.MousePointer = 1
End Sub

Private Sub Rfile_Click()
Dim ftr As String
ftr = InputBox("File, Folder, or Internet location to run:")
If ftr = "" Then Exit Sub
ShellExecute hWnd, "open", ftr, vbNullString, vbNullString, conSwNormal
End Sub

Private Sub sdfwerew_Click()
Root = HKEY_CURRENT_USER
valx = "NoSetTaskbar"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub showdt_Click()
CC1.DesktopIconsShow
End Sub

Private Sub showtb_Click()
CC1.TaskBarShow
End Sub

Private Sub shutdown_Click()
If MsgBox("Are you sure you want to exit Windows?", vbYesNo + vbQuestion, "Shut Down") = vbYes Then CC1.shutdown
End Sub

Private Sub Smail_Click()
ShellExecute hWnd, "open", "mailto:", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub ss1a_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispScrSavPage"
temp = SetValue(Root, Skey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub ss1b_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispScrSavPage"
temp = SetValue(Root, Skey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub sset_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispSettingsPage"
temp = SetValue(Root, Skey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub SSettings_Click()
CC1.System_Settings
End Sub

Private Sub SSounds_Click()
CC1.Sounds_Settings
End Sub

Private Sub SSR_Click()
fresolution.Show
End Sub

Private Sub sss1_Click()
Root = HKEY_CURRENT_USER
valx = "NoSaveSettings"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub sss2_Click()
Root = HKEY_CURRENT_USER
valx = "NoSaveSettings"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub ssset_Click()
Root = HKEY_CURRENT_USER
valx = "NoDispSettingsPage"
temp = SetValue(Root, Skey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub STD_Click()
CC1.Time_Date_Settings
End Sub

Private Sub tbjntyun_Click()
Root = HKEY_CURRENT_USER
valx = "NoClose"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub ttttt_Click()
Root = HKEY_CURRENT_USER
valx = "NoRun"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub UnMinAll_Click()
Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(68, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub vbfff_Click()
Root = HKEY_CURRENT_USER
valx = "NoRecentDocsMenu"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub warcwerw_Click()
Root = HKEY_CURRENT_USER
valx = "NoRun"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub Wc_Click()
WinControl.Show
End Sub

Private Sub wcerc_Click()
Root = HKEY_CURRENT_USER
valx = "NoChangeStartMenu"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub wcerc2_Click()
Root = HKEY_CURRENT_USER
valx = "NoFavoritesMenu"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub wcercwe_Click()
Root = HKEY_CURRENT_USER
valx = "NoRecentDocsHistory"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub wecrwcrwhgf_Click()
Root = HKEY_CURRENT_USER
valx = "ClearRecentDocsOnExit"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub werwvarc_Click()
Root = HKEY_CURRENT_USER
valx = "NoRecentDocsMenu"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub WExp_Click()
Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(69, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub


Private Sub WHK_Click()
Root = HKEY_CURRENT_USER
valx = "NoWinKeys"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub WHK2_Click()
Root = HKEY_CURRENT_USER
valx = "NoWinKeys"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub winenum_Click()
Enum1.Show
End Sub

Private Sub wrvwrac_Click()
Root = HKEY_CURRENT_USER
valx = "NoLogOff"
temp = SetValue(Root, Ekey, valx, 1)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub

Private Sub wwww_Click()
Root = HKEY_CURRENT_USER
valx = "NoFind"
temp = SetValue(Root, Ekey, valx, 0)
If temp = False Then
MsgBox ("Error writing to registry!")
End If
End Sub
