VERSION 5.00
Begin VB.Form Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackd Mage"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   3690
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHur 
      Caption         =   "Auto Haste"
      Height          =   195
      Left            =   2160
      TabIndex        =   46
      Top             =   3240
      Width           =   1155
   End
   Begin VB.CheckBox chkTarget 
      Caption         =   "Attack Target"
      Height          =   195
      Left            =   180
      TabIndex        =   45
      Top             =   3480
      Value           =   2  'Grayed
      Width           =   1275
   End
   Begin VB.CheckBox chkReuse 
      Caption         =   "Click Reuse"
      Height          =   195
      Left            =   2160
      TabIndex        =   44
      Top             =   3480
      Value           =   2  'Grayed
      Width           =   1155
   End
   Begin VB.CheckBox chkMW 
      Caption         =   "MW Timer"
      Height          =   195
      Left            =   180
      TabIndex        =   43
      Top             =   3960
      Value           =   2  'Grayed
      Width           =   1155
   End
   Begin VB.CheckBox chkGold 
      Caption         =   "Chage Gold"
      Height          =   195
      Left            =   180
      TabIndex        =   42
      Top             =   3720
      Value           =   2  'Grayed
      Width           =   1155
   End
   Begin VB.CommandButton cmdSpyDOWN 
      Caption         =   "\/"
      Height          =   235
      Left            =   3240
      TabIndex        =   41
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSpyUP 
      Caption         =   "/\"
      Height          =   235
      Left            =   3240
      TabIndex        =   40
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSpy 
      Caption         =   "Level Spy"
      Height          =   235
      Left            =   3240
      TabIndex        =   39
      ToolTipText     =   "Return Camera"
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkUtamo 
      Caption         =   "Auto Utamo"
      Height          =   195
      Left            =   2160
      TabIndex        =   38
      Top             =   3000
      Width           =   1155
   End
   Begin VB.CheckBox chkSpeed 
      Caption         =   "Speed Hack"
      Height          =   195
      Left            =   180
      TabIndex        =   37
      Top             =   3240
      Width           =   1275
   End
   Begin VB.CommandButton cmdQuickSave 
      Caption         =   "Quick Save"
      Height          =   255
      Left            =   2400
      TabIndex        =   36
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Tibia"
      Height          =   255
      Left            =   2400
      TabIndex        =   34
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtFlash 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1380
      MaxLength       =   6
      TabIndex        =   33
      Text            =   "0"
      Top             =   2940
      Width           =   615
   End
   Begin VB.CheckBox chkFlash 
      Caption         =   "Alert if HP <"
      Height          =   195
      Left            =   180
      TabIndex        =   32
      ToolTipText     =   "Blink Tibia window"
      Top             =   3000
      Width           =   1155
   End
   Begin VB.Timer HealingPot 
      Interval        =   100
      Left            =   120
      Top             =   1080
   End
   Begin VB.ComboBox cmbEat 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2460
      Width           =   975
   End
   Begin VB.CheckBox chkLight 
      Caption         =   "Light Hack"
      Height          =   195
      Left            =   180
      TabIndex        =   30
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CheckBox chkEat 
      Caption         =   "AutoEat"
      Height          =   195
      Left            =   180
      TabIndex        =   29
      ToolTipText     =   "For newer Tibia versions"
      Top             =   2520
      Width           =   915
   End
   Begin VB.ComboBox cmbHotkey2 
      Height          =   315
      Left            =   2460
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   1320
      Width           =   1035
   End
   Begin VB.ComboBox cmbHotkey1 
      Height          =   315
      Left            =   2460
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   960
      Width           =   1035
   End
   Begin VB.TextBox txtHealPot 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1380
      TabIndex        =   24
      Text            =   "0"
      Top             =   1320
      Width           =   675
   End
   Begin VB.TextBox txtManaPot 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1380
      TabIndex        =   23
      Text            =   "0"
      Top             =   960
      Width           =   675
   End
   Begin VB.TextBox txt_mousey 
      Height          =   375
      Left            =   7140
      TabIndex        =   19
      Text            =   "txt_mousey"
      Top             =   540
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txt_mousex 
      Height          =   375
      Left            =   7140
      TabIndex        =   18
      Text            =   "txt_mousex"
      Top             =   120
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtManaTrain 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2940
      TabIndex        =   17
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txtTrainSpell 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1380
      TabIndex        =   15
      Text            =   "utevo lux"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox chkTrain 
      Caption         =   "Mana Train"
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   2040
      Width           =   1155
   End
   Begin VB.Timer tmrMain 
      Interval        =   100
      Left            =   960
      Top             =   2280
   End
   Begin VB.CheckBox chkIdle 
      Caption         =   "Anti-Idle"
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   2280
      Value           =   2  'Grayed
      Width           =   915
   End
   Begin VB.Timer Healing 
      Interval        =   100
      Left            =   120
      Top             =   240
   End
   Begin VB.TextBox txtMPHi 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2940
      MaxLength       =   6
      TabIndex        =   12
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtMPLow 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2940
      MaxLength       =   6
      TabIndex        =   11
      Text            =   "0"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtHi 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1980
      MaxLength       =   6
      TabIndex        =   8
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtLow 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1980
      MaxLength       =   6
      TabIndex        =   7
      Text            =   "0"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtSpellHi 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   660
      TabIndex        =   4
      Text            =   "exura vita"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtSpellLow 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   660
      TabIndex        =   3
      Text            =   "exura gran"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblCharacter 
      Caption         =   " Old Tibia Page - 10jun2021"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   10
      X1              =   60
      X2              =   60
      Y1              =   1920
      Y2              =   4320
   End
   Begin VB.Label Label1 
      Caption         =   "HK"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   26
      Top             =   1380
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "HK"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   25
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label lbl11 
      Caption         =   "Heal Potion HP"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   1380
      Width           =   1155
   End
   Begin VB.Label lbl11 
      Caption         =   "Mana Potion MP"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   9
      X1              =   720
      X2              =   3600
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   8
      X1              =   3600
      X2              =   3600
      Y1              =   1920
      Y2              =   4320
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   7
      X1              =   60
      X2              =   3600
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   6
      X1              =   7080
      X2              =   7080
      Y1              =   3480
      Y2              =   5400
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   5
      X1              =   60
      X2              =   180
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lbl11 
      Caption         =   "Extras"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label lbl10 
      Caption         =   "MP"
      Height          =   195
      Left            =   2640
      TabIndex        =   16
      Top             =   2100
      Width           =   255
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   60
      X2              =   180
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   840
      X2              =   3600
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Label lbl7 
      Caption         =   "MP"
      Height          =   195
      Left            =   2640
      TabIndex        =   10
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lbl6 
      Caption         =   "MP"
      Height          =   195
      Left            =   2640
      TabIndex        =   9
      Top             =   300
      Width           =   255
   End
   Begin VB.Label lbl5 
      Caption         =   "HP"
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lbl4 
      Caption         =   "HP"
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   300
      Width           =   255
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   60
      X2              =   3600
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lbl3 
      Caption         =   "Heavy"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   495
   End
   Begin VB.Label lbl2 
      Caption         =   "Light"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   375
   End
   Begin VB.Label lbl1 
      Caption         =   "Healing"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   675
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   60
      X2              =   60
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Menu mDeveloper 
      Caption         =   "Developer"
   End
   Begin VB.Menu mCaveBot 
      Caption         =   "CaveBot"
   End
   Begin VB.Menu mAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkGold_Click()

If chkGold.Value = 0 Then
    chkGold.Value = 2
Else
    chkGold.Value = 2
End If

End Sub

Private Sub chkIdle_Click()

If chkIdle.Value = 0 Then
    chkIdle.Value = 2
Else
    chkIdle.Value = 2
End If

End Sub

Private Sub chkMW_Click()

If chkMW.Value = 0 Then
    chkMW.Value = 2
Else
    chkMW.Value = 2
End If

End Sub

Private Sub chkReuse_Click()

If chkReuse.Value = 0 Then
    chkReuse.Value = 2
Else
    chkReuse.Value = 2
End If

End Sub

Private Sub chkSpeed_Click()

If chkSpeed.Value = 0 Then
'UpdateSpeedBase
Else
MySpeedBase = MySpeed
End If

End Sub

Private Sub chkTarget_Click()

If chkTarget.Value = 0 Then
    chkTarget.Value = 2
Else
    chkTarget.Value = 2
End If

End Sub

Private Sub cmdBot_Click()
  frmCavebot.WindowState = vbNormal
  frmCavebot.Show
  frmCavebot.SetFocus
  SetWindowPos frmCavebot.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

'TibiaWindow = getSelectedPID 'pega a janela do char selecionado
'ShowCurrentName2 TibiaWindow 'pega o nome do char/ coloca no title
'InitializeCamps 'retorna os campo default
'PreloadChar 'da load nas conf do char

Private Sub cmdQuickSave_Click()
Dim strPath As String
Dim strRes As String
Dim strFPath As String

'If charName <> "" Then
If TibiaWindow <> 0 Then

SaveSettings strPath

    strPath = App.Path & "\Save"
    If Right$(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If

    'strFPath = strPath & charName & ".ini"
    strFPath = strPath & "default.ini"
    strRes = SaveSettings(strFPath)
    If strRes <> "" Then
    'nothing
    End If

Else

Menu.lblCharacter.Caption = " SELECT TIBIA CLIENT !"
Menu.lblCharacter.ForeColor = &HFF&
    
End If

End Sub

Private Sub cmdSelect_Click()

ClientChooser

If TibiaWindow <> 0 Then
    Menu.lblCharacter.Caption = " Old Tibia - " & TibiaWindow
    Menu.lblCharacter.ForeColor = &HFF&
    PreloadChar 'load configs in default.ini for now
Else
    Menu.lblCharacter.Caption = " -TIBIA NOT DETECTED- "
    Menu.lblCharacter.ForeColor = &HFF&
End If

End Sub

Private Sub cmdSpy_Click()

UpdateSpy

End Sub

Private Sub cmdSpyDOWN_Click()

UpdateSpyDOWN

End Sub

Private Sub cmdSpyUP_Click()

UpdateSpyUP

End Sub

Private Sub Form_Load()

'Versions ' lista de vers�es
InitializeCamps 'campos default
InitializeCampsADR
Load frmDebug
Load frmCavebot
PreloadDefault

End Sub

Private Sub Healing_Timer()

If MyHP > 0 Then
    HealingSub ' fun�ao de healing
End If

End Sub

Private Sub HealingPot_Timer()

If MyHP > 0 Then
    HealingPotSub ' fun�ao de heal pot
End If

End Sub

Private Sub mAbout_Click()

 MsgBox ("This Bot is totally Free and Open Source" & vbCrLf & _
 "" & vbCrLf & _
 "[Heal]" & vbCrLf & _
 "Heal HP potion/UH has priority over MP" & vbCrLf & _
 "Type F1~ buttons instead writing the spell" & vbCrLf & _
 "" & vbCrLf & _
 "[CaveBot]" & vbCrLf & _
 "Still in progress, doesnt work" & vbCrLf & _
 "" & vbCrLf & _
 "[Address]" & vbCrLf & _
 "Use Cheat Engine to get Address, it's very easy" & vbCrLf & _
 "You can use this bot to ANY OTClient OTserver" & vbCrLf & _
 "Check out tutorials at TibiaKing, TPForums, TibiaPF" & vbCrLf & _
 "" & vbCrLf & _
 "Contact me at: facebook.com/TibiaOldSchools"), vbOKOnly, "Info"
 
End Sub

Private Sub mCaveBot_Click()
  frmCavebot.WindowState = vbNormal
  frmCavebot.Show
  frmCavebot.SetFocus
  SetWindowPos frmCavebot.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub mDeveloper_Click()
  frmDebug.WindowState = vbNormal
  frmDebug.Show
  frmDebug.SetFocus
  SetWindowPos frmDebug.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub txtFlash_Change()

If IsNumeric(txtFlash) = True Then
    ' ok
Else
    txtFlash.Text = "0"
End If

End Sub

Private Sub tmrMain_Timer()
Dim tibiaclient As Long
Dim letra As String
Dim x As String

tibiaclient = TibiaWindow

'    If GetAsyncKeyState(123) < 0 And GetAsyncKeyState(16) < 0 Then   '   BOT SHOW SHIFT + F12
'        If Menu.Visible = False Then
'           Menu.Show
'           SetWindowPos Menu.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
'        Else
'           Menu.Hide
'        End If
'    End If

'''''''''''''''''''''''''''''''''''' HEALING ADDRESS
OTClientHP

''''''''''''''''''''''''''''''''''''
OTClientMyStatus

'''''''''''''''''''''''''''''''''''' EAT FOOD
If chkEat.Value = 1 And TibiaWindow <> 0 Then
    lngEat = lngEat + 1
    x = cmbEat.Text
    If lngEat = 100 Then
        hotkey x
        lngEat = 0
    End If
Else
    lngEat = 0
End If

'''''''''''''''''''''''''''''''''''' ANTI IDLE
'If chkIdle.Value = 1 And TibiaWindow <> 0 Then
'    lngIdle = lngIdle + 1
'    If lngIdle = 100 Then
'
'PostMessage tibiaclient, WM_KEYDOWN, VK_CONTROL, lparamvar
'PostMessage tibiaclient, WM_KEYDOWN, VK_UP, lparamvar
'PostMessage tibiaclient, WM_KEYUP, VK_UP, lparamvar
'Wait 1
'PostMessage tibiaclient, WM_KEYDOWN, VK_DOWN, lparamvar
'PostMessage tibiaclient, WM_KEYUP, VK_DOWN, lparamvar
'PostMessage tibiaclient, WM_KEYUP, VK_CONTROL, lparamvar
'
'        lngIdle = 0
'    End If
'Else
'    lngIdle = 0
'End If

'''''''''''''''''''''''''''''''''''' MANA TRAINER
If chkTrain.Value = 1 And TibiaWindow <> 0 Then
    OTClientManaTrainer
End If

'''''''''''''''''''''''''''''''''''' AUTO MANA SHIELD
If chkUtamo.Value = 1 And TibiaWindow <> 0 Then
    OTClientAutoUtamo
End If

'''''''''''''''''''''''''''''''''''' AUTO HUR
If chkHur.Value = 1 And TibiaWindow <> 0 Then
    OTClientAutoHur
End If

'''''''''''''''''''''''''''''''''''' SPEED HACK
If chkSpeed.Value = 1 And TibiaWindow <> 0 Then
    UpdateSpeed
End If

'''''''''''''''''''''''''''''''''''' LIGHT HACK
If chkLight.Value = 1 And TibiaWindow <> 0 Then
    UpdateLightOTClient
End If

'''''''''''''''''''''''''''''''''''' FLASH TIBIA
If chkFlash.Value = 1 Then
 If MyHP < CLng(txtFlash.Text) And TibiaWindow <> 0 Then
  Call FlashWindow(tibiaclient, True)
  sndPlaySound App.Path & "\ding.wav", 0
  exaustFlash = exaustFlash + 1
        If exaustFlash < 5 Then
            'nothing to do
        Else
            Call FlashWindow(tibiaclient, True)
            'sndPlaySound App.Path & "\ding.wav", 0
            exaustFlash = 0
        End If
 End If
End If

End Sub

Private Sub txtHealPot_Change()

If IsNumeric(txtHealPot) = True Then
    ' ok
Else
    txtHealPot.Text = "0"
End If

End Sub

Private Sub txtHi_Change()

If IsNumeric(txtHi) = True Then
    ' ok
Else
    txtHi.Text = "0"
End If

End Sub

Private Sub txtLow_Change()

If IsNumeric(txtLow) = True Then
    ' ok
Else
    txtLow.Text = "0"
End If

End Sub

Private Sub txtManaPot_Change()

If IsNumeric(txtManaPot) = True Then
    ' ok
Else
    txtManaPot.Text = "0"
End If

End Sub

Private Sub txtManaTrain_Change()

If IsNumeric(txtManaTrain) = True Then
    ' ok
Else
    txtManaTrain.Text = "0"
End If

End Sub

Private Sub txtMPHi_Change()

If IsNumeric(txtMPHi) = True Then
    ' ok
Else
    txtMPHi.Text = "0"
End If

End Sub

Private Sub txtMPLow_Change()

If IsNumeric(txtMPLow) = True Then
    ' ok
Else
    txtMPLow.Text = "0"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
' fun�ao pra quando fechar o programa, perguntar se quer fechar

    If MsgBox("Close Bot?", vbYesNo + vbQuestion, "Exit") = vbYes Then
        End
    Else
       Cancel = True
    End If
    
End Sub

Private Sub LoadSettingsFromFile()
    On Error GoTo goterr
    
    Dim strRes As String
    Dim strRes2 As String
    Dim sOpen As SelectedFile
    Dim count As Integer
    Dim FileList As String
    
    FileDialog.sFilter = "Ini (*.ini)" & Chr$(0) & "*.ini"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    'FileDialog.sDlgTitle = "Show Open"
    FileDialog.sInitDir = App.Path & "\Save"
    sOpen = ShowOpen(Me.hwnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        strRes = sOpen.sLastDirectory & sOpen.sFiles(1)
        strRes2 = LoadSettings(strRes)
        If strRes2 = "" Then
            'UpdateFormsFromVars
            'MsgBox BString(30) & strRes, vbOKOnly + vbInformation, BString(29)
        Else
            MsgBox strRes2, vbOKOnly + vbCritical, "Error loading " & strRes
        End If
    End If
    Exit Sub
            
goterr:
    If Err.Number <> 32755 Then
        MsgBox "Unexpected error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "LoadSettingsFromFile"
    End If
End Sub

Private Sub SaveSettingsToFile()
    On Error GoTo goterr
    
    Dim strRes As String
    Dim strRes2 As String
    Dim sOpen As SelectedFile
    Dim count As Integer
    Dim FileList As String
    
    FileDialog.sFilter = "Ini (*.ini)" & Chr$(0) & "*.ini"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    'FileDialog.sDlgTitle = "Show Open"
    FileDialog.sInitDir = App.Path & "\Save"
    sOpen = ShowSave(Me.hwnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        strRes = sOpen.sLastDirectory & sOpen.sFiles(1)
        If Right$(strRes, 4) <> ".ini" Then
            strRes = strRes & ".ini"
        End If
        strRes2 = SaveSettings(strRes)
        If strRes2 = "" Then
            'MsgBox BString(32) & strRes, vbOKOnly + vbInformation, BString(31)
        Else
            MsgBox strRes2, vbOKOnly + vbCritical, "Error saving " & strRes
        End If
    End If
    Exit Sub
    
goterr:
    If Err.Number <> 32755 Then
        MsgBox "Unexpected error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    End If
    
End Sub

Private Sub txtSpellHi_Change()

If txtSpellHi.Text = "" Then
    txtSpellHi.Text = "?"
Else
    'ok
End If

End Sub

Private Sub txtSpellLow_Change()

If txtSpellLow.Text = "" Then
    txtSpellLow.Text = "?"
Else
    'ok
End If

End Sub

Private Sub txtTrainSpell_Change()

If txtTrainSpell.Text = "" Then
    txtTrainSpell.Text = "?"
Else
    'ok
End If

End Sub
