VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Quit (ESC)"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Unregister"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Register"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Dim FirstInstallation As String, TimesWorked As Long, LicencedUser As String
Dim Licenc As Long, DemoVersion As Boolean, ReadSeries As String

Private Sub AddText(a As String)

    Text1.Text = Text1.Text & vbCrLf & a
    
End Sub


Private Sub InitializeSystem()
On Error GoTo erro

    AddText ("Initializing...")

    Dim volbuf$, sysname$, serialnum&, sysflags&, componentlength&, res&
    volbuf$ = String$(256, 0)
    sysname$ = String$(256, 0)
    res = GetVolumeInformation("C:\", volbuf$, 255, serialnum, _
            componentlength, sysflags, sysname$, 255)
                 
    AddText ("HD's serial number got: " & serialnum)
    
    'This is the math expression you can apply to get the registering code.
    'Of course, you must build another app that gets the user code and returns the
    'registration code, wich you pass to the user.
    Licenc = Int(2802 * Sqr(serialnum))
    
    AddText ("Licence code is " & Licenc & ", use it when registering the software.")
    
    'Lê data da 1ª instalação:
    
    Dim FirstInstallationSaved As String, ReadDate As String, DateOk As String, FirstTime As Boolean
    ReadDate = GetSetting("DemoApp", "Install", "Installation", "xxx")
    
    If Not ReadDate = "xxx" Then
        DateOk = Decrypt(ReadDate, "alex")
        FirstInstallation = DateOk
        AddText ("FirstInstallation read: " & FirstInstallation)
    Else        'Nothing saved, this is the first time...
        'FirstInstallation = Left(Date, 2) & Mid(Date, 4, 2) & Right(Date, 2)
        FirstInstallation = Date
        FirstInstallationSaved = Encrypt(FirstInstallation, "alex")
        SaveSetting "DemoApp", "Install", "Installation", FirstInstallationSaved
        AddText ("No FirstInstallation saved, doing it now.")
        FirstTime = True
    End If
        
    'Reads encrypted serial number:
    ReadSeries = GetSetting("DemoApp", "Install", "Series", "0")
    
    If ReadSeries = "0" Then      'Doesn't exist, creating one
        DemoVersion = True
        Me.Caption = "Demo App - THIS IS A DEMO VERSION!"
        Command2.Enabled = False
        Dim LimitDate As Date, TimesWorked As Long, TimesWorkedRead As String, TimesWorkedSaved As String
        TimesWorkedRead = GetSetting("DemoApp", "Install", "TimesWorked", "0")
        If Not TimesWorkedRead = "0" Then
            TimesWorked = Decrypt(TimesWorkedRead, "alex")
        Else
            TimesWorked = 0
        End If
        
        'Giving the user 1 month to use the demo
        LimitDate = DateAdd("m", 1, FirstInstallation)
        
        If (TimesWorked >= 100 Or LimitDate < Date) And Not FirstTime Then
            Me.Caption = "Demo App - EXPIRED!!!"
            AddText ("")
            AddText ("This Demo version has EXPIRED!!!")
            AddText ("")
            AddText ("Open Registry Editor and delete the key")
            AddText ("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\DemoApp\Install")
            Command1.Enabled = False
            'End    'Disable further use of the app by the user
        Else
            TimesWorked = TimesWorked + 1
            TimesWorkedSaved = Encrypt(CStr(TimesWorked), "alex")
            SaveSetting "DemoApp", "Install", "TimesWorked", TimesWorkedSaved
            AddText ("This is a DEMO version. You can use it for 1 month or 100 times!")
            AddText ("Times worked: " & TimesWorked & "       First installation: " & FirstInstallation)
            
            'Verify the TimesWorked variable:
            If DemoVersion And TimesWorked >= 95 Then
                If TimesWorked = 100 Then
                    AddText ("")
                    AddText ("WARNING!!   This is the LAST TIME you can run this DEMO version!!!")
                ElseIf TimesWorked = 99 Then
                    AddText ("")
                    AddText ("WARNING!!   You can run only ONE MORE TIME this app!!")
                Else
                    AddText ("")
                    AddText ("WARNING!!   You can run this app " & 100 - TimesWorked & " more times.")
                End If
            End If
            
            'Verify the FirstInstallation variable:
            If Not FirstTime And DemoVersion And DateDiff("d", LimitDate, Date) * (-1) <= 5 Then
                If DateDiff("d", LimitDate, Date) = 0 Then
                    AddText ("")
                    AddText ("WARNING!!   This is the LAST DAY you can run this demo version!")
                ElseIf DateDiff("d", LimitDate, Date) = -1 Then
                    AddText ("")
                    AddText ("WARNING!!   You have only ONE MORE DAY to run this demo version!")
                Else
                    AddText ("")
                    AddText ("WARNING!!   You have " & DateDiff("d", LimitDate, Date) * (-1) & " days to run this demo version!")
                End If
            End If
        End If
    
    ElseIf Decrypt(ReadSeries, "alex") <> CStr(Licenc) Then
        AddText ("The licence code for this app is wrong.   Please contact the support!")
        'End       'Someone have tried to alter the licence, or copy the entire Windows registry
                    'from a registered machine to another one...
    End If
    
    If DemoVersion = False Then
        Dim e As String
        e = GetSetting("DemoApp", "Install", "LicencedUser")
        Command1.Enabled = False
        LicencedUser = Decrypt(e, "alex")
        Me.Caption = "Demo App - REGISTERED VERSION to " & LicencedUser
        AddText ("Registered version to " & LicencedUser)
        
        'you can continue to count the times the app has worked:
        TimesWorkedRead = GetSetting("DemoApp", "Install", "TimesWorked", "0")
        TimesWorked = Decrypt(TimesWorkedRead, "alex")
        TimesWorked = TimesWorked + 1
        TimesWorkedSaved = Encrypt(CStr(TimesWorked), "alex")
        SaveSetting "DemoApp", "Install", "TimesWorked", TimesWorkedSaved
        AddText ("Worked " & TimesWorked & " times.")
    End If
        
        
        
saída:
    Exit Sub
    
erro:
    MsgBox "There was an error:" & vbLf & vbLf & Err.Number & " - " & Err.Description, vbCritical
    Resume saída

End Sub

Public Function Decrypt(texti, salasana)
On Error Resume Next

    Dim t As Byte, sana As String, x1 As Integer, g As Integer, tt As Byte, DeCrypted As String
    
    For t = 1 To Len(salasana)
        sana = Asc(Mid(salasana, t, 1))
        x1 = x1 + sana
    Next

    x1 = Int((x1 * 0.1) / 6)
    salasana = x1
    g = 0

    For tt = 1 To Len(texti)
        sana = Asc(Mid(texti, tt, 1))
        g = g + 1
        If g = 6 Then g = 0
        x1 = 0
        If g = 0 Then x1 = sana + (salasana - 2)
        If g = 1 Then x1 = sana - (salasana - 5)
        If g = 2 Then x1 = sana + (salasana - 4)
        If g = 3 Then x1 = sana - (salasana - 2)
        If g = 4 Then x1 = sana + (salasana - 3)
        If g = 5 Then x1 = sana - (salasana - 5)
        x1 = x1 - g
        DeCrypted = DeCrypted & Chr(x1)
    Next

    Decrypt = DeCrypted

End Function
Public Function Encrypt(texti, salasana)
On Error Resume Next

    Dim t As Byte, tt As Byte, sana As String, x1 As Integer, g As Integer, Crypted As String
    For t = 1 To Len(salasana)
        sana = Asc(Mid(salasana, t, 1))
        x1 = x1 + sana
    Next

    x1 = Int((x1 * 0.1) / 6)
    salasana = x1
    g = 0

    For tt = 1 To Len(texti)
        sana = Asc(Mid(texti, tt, 1))
        g = g + 1
        If g = 6 Then g = 0
        x1 = 0
        If g = 0 Then x1 = sana - (salasana - 2)
        If g = 1 Then x1 = sana + (salasana - 5)
        If g = 2 Then x1 = sana - (salasana - 4)
        If g = 3 Then x1 = sana + (salasana - 2)
        If g = 4 Then x1 = sana - (salasana - 3)
        If g = 5 Then x1 = sana + (salasana - 5)
        x1 = x1 + g
        Crypted = Crypted & Chr(x1)
    Next

    Encrypt = Crypted

End Function




Private Sub Command1_Click()
On Error GoTo erro

    Dim volbuf$, sysname$, serialnum&, sysflags&, componentlength&, res&
    volbuf$ = String$(256, 0)
    sysname$ = String$(256, 0)
    res = GetVolumeInformation("C:\", volbuf$, 255, serialnum, _
            componentlength, sysflags, sysname$, 255)
                        
    'The math expression:
    Licenc = Int(2802 * Sqr(serialnum))
    
    Dim k As String
    k = InputBox("Please input the registration code to this machine:", "Registration")
    If Len(k) = 0 Then Exit Sub
    If k <> Licenc Then
CodErro:
            MsgBox "Nice try, but it's an invalid code.", vbCritical
    Else
        Dim a As String, b As String, c As String
        c = InputBox("Registered user:", "Registration")
            If Len(c) = 0 Then MsgBox "Registration canceled!", vbCritical: Exit Sub
            If MsgBox("Registered user:" & vbLf & vbLf & "" & c & "" & _
                vbLf & vbLf & "Confirm?", vbYesNo + vbQuestion) = vbYes Then
            a = Encrypt(CStr(Licenc), "alex")
            b = Encrypt(c, "alex")
            SaveSetting "DemoApp", "Install", "Series", a
            SaveSetting "DemoApp", "Install", "LicencedUser", b
            LicencedUser = c
            Command1.Enabled = False
            MsgBox "Thanks for registering this app, etc etc...", vbInformation
            Me.Caption = "Demo App - REGISTERED VERSION to " & c
            DemoVersion = False
        End If
    End If
    
saída:
    Exit Sub
    
erro:
    If Err.Number = 13 Then
        Resume CodErro
    Else
        MsgBox "There was an error:" & vbLf & vbLf & Err.Number & " - " & Err.Description, vbCritical
    Resume saída
    End If
    
End Sub

Private Sub Command2_Click()
On Error Resume Next

    If MsgBox("You will cancel the registration." & vbLf & vbLf & "Ok?", vbYesNo) = vbYes Then
        DeleteSetting "DemoApp", "Install", "LicencedUser"
        DeleteSetting "DemoApp", "Install", "Series"
        Me.Caption = "Demo App - THIS IS A DEMO VERSION!"
        Command1.Enabled = True
        Command2.Enabled = False
    End If
    
End Sub


Private Sub Command3_Click()

    Unload Me
    End
    
End Sub

Private Sub Form_Activate()

    InitializeSystem
    
End Sub

