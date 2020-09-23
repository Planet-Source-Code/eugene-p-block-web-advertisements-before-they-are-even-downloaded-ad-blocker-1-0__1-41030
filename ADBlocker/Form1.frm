VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "AD Blocker - By Eugene - http://www.eugenius.tk"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Block !"
      Height          =   405
      Left            =   105
      TabIndex        =   8
      Top             =   2490
      Width           =   1110
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   105
      TabIndex        =   6
      Top             =   1710
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2565
      ItemData        =   "Form1.frx":0000
      Left            =   3225
      List            =   "Form1.frx":0A4E
      TabIndex        =   4
      Top             =   360
      Width           =   3045
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Operating System Detection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   3000
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Windows XP"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Windows 95/98/ME"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   570
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Windows NT/2000"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   870
         Width           =   1935
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VOTE FOR ME !!"
      Height          =   195
      Left            =   1680
      TabIndex        =   10
      Top             =   2400
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.eugenius.tk"
      Height          =   195
      Left            =   1425
      TabIndex        =   9
      Top             =   2190
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   90
      X2              =   3105
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   1515
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AD Server Database:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3225
      TabIndex        =   5
      Top             =   135
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim X As Integer
X = 0
ProgressBar1.Max = List1.ListCount

If Option1.Value = True Then 'IF WINDOWS XP is SELECTED !!
    Open "c:\windows\system32\drivers\etc\hosts.txt" For Append As #1 'OPEN the HOSTS File For Appending Write
    Print #1, vbCrLf 'Print a Carrige Return
    Print #1, "# AD Blocker List By - Eugene"
    
    Do Until X = List1.ListCount 'Begin a Do Until Loop
    DoEvents 'Do events
        Print #1, List1.List(X) 'Append the Server from the Listbox
        X = X + 1 'add 1 to X so next leep we will get onto the NEXT ITEM in the List
        ProgressBar1.Value = X 'increase the Progressbar value by 1
    Loop 'loop (for the new to VB people, it goes back to "Print #1, List1.List(X)" 3 lines up
    Close #1 'Close the file, save the writing
    
    MsgBox "AD Blocker ! Injected into the system, process complete ! Closing AD Blocker ! , thank you for using AD Blocker !  !", vbInformation, "Complete !"
    End
End If
Exit Sub

If Option2.Value = True Then
    Open "c:\windows\hosts" For Append As #1
    Print #1, vbCrLf 'Print a Carrige Return
    Print #1, "# AD Blocker List By - Eugene"
    
    Do Until X = List1.ListCount 'Begin a Do Until Loop
    DoEvents 'Do events
        Print #1, List1.List(X) 'Append the Server from the Listbox
        X = X + 1 'add 1 to X so next leep we will get onto the NEXT ITEM in the List
        ProgressBar1.Value = X 'increase the Progressbar value by 1
    Loop 'loop (for the new to VB people, it goes back to "Print #1, List1.List(X)" 3 lines up
    Close #1 'Close the file, save the writing
    
    MsgBox "AD Blocker ! Injected into the system, process complete ! Closing AD Blocker ! , thank you for using AD Blocker !  !", vbInformation, "Complete !"
    End
End If
Exit Sub

If Option3.Value = True Then
    Open "c:\winnt\system32\drivers\etc\hosts" For Append As #1
    Print #1, vbCrLf 'Print a Carrige Return
    Print #1, "# AD Blocker List By - Eugene"
    
    Do Until X = List1.ListCount 'Begin a Do Until Loop
    DoEvents 'Do events
        Print #1, List1.List(X) 'Append the Server from the Listbox
        X = X + 1 'add 1 to X so next leep we will get onto the NEXT ITEM in the List
        ProgressBar1.Value = X 'increase the Progressbar value by 1
    Loop 'loop (for the new to VB people, it goes back to "Print #1, List1.List(X)" 3 lines up
    Close #1 'Close the file, save the writing
    
    MsgBox "AD Blocker ! Injected into the system, process complete ! Closing AD Blocker ! , thank you for using AD Blocker !  !", vbInformation, "Complete !"
    End
End If
Exit Sub
End Sub

Private Sub Form_Load()
If FExist(App.Path & "\database.txt") = True Then
    Call Loadlistbox(App.Path & "\database.txt", List1)
Else
    Call SaveListBox(App.Path & "\database.txt", List1)
End If

Label1.Caption = "AD Server Database: " & List1.ListCount & " blocked"

End Sub
