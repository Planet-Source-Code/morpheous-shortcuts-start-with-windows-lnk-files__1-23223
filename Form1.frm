VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShortCut Creator"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Add to TaskBar"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Default"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "C:\WINDOWS\Start Menu\Programs\StartUp\"
      Top             =   2880
      Width           =   4215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Form1.frx":0442
      Top             =   120
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Text            =   "ShortCut"
      Top             =   1560
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "......"
      Height          =   285
      Left            =   6120
      TabIndex        =   7
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "......"
      Height          =   285
      Left            =   6120
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton cmdCreateLink 
      Caption         =   "Create"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Destination:"
      Height          =   195
      Left            =   2400
      TabIndex        =   14
      Top             =   2640
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   2400
      TabIndex        =   13
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Exe:"
      Height          =   195
      Left            =   2400
      TabIndex        =   12
      Top             =   2280
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Left            =   2400
      TabIndex        =   11
      Top             =   1920
      Width           =   285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHFileExists Lib "Shell32" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Const conSwNormal = 1

Private Sub cmdCreateLink_Click()
Dim Tex As String
Dim RetVal As String
Dim mIcon As Long

If Text1.Text = "" Or Text2.Text = "" Then 'FileName and Executable name cannot be blank
MsgBox ("Please enter the FileName and Executable!"), vbInformation + vbOKOnly, "ShortCut Creator v" & App.Major & "." & App.Minor & "." & App.Revision

Exit Sub

Else

If Str$(SHFileExists(Text5.Text & Text3.Text & ".lnk")) = 1 Then 'Full path to the shortcut

MsgBox ("A file with the same name already exists at that location!"), vbInformation + vbOKOnly, "ShortCut Creator v" & App.Major & "." & App.Minor & "." & App.Revision

Exit Sub

Else

fCreateShellLink Text1.Text & ".lnk", Text2.Text, "", Text1.Text, 0, 0, SHOWNORMAL

If Text5.Text = "" Then
Tex = "C:\WINDOWS\Start Menu\Programs\StartUp\" & Text3.Text & ".lnk"

Else
Tex = Text5.Text & Text3.Text & ".lnk"

End If

'This next line moves the shortcut from its original path to wherever you want it to
'be according to the text in text5.text

Name Text1.Text & ".lnk" As Tex

'This line creates a shortcut on the taskbar

If Check1.Value = 1 Then
CopyFile Tex, "C:\WINDOWS\Application Data\Microsoft\Internet Explorer\Quick Launch\" & Text3.Text & ".lnk", 1
End If

RetVal = MsgBox("Your shortcut has been created and is located at " & Text5.Text & vbCrLf & "Would you like to locate your shortcut now?", vbInformation + vbOKCancel, "ShortCut Creator v" & App.Major & "." & App.Minor & "." & App.Revision)

If RetVal = vbCancel Then

Exit Sub

Else

If Text5.Text = "C:\WINDOWS\Desktop\" Then

WinExec "Explorer.exe C:\WINDOWS\Desktop", 10 'This function returns the desktop folder which is a special folder!

Else

ShellExecute hWnd, "Open", Text5.Text, vbNullString, vbNullString, conSwNormal

End If

End If

End If

End If

End Sub

Private Sub Command1_Click()
CD1.Filter = "All Files (*.*)|*.*"
CD1.ShowOpen
Text1.Text = CD1.FileName
End Sub

Private Sub Command2_Click()
CD1.Filter = "Program Files (*.exe)|*.exe"
CD1.ShowOpen
Text2.Text = CD1.FileName
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Text3.Text = "ShortCut"
Text5.Text = "C:\WINDOWS\Start Menu\Programs\StartUp\"
Dir1.Path = "C:\WINDOWS\Start Menu\Programs\StartUp\"
Drive1.Drive = "C:\"
End Sub

Private Sub Dir1_Change()
If Len(Dir1.Path) <> 3 Then
Text5.Text = Dir1.Path & "\"
Else
Text5.Text = Dir1.Path
End If
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Dir1.Path = "C:\WINDOWS\Start Menu\Programs\StartUp\"
End Sub




