VERSION 5.00
Begin VB.Form frmFTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Existence"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "frmFTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Before you proceed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   5415
      Begin VB.Label Label6 
         Caption         =   $"frmFTP.frx":0CCA
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Timer Timer 
      Left            =   7200
      Top             =   5040
   End
   Begin VB.Frame Frame1 
      Caption         =   "FTP Information"
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5415
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtFolder 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   1620
         Width           =   4095
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   780
         Width           =   2775
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2085
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Remote Folder"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1665
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1245
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   825
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   405
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Does the file exist on Server?"
      Default         =   -1  'True
      Height          =   375
      Left            =   1485
      TabIndex        =   5
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lblWait 
      AutoSize        =   -1  'True
      Caption         =   "Please Wait...Checking on the server !!!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   525
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   4800
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()
On Error GoTo errHandle

'Author : Anil Raghav, Bangalore, India.

'This is a very simple program to check for the existence of a file on
'a remote FTP Server.  I am basically creating a FTP batchfile at runtime
'which when executed will create another file called Found.txt.  If the
'file search was successful, it would write the filename in the Found.txt file
'else it will be blank.  I then check the size of the file. If its grater than
'0, the file exists.  Simple!!! Is it not.

'I use a function called Supershell which is a variation of shell function.
'This function will make the program wait till the shell execution is complete.
'If ordinary shell function is used, the program will not work.

'If you like to vote....Please vote
    'Variable declaration.
    Dim hndFile         As Integer
    Dim strFTP          As String
    Dim strBAT          As String
    Dim strPath         As String
    
    Dim strContents     As String
    Dim intFTPCode      As Integer
    
    Dim strErrorMessage As String
    
    'Start the timer with interval as one second
    Screen.MousePointer = vbHourglass
    Timer.Interval = 1000
    
    m_blnError = False
    
    'Dynamically created files.....
    strPath = "C:\"
    strFTP = strPath & "FTPFind.txt"
    strBAT = strPath & "FTPFind.BAT"
    
    hndFile = FreeFile
    
    'Create the FTP script file
    Open strFTP For Output As #hndFile
    'Open connection with server
    Print #hndFile, "open " & Trim(txtIP.Text)
    'Supply the username.  Syntax : user <username>
    Print #hndFile, "user " & Trim(txtUsername.Text)
    'Supply the password
    Print #hndFile, Trim(txtPassword.Text)
    'Change directory
    Print #hndFile, "cd " & Trim(txtFolder.Text)
    'Very important. Execute ls command.  Syntax ls <filetobesearched> <localfile>
    'The result of ls command is written to the localfile.
    Print #hndFile, "ls " & Trim(txtFilename.Text) & " c:\found.txt"
    'Close the connection
    Print #hndFile, "close"
    'Bye
    Print #hndFile, "bye"
    'Close the filehandle
    Close #hndFile
    hndFile = FreeFile
    
    'Batchfile uses the file created above as the script file
    'Syntax : FTP -n -s:<scriptfile>
    Open strBAT For Output As #hndFile
        Print #hndFile, "ftp -n -s:""" & strFTP
    Close #hndFile
    
    'Call the supershell function to execute the batch file
    SuperShell (strBAT)

    Open "C:\found.txt" For Input As #1
        'Get the size of the file
        Filesize = LOF(1)
    Close #1
    
    lblWait.Visible = False
    
    'Diable the timer.
    
    Timer.Interval = 0
    Screen.MousePointer = vbArrow
    'Display the search result to the user.
    If Filesize > 0 Then
        MsgBox "File EXISTS on the remote server!", vbOKOnly, "Success!"
    Else
        MsgBox "File DOES NOT EXIST on the remote server!", vbOKOnly, "Failure!"
    End If
    'Destroy all the files generated run-time
    
    Kill ("C:\FTPFind.txt")
    Kill ("C:\FTPFind.bat")
    Kill ("C:\found.txt")
    Exit Sub
errHandle:

    If Err.Number > 0 Then 'If there is an error
        'Show the error and setback the timer....
        MsgBox "Error : " & Err.Description
        Timer.Interval = 0
        Screen.MousePointer = vbArrow
        Err.Clear
    End If
End Sub

Private Sub Timer_Timer()
    'Show the please wait message to the user.
    lblWait.Visible = Not (lblWait.Visible)
End Sub
