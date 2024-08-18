VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LoginForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14685
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7860
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   7200
      Top             =   6480
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=shivam/kumar;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=shivam/kumar;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7920
      TabIndex        =   5
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "login"
      Height          =   360
      Left            =   6600
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   480
      Left            =   5520
      TabIndex        =   1
      Text            =   "Password"
      Top             =   3480
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   480
      Left            =   5520
      TabIndex        =   0
      Text            =   "User Id"
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't have account? Sign up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Forget"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text2.PasswordChar = ""
ElseIf Check1.Value = 0 Then
Text2.PasswordChar = "*"
End If
End Sub

Private Sub Command1_Click()
Adodc1.RecordSource = "select *from login where log_id ='" + Text1.Text + "' and log_pass='" + Text2.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "login faild"
Else
Unload Me
MDIForm1.Show
'conn
'sql = "select log_id,log_pass from login"
'Set r = c.Execute(sql)
'a = r.Fields(0)
'b = r.Fields(1)
'If a = Text1.Text And b = Text2.Text Then
'Unload Me
'MDIForm1.Show
End If
End Sub

Private Sub Text1_Change()
Text1.ForeColor = vbBlack
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Change()
Text2.ForeColor = vbBlack
Text2.PasswordChar = "*"
If Check1.Value = 1 Then
Text2.PasswordChar = ""
End If
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub
