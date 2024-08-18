VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TeacherReg 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19080
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   495
      Left            =   0
      TabIndex        =   70
      Top             =   8880
      Width           =   19215
      Begin VB.CommandButton new 
         Caption         =   "NEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   75
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton save 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   74
         Top             =   0
         Width           =   2055
      End
      Begin VB.CommandButton update 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10440
         TabIndex        =   73
         Top             =   0
         Width           =   2055
      End
      Begin VB.CommandButton delete 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13080
         TabIndex        =   72
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox serch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3480
         TabIndex        =   71
         Text            =   "SERCH"
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Qualilfication"
      Height          =   2175
      Left            =   0
      TabIndex        =   43
      Top             =   4800
      Width           =   19215
      Begin VB.CommandButton Command7 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16680
         TabIndex        =   61
         Top             =   600
         Width           =   1215
      End
      Begin VB.ListBox percentage_list 
         Height          =   960
         Left            =   13440
         TabIndex        =   55
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ListBox pass_year_list 
         Height          =   960
         Left            =   10560
         TabIndex        =   54
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ListBox quali_board_list 
         Height          =   960
         Left            =   5640
         TabIndex        =   53
         Top             =   1080
         Width           =   4575
      End
      Begin VB.ListBox thr_quali_list 
         Height          =   960
         Left            =   1440
         TabIndex        =   52
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox quali_board 
         Height          =   420
         Left            =   5640
         TabIndex        =   51
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox pass_year 
         Height          =   420
         Left            =   10560
         TabIndex        =   50
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox percentage 
         Height          =   420
         Left            =   13440
         TabIndex        =   49
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox thr_quali 
         Height          =   420
         Left            =   1440
         TabIndex        =   48
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         Height          =   300
         Left            =   13920
         TabIndex        =   47
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passing Year"
         Height          =   300
         Left            =   10920
         TabIndex        =   46
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Board/Univercity"
         Height          =   300
         Left            =   6960
         TabIndex        =   45
         Top             =   240
         Width           =   1980
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qualification"
         Height          =   300
         Left            =   2520
         TabIndex        =   44
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Teaching Experience"
      Height          =   1935
      Left            =   0
      TabIndex        =   36
      Top             =   6960
      Width           =   19215
      Begin VB.ListBox work_exp_list 
         Height          =   960
         Left            =   13560
         TabIndex        =   69
         Top             =   840
         Width           =   2175
      End
      Begin VB.ListBox post_list 
         Height          =   960
         Left            =   7920
         TabIndex        =   68
         Top             =   840
         Width           =   3135
      End
      Begin VB.CommandButton Command8 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16680
         TabIndex        =   62
         Top             =   360
         Width           =   1215
      End
      Begin VB.ListBox sch_name_list 
         Height          =   960
         Left            =   2520
         TabIndex        =   59
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox work_exp 
         Height          =   420
         Left            =   13560
         TabIndex        =   42
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox post 
         Height          =   420
         Left            =   7920
         TabIndex        =   40
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox sch_name 
         Height          =   420
         Left            =   2520
         TabIndex        =   38
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Work Experience"
         Height          =   300
         Left            =   11400
         TabIndex        =   41
         Top             =   480
         Width           =   2040
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post"
         Height          =   300
         Left            =   7200
         TabIndex        =   39
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Name"
         Height          =   300
         Left            =   480
         TabIndex        =   37
         Top             =   480
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   3855
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   19215
      Begin VB.TextBox religion 
         Height          =   420
         Left            =   13200
         TabIndex        =   67
         Top             =   1200
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   13200
         TabIndex        =   64
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125304833
         CurrentDate     =   44662
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11760
         TabIndex        =   60
         Top             =   3240
         Width           =   1095
      End
      Begin VB.ListBox sub_list 
         Height          =   660
         Left            =   13080
         TabIndex        =   58
         Top             =   3120
         Width           =   3255
      End
      Begin VB.ComboBox subject 
         Height          =   420
         Left            =   8400
         TabIndex        =   57
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox salary 
         Height          =   420
         Left            =   2400
         TabIndex        =   35
         Top             =   3240
         Width           =   2895
      End
      Begin VB.ComboBox doc 
         Height          =   420
         Left            =   2400
         TabIndex        =   33
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox email 
         Height          =   420
         Left            =   13200
         TabIndex        =   32
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox city 
         Height          =   420
         Left            =   8400
         LinkTimeout     =   0
         TabIndex        =   31
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "UPLOAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17280
         TabIndex        =   30
         Top             =   3120
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   16680
         ScaleHeight     =   2115
         ScaleWidth      =   2235
         TabIndex        =   29
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox state 
         Height          =   420
         Left            =   2400
         TabIndex        =   27
         Top             =   2280
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Female"
         Height          =   375
         Left            =   9480
         TabIndex        =   26
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Male"
         Height          =   375
         Left            =   8400
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox fname 
         Height          =   420
         Left            =   2400
         TabIndex        =   12
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox thr_name 
         Height          =   420
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox mname 
         Height          =   420
         Left            =   8400
         TabIndex        =   10
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox address 
         Height          =   420
         Left            =   2400
         TabIndex        =   9
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox pincode 
         Height          =   420
         Left            =   13200
         TabIndex        =   8
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox phone_no 
         Height          =   420
         Left            =   8400
         TabIndex        =   7
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox doc_no 
         Height          =   420
         Left            =   8400
         TabIndex        =   5
         Top             =   2760
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   13200
         TabIndex        =   6
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   97910785
         CurrentDate     =   44645
      End
      Begin VB.Label serial_no 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1680
         TabIndex        =   81
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial no"
         Height          =   300
         Left            =   360
         TabIndex        =   80
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label reg_date 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label24"
         Height          =   300
         Left            =   17040
         TabIndex        =   79
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   300
         Left            =   16200
         TabIndex        =   78
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teacher Id"
         Height          =   300
         Left            =   7560
         TabIndex        =   77
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label thr_id 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   300
         Left            =   9480
         TabIndex        =   76
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   11280
         TabIndex        =   66
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label gender 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ss"
         Height          =   375
         Left            =   9480
         TabIndex        =   65
         Top             =   840
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joinin Date"
         Height          =   300
         Left            =   11280
         TabIndex        =   63
         Top             =   2760
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   300
         Left            =   6240
         TabIndex        =   56
         Top             =   3360
         Width           =   930
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documents No."
         Height          =   300
         Left            =   6240
         TabIndex        =   34
         Top             =   2880
         Width           =   1860
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
         Height          =   300
         Left            =   6240
         TabIndex        =   28
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name"
         Height          =   300
         Left            =   360
         TabIndex        =   24
         Top             =   1440
         Width           =   1770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   300
         Left            =   360
         TabIndex        =   23
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         Height          =   300
         Left            =   11280
         TabIndex        =   22
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   300
         Left            =   6240
         TabIndex        =   21
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   300
         Left            =   6240
         TabIndex        =   19
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   300
         Left            =   360
         TabIndex        =   18
         Top             =   2400
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code"
         Height          =   300
         Left            =   11280
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documents"
         Height          =   300
         Left            =   360
         TabIndex        =   16
         Top             =   2880
         Width           =   1380
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         Height          =   300
         Left            =   360
         TabIndex        =   15
         Top             =   3360
         Width           =   885
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   300
         Left            =   11280
         TabIndex        =   14
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mother's Name"
         Height          =   300
         Left            =   6240
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19455
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MISSION BOARDING SCHOOL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   540
         Left            =   6720
         TabIndex        =   3
         Top             =   0
         Width           =   6705
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GAUR ROAD BAIRGANIYA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8880
         TabIndex        =   2
         Top             =   480
         Width           =   2310
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TEACHER REGISTRATION "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   8160
         TabIndex        =   1
         Top             =   720
         Width           =   3930
      End
   End
End
Attribute VB_Name = "TeacherReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Change()

End Sub

Private Sub Command6_Click()
If subject.Text = "" Then
MsgBox "Enter Subject Name"
subject.SetFocus
Else
sub_list.AddItem subject.Text
subject.Text = ""
End If
End Sub

Private Sub Command7_Click()
If thr_quali.Text = "" Then
MsgBox "Enter Teacher Qualification "
ElseIf quali_board.Text = "" Then
MsgBox "Enter Board/Univercity of Qulaification "
ElseIf pass_year.Text = "" Then
MsgBox "Enter passing year"
ElseIf percentage.Text = "" Then
MsgBox "Fill the total percentge of marks"
Else
thr_quali_list.AddItem thr_quali.Text
quali_board_list.AddItem quali_board.Text
pass_year_list.AddItem pass_year.Text
percentage_list.AddItem percentage.Text
thr_quali.Text = ""
quali_board.Text = ""
pass_year.Text = ""
percentage.Text = ""
End If
End Sub

Private Sub Command8_Click()
If sch_name.Text = "" Then
MsgBox "Enter School Name"
ElseIf post.Text = "" Then
MsgBox "Enter position in school"
ElseIf work_exp.Text = "" Then
MsgBox "Enter Work Experience"
Else
sch_name_list.AddItem sch_name.Text
post_list.AddItem post.Text
work_exp_list.AddItem work_exp.Text
sch_name.Text = ""
post.Text = ""
work_exp.Text = ""
End If
End Sub

Private Sub fname_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Then
Else
MsgBox "Enter Valid Name"
End If
End Sub

Private Sub Form_Load()
conn
sql = "select count(serial_no) from teacher_registration"
Set r = c.Execute(sql)
serial_no.Caption = r.Fields(0) + 1
reg_date.Caption = Format(Date, "dd mmm yyyy")
thr_id.Caption = "MBS" + Right(reg_date, 4) + "S" & serial_no.Caption
End Sub

Private Sub mname_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Then
Else
MsgBox "Enter Valid Name"
End If
End Sub

'Dim i As Integer
Private Sub Option1_Click()
gender.Caption = "Male"
End Sub

Private Sub Option2_Click()
gender.Caption = "Female"
End Sub

Private Sub percentage_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
MsgBox "Enter only digits"
percentage.Text = ""
KeyAscii = 0
End If
End Sub

Private Sub phone_no_KeyPress(KeyAscii As Integer)
phone_no.MaxLength = 10
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
MsgBox "Only Digits(0 to 9) are Allowed"
phone_no.SetFocus
End If
End Sub

Private Sub pincode_KeyPress(KeyAscii As Integer)
pincode.MaxLength = 6
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
MsgBox "Only Digits are Allowed"
End If
End Sub

Private Sub save_Click()
conn
sql = "insert into teacher_registration values(" + serial_no + ",'" + thr_id + "','" + reg_date + "', '" + thr_name + "','" + gender + " ' ,'" + Format(DTPicker1, "DD MMM YYYY") + "','" + fname + "','" + mname + "','" + religion + "','" + address + "','" + city + "' ," + pincode + ",'" + state + "' ,'" + phone_no + "','" + email + "' ,'" + doc + "' ,'" + doc_no + "','" + Format(DTPicker2, "dd MMM yyyy") + "')"
Set r = c.Execute(sql)
MsgBox "Record saved"
'for i=1 to
'sql = "insert into teacher_subject values('" + thr_id + ",'" + sub_list(i) + "')"
'Set r = c.Execute(sql)
'MsgBox "record save"
'Loop
End Sub

Private Sub thr_name_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Then
Else
MsgBox "Enter Valid Name"
End If
End Sub

