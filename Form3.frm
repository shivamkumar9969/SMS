VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form AdmissionForm 
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18750
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   18750
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "NEW"
      Height          =   495
      Left            =   840
      TabIndex        =   73
      Top             =   8880
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   480
      Left            =   11160
      TabIndex        =   72
      Text            =   "Search"
      Top             =   8880
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   3840
      TabIndex        =   71
      Top             =   8880
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   18735
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADMISSION FORM"
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   8640
         TabIndex        =   69
         Top             =   1200
         Width           =   2685
      End
      Begin VB.Label Label3 
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
         TabIndex        =   68
         Top             =   840
         Width           =   2310
      End
      Begin VB.Label Label2 
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
         TabIndex        =   67
         Top             =   240
         Width           =   6705
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parent's Details"
      Height          =   2295
      Left            =   0
      TabIndex        =   42
      Top             =   6480
      Width           =   18735
      Begin VB.TextBox mname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   52
         Top             =   840
         Width           =   3735
      End
      Begin VB.ComboBox fqualification 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form3.frx":0000
         Left            =   10440
         List            =   "Form3.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox foccupation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form3.frx":0004
         Left            =   15600
         List            =   "Form3.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox moccupation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form3.frx":0008
         Left            =   15600
         List            =   "Form3.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   840
         Width           =   2895
      End
      Begin VB.ComboBox mqualification 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form3.frx":000C
         Left            =   10440
         List            =   "Form3.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox parents_doc_no 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10440
         TabIndex        =   47
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox fname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   46
         Top             =   360
         Width           =   3735
      End
      Begin VB.ComboBox parents_doc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form3.frx":0010
         Left            =   3120
         List            =   "Form3.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox email 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   44
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox phone_no 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10440
         TabIndex        =   43
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name"
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
         Left            =   480
         TabIndex        =   62
         Top             =   480
         Width           =   1770
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qulification"
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
         Left            =   8520
         TabIndex        =   61
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label occupation 
         AutoSize        =   -1  'True
         Caption         =   "Occupation"
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
         Left            =   13920
         TabIndex        =   60
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mother's Name"
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
         Left            =   480
         TabIndex        =   59
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label mmqualification 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qulification"
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
         Left            =   8520
         TabIndex        =   58
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Occupation"
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
         Left            =   13920
         TabIndex        =   57
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parent's Document"
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
         Index           =   0
         Left            =   480
         TabIndex        =   56
         Top             =   1440
         Width           =   2310
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8520
         TabIndex        =   55
         Top             =   1440
         Width           =   1725
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
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
         Left            =   480
         TabIndex        =   54
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
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
         Left            =   8520
         TabIndex        =   53
         Top             =   1920
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Information"
      Height          =   5175
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   18735
      Begin VB.TextBox relligion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10320
         TabIndex        =   41
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox session 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   24
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox admission_no 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         HideSelection   =   0   'False
         Left            =   3120
         TabIndex        =   23
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox roll_no 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10320
         TabIndex        =   22
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox class 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form3.frx":0014
         Left            =   10320
         List            =   "Form3.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Please select  class"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox stud_name 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   20
         Top             =   1920
         Width           =   2775
      End
      Begin VB.ComboBox blood_group 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form3.frx":0018
         Left            =   3120
         List            =   "Form3.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Uplaod Picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   15840
         TabIndex        =   18
         Top             =   3240
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Height          =   2175
         Left            =   15360
         ScaleHeight     =   2115
         ScaleWidth      =   3075
         TabIndex        =   17
         Top             =   960
         Width           =   3135
         Begin VB.Image Image1 
            Height          =   1935
            Left            =   120
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2775
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   8520
         TabIndex        =   13
         Top             =   2520
         Width           =   4335
         Begin VB.OptionButton Option1 
            Caption         =   "Male"
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
            Left            =   1800
            TabIndex        =   15
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Female"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   14
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Gender"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   8520
         TabIndex        =   9
         Top             =   3480
         Width           =   4215
         Begin VB.OptionButton Option3 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   11
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3120
            TabIndex        =   10
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Transport"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.ComboBox section 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form3.frx":001C
         Left            =   10320
         List            =   "Form3.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox trns_loc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form3.frx":0020
         Left            =   12000
         List            =   "Form3.frx":0022
         TabIndex        =   7
         Top             =   3840
         Width           =   3135
      End
      Begin VB.ComboBox stud_doc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form3.frx":0024
         Left            =   3120
         List            =   "Form3.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4440
         Width           =   3135
      End
      Begin VB.TextBox stud_doc_no 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   12000
         TabIndex        =   5
         Top             =   4320
         Width           =   3135
      End
      Begin VB.TextBox pincode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5760
         TabIndex        =   4
         Text            =   "Pincode"
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox city 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   3
         Text            =   "City"
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox stud_add 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   2
         Top             =   3480
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   3120
         TabIndex        =   25
         Top             =   2400
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
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
         CustomFormat    =   "dd- MMM-yyyy"
         Format          =   125108227
         CurrentDate     =   44636
      End
      Begin VB.Label serial_no 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
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
         Left            =   1680
         TabIndex        =   70
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   15480
         TabIndex        =   66
         Top             =   480
         Width           =   720
      End
      Begin VB.Label adm_date 
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
         Left            =   16560
         TabIndex        =   65
         Top             =   480
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serial No."
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
         Left            =   240
         TabIndex        =   64
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Relligion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8640
         TabIndex        =   40
         Top             =   3000
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Session"
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
         Left            =   480
         TabIndex        =   39
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Class"
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
         Left            =   8640
         TabIndex        =   38
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admission No."
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
         Left            =   480
         TabIndex        =   37
         Top             =   1560
         Width           =   1725
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   480
         TabIndex        =   36
         Top             =   2040
         Width           =   1740
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   480
         TabIndex        =   35
         Top             =   2640
         Width           =   1560
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Section"
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
         Left            =   8640
         TabIndex        =   34
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Roll No."
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
         Left            =   8640
         TabIndex        =   33
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Group"
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
         Left            =   480
         TabIndex        =   32
         Top             =   3120
         Width           =   1530
      End
      Begin VB.Label gender 
         Height          =   15
         Left            =   9000
         TabIndex        =   31
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
         WordWrap        =   -1  'True
      End
      Begin VB.Label transport 
         Height          =   15
         Left            =   9480
         TabIndex        =   30
         Top             =   3600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Pickup Location"
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
         Left            =   8640
         TabIndex        =   29
         Top             =   3960
         Width           =   2910
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documents"
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
         Index           =   1
         Left            =   480
         TabIndex        =   28
         Top             =   4560
         Width           =   1380
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document No."
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
         Index           =   1
         Left            =   8640
         TabIndex        =   27
         Top             =   4440
         Width           =   1725
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Index           =   1
         Left            =   480
         TabIndex        =   26
         Top             =   3600
         Width           =   1005
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SUBMIT AND PAY"
      Height          =   480
      Left            =   7440
      TabIndex        =   0
      Top             =   8880
      Width           =   3135
   End
End
Attribute VB_Name = "AdmissionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String

Private Sub admission_no_GotFocus()
admission_no.Text = Right(adm_date, 4) + "S" & serial_no.Caption + "C" & class.Text
End Sub

Private Sub city_Click()
city.Text = ""
End Sub

Private Sub Combo1_Click()
conn
sql = "select *from stud_admission where admission_no = '" & Combo1.Text & "'"
Set r = c.Execute(sql)
serial_no.Caption = r.Fields(0)
adm_date.Caption = r.Fields(1)
session.Text = r.Fields(2)
class.Text = r.Fields(3)
admission_no.Text = r.Fields(4)
section.Text = r.Fields(5)
stud_name.Text = r.Fields(6)
roll_no.Text = r.Fields(7)
DTPicker1.Value = r.Fields(8)
gender.Caption = r.Fields(9)
blood_group.Text = r.Fields(10)
relligion.Text = r.Fields(11)
stud_add.Text = r.Fields(12)
city.Text = r.Fields(13)
pincode.Text = r.Fields(14)
stud_doc.Text = r.Fields(14)
stud_doc_no.Text = r.Fields(14)
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
P = CommonDialog1.FileName
Image1.Picture = LoadPicture(P)
End Sub
Private Sub Command2_Click()
If session = "" Then
MsgBox "Enter a sessoion!"
session.SetFocus
ElseIf class = "" Then
ToolTipText = "slect class"
MsgBox "Enter class!"
class.SetFocus
ElseIf admission_no = "" Then
MsgBox "please Enter Admission number!"
admission_no.SetFocus
ElseIf section = "" Then
MsgBox "Enter a class Section"
ElseIf stud_name = "" Then
MsgBox "Enter Student name!"
stud_name.SetFocus
ElseIf roll_no = "" Then
MsgBox "please Roll Number number!"
roll_no.SetFocus
' ElseIf Frame1 = "" Then
' MsgBox "Please Select Gender!"
ElseIf fname = "" Then
MsgBox "Enter Father's Name"
fname.SetFocus
'ElseIf fqulification = "" Then
'MsgBox "Enter Father's Qulification!"
'fqulification.SetFocus
'ElseIf foccupation = "" Then
'MsgBox "Enter father's occupation!"
'foccupation.SetFocus
'ElseIf mname = "" Then
'MsgBox "Enter Mother's name!"
'mname.SetFocus
'ElseIf mqulification = "" Then
'MsgBox "Enter Mother's Qulification!"
'mqulification.SetFocus
'ElseIf moccupation = "" Then
'MsgBox "Enter Mother's occupation!"
'moccupation.SetFocus
'ElseIf relligion = "" Then
'MsgBox "Enter Relligion!"
'relligion.SetFocus
'ElseIf address = "" Then
'MsgBox "Enter Address!"
'address.SetFocus
'ElseIf city = "" Then
'MsgBox "Enter city name!"
'city.SetFocus
' ElseIf pincode = "" Then
' ToolTipText = "Enter Zip Code"
 MsgBox "Enter area Pin Code!"
 pincode.SetFocus
 ElseIf stud_doc = "" Then
 MsgBox "Select Student Document Type!"
 stud_doc.SetFocus
 ElseIf stud_doc_no = "" Then
 MsgBox "Enter student document Number!"
 stud_doc_no.SetFocus
ElseIf parents_doc = "" Then
MsgBox "Select Parents Document Type!"
parents_doc.SetFocus
ElseIf parents_doc_no = "" Then
MsgBox "Enter Parents document Number!"
parents_doc_no.SetFocus
ElseIf email = "" Then
MsgBox "Enter Email address!"
email.SetFocus
ElseIf phone_no = "" Then
MsgBox "Enter Phone Number!"
phone_no.SetFocus
Else
conn 'call function
sql = "insert into stud_admission values(" + serial_no + ", '" + adm_date + "', '" + session + "','" + class + "',  '" + admission_no + "','" + section + "', '" + stud_name + "', " + roll_no + ",'" + Format(DTPicker1, "dd MMM yyyy") + "','" + gender.Caption + "' , '" + blood_group + "','" + relligion + "', '" + stud_add + "', '" + city + "', " + pincode + ",'" + stud_doc + "', '" + stud_doc_no + "')"
Set r = c.Execute(sql)
sql = "insert into stud_transport values('" + admission_no + "','" + trns_loc + "')"
Set r = c.Execute(sql)
sql = "insert into stud_parents values('" + admission_no + "',' " + fname + " ', '" + fqualification + "','" + foccupation + "','" + mname + "','" + mqulification + "','" + moccupation + "','" + parents_doc + "','" + parents_doc_no + "','" + email + "'," + phone_no + ")"
Set r = c.Execute(sql)
MsgBox "record saved"
serial_no.Caption = serial_no.Caption + 1
session.Text = ""
admission_no.Text = ""
stud_name.Text = ""
city.Text = ""
pincode.Text = ""
'class.Text = ""
'section.Text = ""
'stud_doc.Text = ""
'blood_group.Text = ""
roll_no.Text = ""
stud_add.Text = ""
stud_doc_no.Text = ""
trns_loc.Text = ""
fname.Text = ""
parents_doc_no.Text = ""
mname.Text = ""
email.Text = ""
phone_no.Text = ""
End If
End Sub



Private Sub Command3_Click()
conn 'call function
sql = "insert into stud_admission values(" + serial_no + ", '" + adm_date + "', '" + session + "','" + class + "',  '" + admission_no + "','" + section + "', '" + stud_name + "', " + roll_no + ",'" + Format(DTPicker1, "dd MMM yyyy") + "','" + gender.Caption + "' , '" + blood_group + "','" + relligion + "', '" + stud_add + "', '" + city + "', " + pincode + ",'" + stud_doc + "', '" + stud_doc_no + "')"
Set r = c.Execute(sql)
sql = "insert into stud_transport values('" + admission_no + "','" + trns_loc + "')"
Set r = c.Execute(sql)
sql = "insert into stud_parents values('" + admission_no + "',' " + fname + " ', '" + fqualification + "','" + foccupation + "','" + mname + "','" + mqulification + "','" + moccupation + "','" + parents_doc + "','" + parents_doc_no + "','" + email + "'," + phone_no + ")"
Set r = c.Execute(sql)
MsgBox "RECORD UPDATED"
End Sub

Private Sub Command4_Click()
conn
sql = "select count(serial_no) from stud_admission"
Set r = c.Execute(sql)
serial_no.Caption = r.Fields(0) + 1
session.Text = ""
admission_no.Text = ""
stud_name.Text = ""
city.Text = ""
city.Text = "City"
pincode.Text = ""
pincode.Text = "Pincode"
'class.Text = ""
'section.Text = ""
'stud_doc.Text = ""
'blood_group.Text = ""
roll_no.Text = ""
stud_add.Text = ""
stud_doc_no.Text = ""
trns_loc.Text = ""
fname.Text = ""
parents_doc_no.Text = ""
mname.Text = ""
email.Text = ""
phone_no.Text = ""
End Sub

Private Sub fname_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Then
Else
MsgBox "Enter Valid Name"
End If
End Sub

Private Sub Form_Load()
                          
conn
sql = "select admission_no from stud_admission "
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
                        
                            
                            blood_group.AddItem "A+"
                            blood_group.AddItem "A-"
                            blood_group.AddItem "B+"
                            blood_group.AddItem "B-"
                            blood_group.AddItem "AB+"
                            blood_group.AddItem "AB-"
                            blood_group.AddItem "O+"
                            blood_group.AddItem "O-"
                            
                            trns_loc.AddItem "Barhi"
                            trns_loc.AddItem "Bazarsmiti"
                            trns_loc.AddItem "Bahiri"
                            trns_loc.AddItem "Belganj"
                            trns_loc.AddItem "Begahi"
                            trns_loc.AddItem "Bhatauliya"
                            trns_loc.AddItem "Birti Tola"
                            trns_loc.AddItem "Jamua"
                            trns_loc.AddItem "Marpa Tahir"
                            trns_loc.AddItem "Masha"
                            trns_loc.AddItem "Musachak"
                            trns_loc.AddItem "Nandwara"
                            trns_loc.AddItem "Pachtaki Ram"
                            trns_loc.AddItem "Pachtaki Jadu"
                            trns_loc.AddItem "Parsanui"
                            trns_loc.AddItem "Patahi Bazar"
                            trns_loc.AddItem "Railway Gumti"
                            trns_loc.AddItem "Sekhona"
                        
                            stud_doc.AddItem "Aadhar Card"
                            stud_doc.AddItem "School Leaving Certificate"
                            stud_doc.AddItem "DOB Certificate"
                            
                            fqualification.AddItem "Not literate"
                            fqualification.AddItem "Bellow Primary"
                            fqualification.AddItem "Primary"
                            fqualification.AddItem "Midle"
                            fqualification.AddItem "Matriculation"
                            fqualification.AddItem "Intermidiate"
                            fqualification.AddItem "Gradute"
                            fqualification.AddItem "Post Gradute And above"
                            
                            foccupation.AddItem "Artist"
                            foccupation.AddItem "Business Man"
                            foccupation.AddItem "Designer"
                            foccupation.AddItem "Driver"
                            foccupation.AddItem "Doctor"
                            foccupation.AddItem "Engineer"
                            foccupation.AddItem "Goverment Employee"
                            foccupation.AddItem "Lobour"
                            foccupation.AddItem "Private Employee"
                            foccupation.AddItem "Self Employee"
                            foccupation.AddItem "Social worker"
                            foccupation.AddItem "Teaching"
                            foccupation.AddItem "Other worker"
                            
                            mqualification.AddItem "Not literate"
                            mqualification.AddItem "Bellow Primary"
                            mqualification.AddItem "Primary"
                            mqualification.AddItem "Midle"
                            mqualification.AddItem "Matriculation"
                            mqualification.AddItem "Intermidiate"
                            mqualification.AddItem "Gradute"
                            mqualification.AddItem "Post Gradute And above"
                            
                            moccupation.AddItem "Artist"
                            moccupation.AddItem "Business Man"
                            moccupation.AddItem "Designer"
                            moccupation.AddItem "Doctor"
                            moccupation.AddItem "Engineer"
                            moccupation.AddItem "Goverment Employee"
                            moccupation.AddItem "Housewife"
                            moccupation.AddItem "Nurses"
                            moccupation.AddItem "Private Employee"
                            moccupation.AddItem "Self Employee"
                            moccupation.AddItem "Social worker"
                            moccupation.AddItem "Teaching"
                             
                            
                            
                           parents_doc.AddItem "Aadhaar card"
                           parents_doc.AddItem "Driving license"
                           parents_doc.AddItem "Indian passport"
                           parents_doc.AddItem "PAN card"
                           parents_doc.AddItem "Voter ID Card"
                           parents_doc.AddItem "PAN card"
                           parents_doc.AddItem "School leaving certificate"
                           
                                                   
                           

Label17.Visible = False
trns_loc.Visible = False
'address.ToolTipText = "Enter Address Line"
'stud_doc.ToolTipText = "Select Student Document Type"
'stud_doc_no.ToolTipText = "Enter Student Document Number"
parents_doc.ToolTipText = "Select Parents Document Type"
parents_doc_no.ToolTipText = "Enter parents Document number"
'city.ToolTipText = "Enter City"
'pincode.ToolTipText = "Enter Area Zip Code"
adm_date.Caption = Format(Date, "dd mmm yyyy")
conn
sql = "select count(serial_no) from stud_admission"
Set r = c.Execute(sql)
serial_no.Caption = r.Fields(0) + 1

sql = "select distinct class_name from class order by class_name"
Set r = c.Execute(sql)
Do While Not r.EOF
class.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select distinct section_name from section order by section_name"
Set r = c.Execute(sql)
Do While Not r.EOF
section.AddItem r.Fields(0)
r.MoveNext
Loop
session.TabIndex = 0
End Sub

Private Sub Frame1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)


Label17.Visible = False
Combo1.Visible = False
address.ToolTipText = "Enter Address Line"
stud_doc.ToolTipText = "Select Student Document Type"
stud_doc_no.ToolTipText = "Enter Student Document Number"
parents_doc.ToolTipText = "Select Parents Document Type"
parents_doc_no.ToolTipText = "Enter parents Document number"
city.ToolTipText = "Enter City"
pincode.ToolTipText = "Enter Area Zip Code"
ad_date.Caption = Format(Date, "dd mmm yyyy")
conn
sql = "select count(serial_no) from stud_admission"
Set r = c.Execute(sql)
serial_no.Caption = r.Fields(0) + 1
serial_no.Locked = True
session.TabIndex = 0
End Sub

Private Sub mname_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Then
Else
MsgBox "Enter Valid Name"
End If
End Sub

Private Sub Option1_Click()
gender.Caption = "Male"
End Sub
Private Sub Option2_Click()
gender.Caption = "Famale"
End Sub

Private Sub Option3_Click()
transport.Caption = "Yes"
If transport.Caption = "Yes" Then
Label17.Visible = True
trns_loc.Visible = True
End If
End Sub

Private Sub Option4_Click()
transport.Caption = "No"
If transport.Caption = "No" Then
Label17.Visible = False
trns_loc.Visible = False
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

Private Sub phone_no_LostFocus()
If Len(phone_no) <> 10 And phone_no <> "" Then
MsgBox "Please Input Valid mobile Number! Mobile Number Must be 10 Digits"
phone_no.SetFocus
End If
End Sub



Private Sub pincode_Click()
pincode.Text = ""
End Sub

Private Sub pincode_KeyPress(KeyAscii As Integer)
pincode.MaxLength = 6
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
MsgBox "Only Digits are Allowed"
End If
End Sub

Private Sub pincode_LostFocus()
If Len(pincode) <> 6 And pincode <> "" Then
MsgBox "Please Enter a valid Area Pincode"
pincode.Text = ""
pincode.SetFocus
End If
End Sub

Private Sub section_Click()
sql = "select count(roll_no) from stud_admission where section='" + section.Text + "'"
Set r = c.Execute(sql)
roll_no.Text = r.Fields(0) + 1
End Sub

Private Sub stud_name_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Then
Else
MsgBox "Enter Valid Name"
End If
End Sub

