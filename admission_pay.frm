VERSION 5.00
Begin VB.Form AdmissionPayForm 
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15165
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
   ScaleHeight     =   7350
   ScaleWidth      =   15165
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "PAY"
      Height          =   615
      Left            =   6000
      TabIndex        =   9
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   7560
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   4440
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   7560
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7560
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dues Amount"
      Height          =   360
      Left            =   4680
      TabIndex        =   5
      Top             =   4680
      Width           =   1905
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Amount"
      Height          =   360
      Left            =   4680
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   360
      Left            =   4680
      TabIndex        =   3
      Top             =   3000
      Width           =   645
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admission No."
      Height          =   360
      Left            =   4680
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADMISSION FEE PAYMENT"
      Height          =   360
      Left            =   4920
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "AdmissionPayForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Label3.Caption = AdmissionForm.admission_no.Text
End Sub

