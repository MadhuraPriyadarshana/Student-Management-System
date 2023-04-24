VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15300
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   MouseIcon       =   "Form1.frx":C922
   Picture         =   "Form1.frx":19244
   ScaleHeight     =   8235
   ScaleWidth      =   15300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18975
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   600
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc studentdb 
         Height          =   735
         Left            =   15960
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"Form1.frx":33E49
         OLEDBString     =   $"Form1.frx":33ED1
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select*from student_info"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtad 
         DataField       =   "Admission Date"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   33
         Top             =   8400
         Width           =   9015
      End
      Begin VB.PictureBox image1 
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2835
         ScaleWidth      =   3315
         TabIndex        =   31
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CommandButton uploadbtn 
         Caption         =   "upload"
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton updatebtn 
         Caption         =   "update"
         Height          =   495
         Left            =   2160
         TabIndex        =   28
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtm 
         DataField       =   "Mother/Father OR Guardian"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   27
         Top             =   7680
         Width           =   9015
      End
      Begin VB.TextBox txts 
         DataField       =   "Status"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   25
         Top             =   6960
         Width           =   9015
      End
      Begin VB.TextBox Text2 
         DataField       =   "Namewithinitials"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   23
         Top             =   1920
         Width           =   9015
      End
      Begin VB.CommandButton searchbtn 
         BackColor       =   &H8000000A&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   21
         Top             =   6360
         Width           =   1215
      End
      Begin VB.TextBox txtsearch 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   615
         Left            =   840
         TabIndex        =   20
         Top             =   5640
         Width           =   2895
      End
      Begin VB.CommandButton cancelbtn 
         BackColor       =   &H0000C0C0&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   16080
         TabIndex        =   18
         Top             =   9360
         Width           =   2055
      End
      Begin VB.CommandButton nxtbtn 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   13320
         TabIndex        =   17
         Top             =   9360
         Width           =   2055
      End
      Begin VB.CommandButton prevbtn 
         BackColor       =   &H0080C0FF&
         Caption         =   "Pevious"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10680
         MaskColor       =   &H0000C0C0&
         TabIndex        =   16
         Top             =   9360
         Width           =   2055
      End
      Begin VB.TextBox txtgend 
         DataField       =   "Gender"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   15
         Top             =   6240
         Width           =   9015
      End
      Begin VB.TextBox txtgrade 
         DataField       =   "Grade"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   14
         Top             =   5520
         Width           =   9015
      End
      Begin VB.TextBox txtphone 
         DataField       =   "Contact Number"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   13
         Top             =   4800
         Width           =   9015
      End
      Begin VB.TextBox txtbirth 
         DataField       =   "Bithday"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   12
         Top             =   4080
         Width           =   9015
      End
      Begin VB.TextBox txtaddress 
         DataField       =   "Address"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   11
         Top             =   3360
         Width           =   9015
      End
      Begin VB.TextBox txtroll 
         DataField       =   "Index"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   5
         Top             =   2640
         Width           =   9015
      End
      Begin VB.TextBox txtnwi 
         DataField       =   "Full Name"
         DataSource      =   "studentdb"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   3
         Top             =   1200
         Width           =   9015
      End
      Begin VB.Label txta 
         BackColor       =   &H0080FFFF&
         Caption         =   "Admission Date ---------------------"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   5040
         TabIndex        =   32
         Top             =   8400
         Width           =   4455
      End
      Begin VB.Label Label14 
         Caption         =   "Label12"
         Height          =   2895
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FFFF&
         Caption         =   "Mother/Father OR Guardian"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5040
         TabIndex        =   26
         Top             =   7680
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Status ----------------------------------------"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5160
         TabIndex        =   24
         Top             =   6960
         Width           =   4335
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FFFF&
         Caption         =   "Name with Initials ----------------"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   5160
         TabIndex        =   22
         Top             =   1920
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Index /             Name with Initials"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FFFF&
         Caption         =   "Gender --------------------------------------"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5160
         TabIndex        =   10
         Top             =   6240
         Width           =   4335
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FFFF&
         Caption         =   "Grade ------------------------------------------"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5040
         TabIndex        =   9
         Top             =   5520
         Width           =   4335
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "Contact Number --------------------"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5160
         TabIndex        =   8
         Top             =   4800
         Width           =   4335
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Date of Birth -------------------------"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5160
         TabIndex        =   7
         Top             =   4080
         Width           =   4335
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Address ------------------------------------"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5160
         TabIndex        =   6
         Top             =   3360
         Width           =   4335
      End
      Begin VB.Label label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Index Number ------------------------"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5160
         TabIndex        =   4
         Top             =   2640
         Width           =   4215
      End
      Begin VB.Label lable2 
         BackColor       =   &H0080FFFF&
         Caption         =   " Full Name ------------------------------"
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5160
         TabIndex        =   2
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daham School Student Data System"
         BeginProperty Font 
            Name            =   "Magneto"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   735
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   12375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String

Private Sub cancelbtn_Click()
txtroll.Text = ""
Text2.Text = ""
txtnwi.Text = ""
txtaddress.Text = ""
txtgend.Text = ""
txtphone = ""
txtbirth.Text = ""
txtgrade.Text = ""
txtm.Text = ""
txts.Text = ""

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
combo1.AddItem
End Sub

Private Sub nxtbtn_Click()
studentdb.Recordset.MoveNext
End Sub

Private Sub prevbtn_Click()
studentdb.Recordset.MovePrevious
End Sub



Private Sub Savebtn_Click()

End Sub

Private Sub searchbtn_Click()
k = txtsearch.Text
studentdb.RecordSource = ("select*from student_info where Index  = " + "'" + k + "'or namewithinitials = " + "'" + k + "'")
studentdb.Refresh
If studentdb.Recordset.EOF Then
MsgBox "record Not Found Try Another Index Number or Name", vbInformation
Else
studentdb.Caption = studentdb.RecordSource
End If
End Sub


Private Sub updatebtn_Click()

studentdb.Recordset("index").Value = txtroll.Text
studentdb.Recordset.Fields("Full name") = txtnwi.Text
studentdb.Recordset.Fields("address") = txtaddress.Text
studentdb.Recordset.Fields("index") = txtroll.Text
MsgBox "data is updated", vbInformation, "Message"
studentdb.Recordset.Update
End Sub



Private Sub uploadbtn_Click()
CommonDialog1.FileName = ""
CommonDialog1.Filter = "JPEG fILES|*.jpg|GIF Files|*.gif|All Files|*.*"
CommonDialog1.showopen
label4 = CommonDialog1.FileName
If Len(Trim(label4)) < 1 Then
Exit Sub
End If
image1.Picture = LoadPicture(label4)
End Sub
