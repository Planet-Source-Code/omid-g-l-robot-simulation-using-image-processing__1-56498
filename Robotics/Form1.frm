VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Load Map"
      Height          =   405
      Left            =   4005
      TabIndex        =   7
      Top             =   105
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map"
      Height          =   4140
      Left            =   60
      TabIndex        =   5
      Top             =   675
      Width           =   5760
      Begin VB.PictureBox Map 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H0080FF80&
         Height          =   4050
         Left            =   180
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   4050
         ScaleWidth      =   6030
         TabIndex        =   6
         Top             =   450
         Width           =   6030
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C000&
            FillStyle       =   0  'Solid
            Height          =   75
            Left            =   4050
            Shape           =   3  'Circle
            Top             =   45
            Width           =   75
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5115
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.bmp"
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2205
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   150
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   885
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   150
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start !!!"
      Height          =   390
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Start Y"
      Height          =   195
      Index           =   1
      Left            =   1650
      TabIndex        =   4
      Top             =   195
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Start X"
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   3
      Top             =   195
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim bk As Integer
Dim DST As Destination
Dim DST1 As Destination
Dim PS As Position
Dim d As Long
PS.X = CInt(Text1.Text) * 15
PS.Y = CInt(Text2.Text) * 15
 'Map.PSet (Map.Width - 30, 45)
 Xmain 3, PS
 If Not IFoundit Then Xmain 1, PS
 If Not IFoundit Then Xmain 2, PS
 If Not IFoundit Then Xmain 4, PS
End Sub



Private Sub Command2_Click()
Dim Fn As String
Fn = OpenDialog("Picture Files (*.bmp,*.gif)" + Chr$(0) + "*.bmp;*.dib;*.gif" + Chr$(0), Me)
If Fn <> "" Then
    Map.Picture = LoadPicture(Fn)
    Frame1.Width = Map.Width + 400
    Me.Width = Map.Width + 600
    Frame1.Height = Map.Height + 600
    Me.Height = Map.Height + 1700
    Text1.Text = (Map.Width / 15) - 2
    Text2.Text = 3
End If
End Sub

Private Sub Form_Load()
Text1.Text = (Map.Width / 15) - 2
Text2.Text = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Map_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Form1.Caption = "X= " & CStr(X / 15) & "  Y= " & CStr(Y / 15)
End Sub
