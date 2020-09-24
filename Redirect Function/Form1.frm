VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "---COMPILE CODE++++"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Message Box"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Memory"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Fastest way to call external STDCALL"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "Redirect Call By Vanja Fuckar,EMAIL:INGA@VIP.HR ----------->@2003"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim S As String
Dim C As String
Dim sLen As Long


S = "Vanja Fuckar"
sLen = Len(S) * 2

C = Space(sLen)


CopyMemory StrPtr(C), StrPtr(S), sLen

TC = GetTickCount - TC
MsgBox "Copy:" & vbCrLf & C, , "Info"

End Sub



Private Sub Command3_Click()
MessageBox_W 0, "By Vanja Fuckar", "Redirect Function", 0
End Sub


