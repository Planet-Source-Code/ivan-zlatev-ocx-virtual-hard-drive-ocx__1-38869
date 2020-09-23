VERSION 5.00
Begin VB.Form VHDD 
   BackColor       =   &H00FDD7BB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Hard Drive Active-X "
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin Project1.izVirtualHDD izVirtualHDD1 
      Left            =   2760
      Top             =   4680
      _extentx        =   847
      _extenty        =   847
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDD7BB&
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   3375
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Text            =   "Virtual HDD Letter"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Create 
         BackColor       =   &H00FFC175&
         Caption         =   "Create"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Delete 
         BackColor       =   &H00FFC175&
         Caption         =   "Delete"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Exit 
         BackColor       =   &H00FFC175&
         Caption         =   "Exit"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox path 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDD7BB&
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00C0E0FF&
         Height          =   3015
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "By Ivan Zlatev (programs@mail.bg)"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   5760
      Width           =   2655
   End
End
Attribute VB_Name = "VHDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Create_Click()
'Creates Virtual HDD
Me.izVirtualHDD1.VirtualDrive_Create Me.Combo1.Text, Me.path.Text
MsgBox "Done! Check 'My computer'.", vbInformation

End Sub

Private Sub Delete_Click()
'delets Virtual HDD

Me.izVirtualHDD1.VirtualDrive_Delete Combo1.Text

MsgBox "Done! Check 'My computer'.", vbInformation
End Sub

Private Sub Dir1_Change()
path.Text = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
' adds all letters from a: to z:

For i = 65 To (65 + 25)
Combo1.AddItem Chr(i) & ":"

Next i

Combo1.ListIndex = 0
End Sub
