VERSION 5.00
Begin VB.UserControl izVirtualHDD 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UserControl1.ctx":0000
   ScaleHeight     =   525
   ScaleWidth      =   540
   Begin VB.TextBox dir 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "UserControl1.ctx":0CCA
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "izVirtualHDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Virtual Hard Drive Active-X
' By Ivan Zlatev (programs@mail.bg)


Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Event VirtualDriveCreate(VirtualDrive_Create)
Event VirtualDriveDelete(VirtualDrive_Delete)




' Converts a long path into it's 8.5 short path equilvent

Private Function GetShortPath(ByVal sLongPath As String) As String


    Dim lLen As Long

        ' Setup the buffer for the API call
    GetShortPath = Space(1024)

        ' Call the API, strip away the unwanted characters and return
    lLen = GetShortPathName(sLongPath, GetShortPath, Len(GetShortPath))
    GetShortPath = Left(GetShortPath, lLen)
End Function
Function VirtualDrive_Create(ByVal Virtual_Hard_Drive_Letter As String, ByVal The_Directory_from_which_the_drive_will_be_created As String)

'We are converting the LongPath to Short using the API
'because dos does not support the longpathname

dir.Text = GetShortPath(The_Directory_from_which_the_drive_will_be_created)

Shell "Subst " & Virtual_Hard_Drive_Letter & " " & dir.Text, vbHide

End Function

Private Sub UserControl_Resize()

UserControl.Height = UserControl.Image1.Height
UserControl.Width = UserControl.Image1.Height

End Sub



Function VirtualDrive_Delete(ByVal Virtual_Hard_Drive_Letter_To_Delete As String)


Shell "Subst " & Virtual_Hard_Drive_Letter_To_Delete & " /d", vbHide

End Function

