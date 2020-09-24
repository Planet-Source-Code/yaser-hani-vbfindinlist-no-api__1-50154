VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "vbFindInList"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmMain.frx":0000
      Top             =   2520
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   2640
      ScaleHeight     =   1755
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enter search criteria here"
      Top             =   120
      Width           =   2415
   End
   Begin VB.ListBox AnimeList 
      Height          =   1815
      ItemData        =   "frmMain.frx":0190
      Left            =   120
      List            =   "frmMain.frx":01A6
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by: Yaser Hani (aka tHe bEaSt)
'E-mail: info@solarwind.tk

'================================================================================
'This function searches for an item in a list and it retrieves multiple results
'and can find the target using only a partial string, also it's highly custmoizable
'and the code is easy to understand...I hope you like it.
'================================================================================
'vbString: The string you will search for
'vbList: The lisbox that will search in
'iStart: the starting position of the search
Function vbFind(vbString As String, vbList As ListBox, Optional iStart As Integer) As Boolean
    Dim vbWhere As String 'a variable to hold the search target
    vbFind = False 'intialization of the function
    If iStart < 0 Then iStart = 0 ' handling the optional argument
    If vbString <> "" Then 'makes sure that the function doesn't search for null
        For i = iStart To vbList.ListCount - 1 'start loop
            vbString = LCase(vbString) ' making the search NOT case sensitive
            vbWhere = LCase(vbList.List(i)) 'making the search NOT case sensitive
            If InStr(1, vbWhere, vbString) > 0 Then 'InStr returns 0 if the string id not found
                vbFind = True 'string found
                Picture1.Print (i) 'you can enter whatever action you like here
                vbList.Selected(i) = True 'you can enter whatever action you like here
                iStart = i + 1 ' this enables the funstion get multiple results by continueing from after last found result
            End If
        Next i
    End If
    'When string is not found or null
    If vbFind = False Then MsgBox "File not found, Please make sure that you've entered the file name correctly and try again.", vbCritical, "Error"
End Function

Private Sub cmdSearch_Click()
Picture1.Cls
For i = 0 To AnimeList.ListCount - 1
    AnimeList.Selected(i) = False
Next i
x = vbFind(txtSearch, AnimeList)
Me.Caption = x

End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "If you found this code useful pls. vote for me, if not then at least write some feedback...thanx.."
End Sub
