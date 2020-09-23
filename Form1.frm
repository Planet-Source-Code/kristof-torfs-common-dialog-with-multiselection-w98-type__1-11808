VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "Common Dialog Example"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin Project1.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   2400
      ItemData        =   "Form1.frx":01DA
      Left            =   0
      List            =   "Form1.frx":01E1
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "You have opened these files:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public filenames As String   ' This must always be declared on the form

Public Sub Cdlg_Init()
    ' Even if nothing is done here, this function has to stay
End Sub

Public Sub Cdlg_UserAction(id As CtrlID, bCancel As Boolean, nAction As ActionType)
    If id = ID_FILETEXT And nAction = EN_CHANGE Then filenames = GetCtrlText(ID_FILETEXT)
    ' This must always be present on the ownerform
End Sub

Private Sub cmdClose_Click()
    Unload Me ' End the program
End Sub

Private Sub cmdOpen_Click()
    CommonDialog1.Filter = "All files (*.*)|*.*"    ' Find all files
    CommonDialog1.ShowOpenMultiSelect Me            ' Show the dialog
    If CommonDialog1.FileName = "" Then             ' Cancel was pressed
        MsgBox "You haven't opened any files!!!"
    Else                                            ' There is 1 or more files opened
        List1.Clear                                 ' Empty the list
        For i = LBound(Files) To UBound(Files)
            List1.AddItem Files(i)                  ' Add all files to the list
        Next i
        List1.Enabled = True                        ' Make items in list clickable
    End If
End Sub

Private Sub Form_Resize()
    If Me.Width < 2000 Then Me.Width = 4935
    List1.Width = Me.Width                          ' Adjust listbox width
    Me.Height = 3645                                ' Form height cannot change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ret As VbMsgBoxResult
    ret = MsgBox("Are you sure you want to leave ?", vbQuestion + vbYesNo + vbDefaultButton2, "Exit?")
    If ret = vbNo Then Cancel = True
End Sub
