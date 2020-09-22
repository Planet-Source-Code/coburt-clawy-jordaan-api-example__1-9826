VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API : Windows ""hwnd"" Example"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmSize 
      Caption         =   " Windows Size Contols : "
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   5415
      Begin VB.CommandButton cmdMaximize 
         Caption         =   "Maximize"
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdMinimize 
         Caption         =   "Minimize"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdNormal 
         Caption         =   "Normal"
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Windows : "
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5415
      Begin VB.CommandButton cmdListWindows 
         Caption         =   "List Windows"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   3720
         Width           =   5175
      End
      Begin VB.ListBox listWindows 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ================================================================================== '
' API Example : Getting all the windows captions and "hwnd"s and to Minimize, Normal
' and maximizing a windows using it's "hwnd". Without having to point the mouse at
' the window first.
' Note : This does note include child windows.
' ================================================================================== '

Private Sub cmdListWindows_Click()
' ---------------------------------------------------------------------------------- '
    Set TargetList = frmMain.listWindows
    TargetList.Clear
    EnumWindows AddressOf EnumWindowsProc, 0
' ---------------------------------------------------------------------------------- '
End Sub

Private Sub cmdMinimize_Click()
' ---------------------------------------------------------------------------------- '
    If listWindows.ListCount > 0 Then
        If listWindows.ListIndex > -1 Then
            ShowWindow listWindows.ItemData(listWindows.ListIndex), SW_Minimize
        End If
    End If
' ---------------------------------------------------------------------------------- '
End Sub

Private Sub cmdNormal_Click()
' ---------------------------------------------------------------------------------- '
    If listWindows.ListCount > 0 Then
        If listWindows.ListIndex > -1 Then
            ShowWindow listWindows.ItemData(listWindows.ListIndex), SW_Normal
        End If
    End If
' ---------------------------------------------------------------------------------- '
End Sub

Private Sub cmdMaximize_Click()
' ---------------------------------------------------------------------------------- '
    If listWindows.ListCount > 0 Then
        If listWindows.ListIndex > -1 Then
            ShowWindow listWindows.ItemData(listWindows.ListIndex), SW_Maximize
        End If
    End If
' ---------------------------------------------------------------------------------- '
End Sub
