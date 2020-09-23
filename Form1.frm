VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "PSC VB Ticker"
   ClientHeight    =   4905
   ClientLeft      =   3450
   ClientTop       =   1785
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   3285
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   225
      TabIndex        =   2
      Top             =   4455
      Width           =   2805
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   225
      TabIndex        =   1
      Top             =   4095
      Width           =   2805
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   3930
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3075
      ExtentX         =   5424
      ExtentY         =   6932
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()
wb.Refresh
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
wb.navigate App.Path & "/ticker.html"

End Sub
