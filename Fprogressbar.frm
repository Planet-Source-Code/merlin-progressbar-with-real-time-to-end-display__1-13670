VERSION 5.00
Begin VB.Form FProgressBar 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Berechne das Ergebnis..."
   ClientHeight    =   1320
   ClientLeft      =   4785
   ClientTop       =   6000
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   5715
         TabIndex        =   1
         Top             =   240
         Width           =   5775
      End
   End
End
Attribute VB_Name = "FProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    FProgressBar.Command1.Cancel = True
End Sub

