VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   " DiffAdd Version 1.01 written by Merlin"
   ClientHeight    =   1830
   ClientLeft      =   4215
   ClientTop       =   4485
   ClientWidth     =   6765
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6765
   StartUpPosition =   1  'Fenstermitte
   Begin VB.OptionButton Option2 
      Caption         =   "Formel   (schnell)"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Schleife (langsam)"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   360
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Berechne"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2520
      MaxLength       =   150
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0,000000000000000000000000000000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   150
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   " 1 Zahl eingeben: "
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   " 2 Zahl eingeben: "
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame5 
      Caption         =   " Berechnung per "
      Height          =   855
      Left            =   4680
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abbrechen"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Frame Frame3 
      Caption         =   " Ergebnis: "
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   4335
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   269
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents cBar As CProgressBar
Attribute cBar.VB_VarHelpID = -1

Private Sub cBar_ChangePercentValue(Value As String)
Form1.Caption = Value & " - DiffAdd Version 1.01 written by Merlin"
End Sub

Private Sub cBar_ChangeTimeToGo(Time As String)
    Form1.Frame4.Caption = "verbleibende Zeit: ca " & Time
End Sub

Private Sub cBar_Error(ByVal ErrNr As Long, ErrDescription As String)
    MsgBox "Exception: " & ErrNr & "  ---->  " & ErrDescription, vbExclamation
    Command2.Cancel = True
End Sub


Private Sub Command1_Click()

If Form1.Option1.Value = True Then
    Text3.Text = PastePoint(LTrim(Str(VBDiffAdd(Val(Text1.Text), Val(Text2.Text)))))
End If

If Form1.Option2.Value = True Then
    x = Val(Text1.Text)
    Y = Val(Text2.Text)
    
    'Bringe Werte in Format x<y
    If x > Y Then
        z = x: x = Y: Y = z
    End If
    
    Text3.Text = PastePoint(LTrim(Str(((x + Y) / 2) * ((Y - x) + 1))))

End If

End Sub

Private Sub Command2_Click()
    Command2.Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Text1_Change()

Text1.Text = RemoveChar(Text1.Text)
    
End Sub
Private Function PastePoint(sString As String) As String

If Len(sString) < 4 Then
    PastePoint = sString
    Exit Function
End If

If InStr(1, sString, "E") <> 0 Then
    PastePoint = sString
    Exit Function
End If

z = 0
For x = Len(sString) To 1 Step -1
    
    z = z + 1
    If z / 3 = Int(z / 3) And x <> 1 Then
        Y = "." & Mid(sString, x, 1) & Y
    Else
        Y = Mid(sString, x, 1) & Y
    End If

Next x

PastePoint = Y

End Function
Private Function RemoveChar(sString As String) As String

If sString = "" Then Exit Function

For x = 1 To Len(sString)

    cut = Mid(sString, x, 1)
    For Y = 48 To 57
        If Asc(cut) = Y Then
            RemoveChar = RemoveChar & cut
            Exit For
        End If
    Next Y
    
Next x


End Function

Private Sub Text2_Change()

Text2.Text = RemoveChar(Text2.Text)

End Sub
Private Function VBDiffAdd(x As Double, Y As Double) As Double
Set cBar = New CProgressBar
Set cBar.cPictureObjekt = Form1.Picture1

'Bringe Werte in Format x<y
If x > Y Then
    z = x: x = Y: Y = z
End If

If Y - x > 500 Then
    Form1.Option1.Enabled = False
    Form1.Option2.Enabled = False
    Form1.Text1.Enabled = False
    Form1.Text2.Enabled = False
    Form1.Text3.Visible = False
    Form1.Frame3.Visible = False
    Form1.Command1.Visible = False
    Form1.Frame4.Visible = True
    Form1.Picture1.Visible = True
    Form1.Command2.Visible = True
    Form1.Frame4.Caption = " Berechne das Ergebnis... "
End If

z = 0
For d1 = x To Y
    z = z + d1
    DoEvents
    
    If Y - x > 500 Then
        cBar.cBar x, Y, d1
    
        If Form1.Command2.Cancel = True Then
            z = 0
            Form1.Command2.Cancel = False
            Exit For
        End If
    End If

Next d1


If Y - x > 500 Then
    Form1.Text3.Visible = True
    Form1.Frame3.Visible = True
    Form1.Command1.Visible = True
    Form1.Frame4.Visible = False
    Form1.Picture1.Visible = False
    Form1.Command2.Visible = False
    Form1.Text1.Enabled = True
    Form1.Text2.Enabled = True
    Form1.Option1.Enabled = True
    Form1.Option2.Enabled = True
    Form1.Caption = " DiffAdd Version 1.01 written by Merlin"
End If

VBDiffAdd = z
End Function
