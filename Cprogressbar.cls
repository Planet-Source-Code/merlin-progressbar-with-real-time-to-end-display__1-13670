VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Klasse:   CProgressbar
'Author:   Sigurd Pekrul
'Funktion: Progressbar mit realer Zeitberechnung

Public Event Error(ByVal ErrNr As Long, ErrDescription As String)
Public Event ChangeTimeToGo(Time As String)
Public Event ChangeValue(Value As Variant)
Public Event ChangePercentValue(Value As String)

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Type Size
        cx As Long
        cy As Long
End Type


Private Const iIntervall As Integer = 1

Public cPictureObjekt As Variant

Private cInitCheck As Boolean
Private ShowTime As Double
Private RunTime As Double
Private sOut As String

Private Sub Class_Initialize()
On Error GoTo err_

cInitCheck = False
  
err_:
If Err.Number <> 0 Then
    RaiseEvent Error(Err.Number, Err.Description)
End If
  
End Sub
Public Function cBar(Min As Variant, Max As Variant, Value As Variant)
On Error GoTo err_

RaiseEvent ChangeValue(Value)

If cInitCheck = False Then
    cInit
    cInitCheck = True
End If


Dim KopieMin, KopieMax, KopieValue

KopieValue = (Value - Min) + 1
KopieMax = (Max - Min) + 1

If KopieValue >= KopieMax Then cInitCheck = False


With cPictureObjekt
    vValue = KopieValue / KopieMax * .ScaleWidth
    cPictureObjekt.Line (-1, -1)-(vValue, .ScaleWidth), RGB(0, 0, 255), BF
End With

'_____P A T C H _________________________________________

Dim lOut As Long
Static tTextSize As Size


If sOut <> vbInt(KopieValue / KopieMax * 100) & " %" Then
    sOut = vbInt(KopieValue / KopieMax * 100) & " %"
    RaiseEvent ChangePercentValue(sOut)
    lOut = GetTextExtentPoint32(cPictureObjekt.hdc, sOut, Len(sOut), tTextSize)
End If


lOut = TextOut(cPictureObjekt.hdc, (vValue / 2) - (tTextSize.cx / 2), (cPictureObjekt.ScaleHeight / 2) - (tTextSize.cy / 2), sOut, Len(sOut))
'_____P A T C H _________________________________________


If Intervall Then RaiseEvent ChangeTimeToGo(CTimer(KopieValue, KopieMax))


err_:
If Err.Number <> 0 Then
    RaiseEvent Error(Err.Number, Err.Description)
End If

End Function
Private Function Intervall() As Boolean

On Error GoTo err_:
    
    If (Timer - ShowTime) >= iIntervall Then
        Intervall = True
    Else
        Intervall = False
    End If

err_:
If Err.Number <> 0 Then
    RaiseEvent Error(Err.Number, Err.Description)
End If

End Function
Private Function TimeToString(dSec As Double) As String

On Error GoTo err_

Dim s As String
Dim iCheck As Integer


    s = ""

    
    iCheck = 0
    
    If iCheck < 2 Then
        If dSec / 31536000 >= 1 Then
            s = s & LTrim(Str(Int(dSec / 31536000))) & " Jahr"
            s = IIf(Int(dSec / 31536000) = 1, s & " ", s & "e ")
            dSec = dSec Mod 31536000
            iCheck = iCheck + 1
        End If
    End If
    
    
    If iCheck < 2 Then
        If dSec / 86400 >= 1 Then
            s = s & LTrim(Str(Int(dSec / 86400))) & " Tag"
            s = IIf(Int(dSec / 86400) = 1, s & " ", s & "e ")
            dSec = dSec Mod 86400
            iCheck = iCheck + 1
        End If
    End If

    
    If iCheck < 2 Then
        If dSec / 3600 >= 1 Then
            s = s & LTrim(Str(Int(dSec / 3600))) & " Stunde"
            s = IIf(Int(dSec / 3600) = 1, s & " ", s & "n ")
            dSec = dSec Mod 3600
            iCheck = iCheck + 1
        End If
    End If


    If iCheck < 2 Then
        If dSec / 60 >= 1 Then
            s = s & LTrim(Str(Int(dSec / 60))) & " Minute"
            s = IIf(Int(dSec / 60) = 1, s & " ", s & "n ")
            dSec = dSec Mod 60
            iCheck = iCheck + 1
        End If
    End If

    If iCheck < 2 Then
        If dSec >= 1 Then
            s = s & LTrim(Str(Int(dSec))) & " Sekunde"
            s = IIf(Int(dSec) = 1, s & " ", s & "n")
            iCheck = iCheck + 1
        End If
    End If
    
    TimeToString = s

    

err_:
If Err.Number <> 0 Then
    RaiseEvent Error(Err.Number, Err.Description)
End If

End Function
Private Function CTimer(Value As Variant, Max As Variant) As String
On Error GoTo err_

Dim sEndTime As String, dEndTime As Double, bla As String

If (Timer - ShowTime) >= iIntervall Then
    dEndTime = (((Timer - RunTime) / Value) * Max) - (Timer - RunTime)
    sEndTime = TimeToString(dEndTime)
    bla = sEndTime  ' & "   [ " & Int(Value / Max * 100) & " % ] "
    
    ShowTime = Timer
    If dEndTime < 1 And dEndTime > 0 Then
        bla = ""
    End If
    
        
End If

CTimer = bla

err_:
If Err.Number <> 0 Then
    RaiseEvent Error(Err.Number, Err.Description)
End If

End Function

Private Function vbInt(x As Variant) As Integer
On Error GoTo err_


X1 = Int(x)
X1 = X1 + 0.5
If x > X1 Then
    vbInt = Int(x) + 1
Else
    vbInt = Int(x)
End If

err_:
If Err.Number <> 0 Then
    RaiseEvent Error(Err.Number, Err.Description)
End If

End Function

Private Sub cInit()
On Error GoTo err_

RunTime = Timer
ShowTime = Timer
   
With cPictureObjekt
    .AutoSize = True
    .AutoRedraw = True
    .BackColor = RGB(255, 255, 255)
    .ForeColor = RGB(255, 255, 255)
    .FontTransparent = True
    .ScaleMode = 3
End With

err_:
If Err.Number <> 0 Then
    RaiseEvent Error(Err.Number, Err.Description)
End If

End Sub
