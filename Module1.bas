Attribute VB_Name = "Module1"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Function Wave_Vertical(target As PictureBox, source As PictureBox, strength As Long)

Dim pos As Long, t As Long

pos = 0

Do

t = t + 1

For i = 0 To source.ScaleHeight
If i = source.ScaleHeight Then pos = pos + 1

xp = Sin(t / 5) * strength

target.PSet (pos, i), source.Point(pos, i + xp)

Next i
Loop While pos < source.ScaleWidth

DoEvents
target.Refresh

End Function

Public Function Wave_Horizonal(target As PictureBox, source As PictureBox, strength As Long)

Dim row As Long, t As Long

row = 0

Do

t = t + 1

For i = 0 To source.ScaleWidth
If i = source.ScaleWidth Then row = row + 1

yp = Sin(t / 5) * strength

target.PSet (i, row), source.Point(i + yp, row)

Next i
Loop While row < source.ScaleHeight

DoEvents
target.Refresh

End Function

