VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Range("A1:G9999").ClearContents

Range("E1") = "�������"
Range("E2") = "�������� ������"
Range("E3") = "d1 �������"
Range("E4") = "d2 �������"
Range("E5") = "s �������"
Range("E6") = "L1R �� ������� �����"
Range("E7") = "L1L �� ������ �����"
Range("E8") = "L2R �� ������� �����"
Range("E9") = "L2L �� ������ �����"
Range("E10") = "r ����������"
Range("E11") = "�������� �� z"
Range("E12") = "����� �� x"


Range("F1") = 1
Range("F2") = "��������"
Range("F3") = "52.4"
Range("F4") = "58"
Range("F5") = "6.2"
Range("F6") = "35"
Range("F7") = "Z"
Range("F8") = "47"
Range("F9") = "Z"
Range("F10") = "1"
Range("F11") = "0.05"
Range("F12") = "0.2"


Range("H2") = "�������� �����-��"
Range("H3") = "s ���� �����������"
Range("H4") = "s ���� �����������"
Range("H5") = "�������� ����, �/���"
Range("H6") = "�������� ����, �/���"
Range("I2") = "��������"
End Sub

Private Sub CommandButton3_Click()
Range("A1:G9999").ClearContents

Range("E1") = "������� ��� �����"
Range("E2") = "�������� ������"
Range("E3") = "d1 ��� ����� (min)"
Range("E4") = "d2 ��� ����� (max)"
Range("E5") = "d3 �������"
Range("E6") = "d4 ����������"
Range("E7") = "LL �� ������ �����"
Range("E8") = "s ������ ��� �����"
Range("E9") = "s1 ������ ������ �������"
Range("E10") = "s2 ������ ������������_"
Range("E11") = "s3 ������ �������_"
Range("E12") = "����� �� x"
Range("E13") = "r ����� � ������ �������"


Range("F1") = 2
Range("F2") = "��������"
Range("F3") = "40.9"
Range("F4") = "41.92"
Range("F5") = "50"
Range("F6") = "60"
Range("F7") = "100"
Range("F8") = "40.1"
Range("F9") = "10"
Range("F10") = "18.1"
Range("F11") = "Z"
Range("F12") = "0.2"
Range("F13") = "1"


Range("H2") = "�������� �����-��"
Range("H3") = "s ���� �����������"
Range("H4") = "s ���� �����������"
Range("H5") = "�������� ����, �/���"
Range("H6") = "�������� ����, �/���"
Range("I2") = "��������"
End Sub

Private Sub CommandButton4_Click()
Range("A1:G9999").ClearContents

Range("E1") = "��������"
Range("E2") = "�������� ������"
Range("D2") = "��� ��-��"
Range("C2") = "��� ��-��"
Range("B2") = "� ��-��"
Range("F1") = 3
Range("F2") = "��������"
Range("E3") = "d"
Range("F3") = 76
Range("E4") = "D"
Range("F4") = 30
Range("E5") = "Z0"
Range("F5") = 0
Range("E6") = "��� �������"
Range("F6") = "��������"
Range("g6") = "(������)"
Range("E7") = "� �������?"
Range("F7") = "��"

End Sub


Private Sub CommandButton5_Click()
Range("C10") = "G code"
Range("D10") = "�"
Range("E10") = "X"
Range("F10") = "Z"
Range("G10") = "R"
Range("F11") = Range("F5")
Range("D11") = 0


UserForm1.Show
End Sub

Private Sub CommandButton2_Click()
If Range("F1") = 1 Then
    Kanavki
ElseIf Range("F1") = 2 Then
    PodGayku
ElseIf Range("F1") = 3 Then
    Protochka
End If
End Sub


Sub PodGayku()
Dim d1 As Double, d2 As Double, d3 As Double, d4 As Double, s As Double, s1 As Double, s2 As Double, s3 As Double, LL As Double, LR As Double, R As Double, xZazor As Double
Dim SrChern As Double, SrChist As Double, SpeedChern As Long, SpeedChist As Long
Dim i As Long, j As Long, temp As Double

Range("A1:A9999").ClearContents

'��������� �������� ����������
d1 = Range("F3")
d2 = Range("F4")
d3 = Range("F5")
d4 = Range("F6")
LL = Range("F7")
s = Range("F8")
xZazor = Range("F12")
R = Range("F13")
If IsNumeric(Range("F9")) Then s1 = Range("F9")
If IsNumeric(Range("F10")) Then s2 = Range("F10")
If IsNumeric(Range("F11")) Then s3 = Range("F11")
If s1 = 0 And s2 <> 0 And s3 <> 0 Then s1 = s - s2 - s3
If s2 = 0 And s1 <> 0 And s3 <> 0 Then s2 = s - s1 - s3
If s3 = 0 And s2 <> 0 And s1 <> 0 Then s3 = s - s2 - s1
LR = LL - s

SrChern = Range("I3")
SrChist = Range("I4")
SpeedChern = Range("I5")
SpeedChist = Range("I6")
'/��������� �������� ����������

If s - s1 - s2 <> s3 Then
MsgBox "������� ������� �������� �������", vbExclamation
Exit Sub
End If

'�������� ���������� �����
Worksheets("����2").Range("C5") = "(chern. kanavochn. " & SrChern & "mm)"
Worksheets("����2").Range("C7") = "G00 Z-" & Replace(LR + SrChern + 0.5, ",", ".")
Worksheets("����2").Range("C8") = "X" & Replace(d3 + 10, ",", ".")
Worksheets("����2").Range("C9") = "G96 S" & SpeedChern
Worksheets("����2").Range("C10") = "G01 X" & Replace(d3 + 1, ",", ".") & " F0.5 M08"
Worksheets("����2").Range("C11") = "G75 I0.1 X" & Replace(d2 + 0.5, ",", ".") & " W-" & Replace(s3 - SrChern + 1, ",", ".") & " K" & Replace(SrChern - 1, ",", ".") & ". F0.1"
Worksheets("����2").Range("C12") = "G01 W-" & Replace(SrChern - 1, ",", ".") & " F1."
Worksheets("����2").Range("C13") = "G75 I0.1 X" & Replace(d1 + 0.5, ",", ".") & " W-" & Replace(s2 - SrChern - 1, ",", ".") & " K" & Replace(SrChern - 1, ",", ".") & ". F0.1"
Worksheets("����2").Range("C14") = "G01 W-" & Replace(SrChern - 1, ",", ".") & " F1."
Worksheets("����2").Range("C15") = "G75 I0.1 X" & Replace(d2 + 0.5, ",", ".") & " W-" & Replace(s1 - SrChern + 1, ",", ".") & " K" & Replace(SrChern - 1, ",", ".") & ". F0.1"
'/�������� ���������� �����


'�������� ���������� �����
Worksheets("����2").Range("C24") = "(chist. kanavochn. " & SrChist & "mm)"
Worksheets("����2").Range("C26") = "G00 Z-" & Replace(LR + SrChist - R, ",", ".")
Worksheets("����2").Range("C27") = "X" & Replace(d3 + 10, ",", ".")
Worksheets("����2").Range("C28") = "G96 S" & SpeedChist
Worksheets("����2").Range("C29") = "G01 X" & Replace(d3 + 1, ",", ".") & " F0.5 M08"
Worksheets("����2").Range("C30") = "X" & Replace(d3 + xZazor, ",", ".") & " F0.08"
Worksheets("����2").Range("C31") = "W-" & Replace(R, ",", ".") & " R-" & Replace(R, ",", ".")
Worksheets("����2").Range("C32") = "X" & Replace(d2, ",", ".")
Worksheets("����2").Range("C33") = "W-" & Replace(s3, ",", ".") & " R-0.4"
Worksheets("����2").Range("C34") = "X" & Replace(d1, ",", ".")
Worksheets("����2").Range("C35") = "W-" & Replace(s2 - SrChist, ",", ".")
Worksheets("����2").Range("C36") = "G00 X" & Replace(d4 + 4, ",", ".")
Worksheets("����2").Range("C37") = "G01 W-" & Replace(s1 + R, ",", ".") & " F1."
Worksheets("����2").Range("C38") = "X" & Replace(d4 + xZazor, ",", ".") & " F0.08"
Worksheets("����2").Range("C39") = "W" & Replace(R, ",", ".") & " R-" & Replace(R, ",", ".")
Worksheets("����2").Range("C40") = "X" & Replace(d2, ",", ".")
Worksheets("����2").Range("C41") = "W" & Replace(s1, ",", ".") & " R-0.4"
Worksheets("����2").Range("C42") = "X" & Replace(d1, ",", ".")
Worksheets("����2").Range("C43") = "G00 X" & Replace(d3 + 2, ",", ".") & " M09"
'/�������� ���������� �����

Tochki

End Sub

Sub Kanavki()
Dim d1 As Double, d2 As Double, s As Double, L1R As Double, L1L As Double, L2R As Double, L2L As Double, R As Double
Dim SrChern As Double, SrChist As Double, SpeedChern As Long, SpeedChist As Long
Dim i As Long, j As Long, temp As Double, ZRazgon As Double, xZazor

Range("A1:A9999").ClearContents

'��������� �������� ����������
d1 = Range("F3")
d2 = Range("F4")
s = Range("F5")
If IsNumeric(Range("F6")) Then L1R = Range("F6")
If IsNumeric(Range("F7")) Then L1L = Range("F7")
If IsNumeric(Range("F8")) Then L2R = Range("F8")
If IsNumeric(Range("F9")) Then L2L = Range("F9")
If L1L = 0 And L1R <> 0 Then L1L = L1R + s
If L2L = 0 And L2R <> 0 Then L2L = L2R + s
If L1R = 0 And L1L <> 0 Then L1R = L1L - s
If L2R = 0 And L2L <> 0 Then L2R = L2L - s
R = Range("F10")
ZRazgon = Range("F11")
xZazor = Range("F12")

SrChern = Range("I3")
SrChist = Range("I4")
SpeedChern = Range("I5")
SpeedChist = Range("I6")
'/��������� �������� ����������

If Not L2L - s - L2R < 0.001 And Not L2L = 0 And Not L2R = 0 Or Not L1L - s - L1R < 0.001 Then
MsgBox "������� ������� �������� ������� �������", vbExclamation
Exit Sub
ElseIf s < SrChist Or s < SrChern Then
    MsgBox "����� �� ����� ���� ���� �������", vbExclamation
    Exit Sub
End If


'�������� ���������� �����
Worksheets("����2").Range("A5") = "(chern. kanavochn. " & SrChern & "mm)"
If s - SrChern > 1.2 Then '���� �������� ����� ������� ��� �������
    Worksheets("����2").Range("A7") = "G00 Z-" & Replace(L1R + SrChern + 0.3, ",", ".")
    Worksheets("����2").Range("A8") = "X" & Replace(d2 + 10, ",", ".")
    Worksheets("����2").Range("A9") = "G96 S" & SpeedChern
    Worksheets("����2").Range("A10") = "G01 X" & Replace(d2 + 1, ",", ".") & " F0.5 M08"
    Worksheets("����2").Range("A11") = "G75 I0.1 X" & Replace(d1 + 0.5, ",", ".") & " Z-" & Replace(L1L - 0.3, ",", ".") & " K" & Replace(SrChern - 1, ",", ".") & ". F0.08"
    If L2R <> 0 Then
        Worksheets("����2").Range("A12") = "G01 Z-" & Replace(L2R + SrChern + 0.3, ",", ".") & " F0.5"
        Worksheets("����2").Range("A13") = "G01 X" & Replace(d2 + 1, ",", ".") & " F0.5"
    Worksheets("����2").Range("A14") = "G75 I0.1 X" & Replace(d1 + 0.5, ",", ".") & " Z-" & Replace(L2L - 0.3, ",", ".") & " K" & Replace(SrChern - 1, ",", ".") & ". F0.08"
    Else
        Worksheets("����2").Range("A12") = "z"
        Worksheets("����2").Range("A13") = "z"
        Worksheets("����2").Range("A14") = "z"
    End If

Else
    Worksheets("����2").Range("A7") = "G00 Z-" & Replace(L1R + SrChern + (s - SrChern) / 2, ",", ".")
    Worksheets("����2").Range("A8") = "X" & Replace(d2 + 10, ",", ".")
    Worksheets("����2").Range("A9") = "G96 S" & SpeedChern
    Worksheets("����2").Range("A10") = "G01 X" & Replace(d2 + 1, ",", ".") & " F0.5 M08"
    Worksheets("����2").Range("A11") = "G75 I0.1 X" & Replace(d1 + 0.5, ",", ".") & " F0.08"
    If L2R <> 0 Then
        Worksheets("����2").Range("A12") = "G01 Z-" & Replace(L2R + SrChern + (s - SrChern) / 2, ",", ".") & " F0.5"
        Worksheets("����2").Range("A13") = "G01 X" & Replace(d2 + 1, ",", ".") & " F0.5"
        Worksheets("����2").Range("A14") = "G75 I0.1 X" & Replace(d1 + 0.5, ",", ".") & " F0.08"
    Else
        Worksheets("����2").Range("A12") = "z"
        Worksheets("����2").Range("A13") = "z"
        Worksheets("����2").Range("A14") = "z"
    End If
End If
'/�������� ���������� �����

'�������� ���������� �����
Worksheets("����2").Range("A23") = "(chist. kanavochn. " & SrChist & "mm)"
Worksheets("����2").Range("A25") = "G00 Z-" & Replace(L1L, ",", ".")
Worksheets("����2").Range("A26") = "X" & Replace(d2 + 10, ",", ".")
Worksheets("����2").Range("A27") = "G96 S" & SpeedChist
Worksheets("����2").Range("A28") = "G01 X" & Replace(d2 + 1, ",", ".") & " F0.5 M08"
Worksheets("����2").Range("A29") = "G01 U-" & Replace(1 - xZazor, ",", ".") & " F0.06"
Worksheets("����2").Range("A30") = "W-" & Replace(R, ",", ".")
Worksheets("����2").Range("A31") = "W" & Replace(R, ",", ".") & " R-" & Replace(R, ",", ".")
Worksheets("����2").Range("A32") = "U-" & Replace(d2 - d1 + xZazor, ",", ".")
temp = L1L - L1R - SrChern - (s - SrChern) / 2 + SrChern - SrChist '������� ����� ������ �� z ��������
If temp <= 0.5 Then
    temp = Round(temp / 2, 2)
Else
    temp = 0.5
End If
Worksheets("����2").Range("A33") = "U1. W" & Replace(temp, ",", ".")
Worksheets("����2").Range("A34") = "U" & Replace(d2 - d1 - 1 + xZazor, ",", ".") & " F0.5"
Worksheets("����2").Range("A35") = "W" & Replace(s + R - temp + ZRazgon - SrChist, ",", ".") & " F0.2"
Worksheets("����2").Range("A36") = "W-" & Replace(R, ",", ".") & " R-" & Replace(R, ",", ".") & " F0.06"
Worksheets("����2").Range("A37") = "U-" & Replace(d2 - d1 + xZazor, ",", ".")
Worksheets("����2").Range("A38") = "W-" & Replace(s - SrChist + ZRazgon, ",", ".")
Worksheets("����2").Range("A39") = "U1. W" & Replace(temp, ",", ".")
Worksheets("����2").Range("A40") = "U" & Replace(d2 - d1, ",", ".") & " F0.5"
If L2L <> 0 Then
    Worksheets("����2").Range("A41") = "G01 Z-" & Replace(L2L, ",", ".") & " F0.5"
    Worksheets("����2").Range("A42") = Worksheets("����2").Range("A29")
    Worksheets("����2").Range("A43") = Worksheets("����2").Range("A30")
    Worksheets("����2").Range("A44") = Worksheets("����2").Range("A31")
    Worksheets("����2").Range("A45") = Worksheets("����2").Range("A32")
    Worksheets("����2").Range("A46") = Worksheets("����2").Range("A33")
    Worksheets("����2").Range("A47") = Worksheets("����2").Range("A34")
    Worksheets("����2").Range("A48") = Worksheets("����2").Range("A35")
    Worksheets("����2").Range("A49") = Worksheets("����2").Range("A36")
    Worksheets("����2").Range("A50") = Worksheets("����2").Range("A37")
    Worksheets("����2").Range("A51") = Worksheets("����2").Range("A38")
    Worksheets("����2").Range("A52") = Worksheets("����2").Range("A39")
    Worksheets("����2").Range("A53") = Worksheets("����2").Range("A40")
Else
    Worksheets("����2").Range("A41") = "z"
    Worksheets("����2").Range("A42") = "z"
    Worksheets("����2").Range("A43") = "z"
    Worksheets("����2").Range("A44") = "z"
    Worksheets("����2").Range("A45") = "z"
    Worksheets("����2").Range("A46") = "z"
    Worksheets("����2").Range("A47") = "z"
    Worksheets("����2").Range("A48") = "z"
    Worksheets("����2").Range("A49") = "z"
    Worksheets("����2").Range("A50") = "z"
    Worksheets("����2").Range("A51") = "z"
    Worksheets("����2").Range("A52") = "z"
    Worksheets("����2").Range("A53") = "z"
End If
'/�������� ���������� �����

Tochki


End Sub
Sub Tochki() '����������� ����� � ����� ��������� �� ����1
Dim s As String, s1 As String, swDot As Boolean
Dim i As Long, j As Long, k As Long
Dim colName As String

If Worksheets("����1").Range("F1") = 1 Then
    colName = "A"
ElseIf Worksheets("����1").Range("F1") = 2 Then
    colName = "C"
End If

k = 1
For i = 2 To 9999 '�������� ������� ��������� �� �����2 � ���������� ����� ��� ����������. � ����� - ����� �� ����1
If Worksheets("����2").Range(colName & i) = "end" Then Exit Sub
If Worksheets("����2").Range(colName & i) Like "*[XYZWUIKR]*" Then 'Like "*[(X#)(Y#)(Z#)(W#)(U#)(I#)(K#)(R#)]*" Then
    s1 = Worksheets("����2").Range(colName & i)
    s = vbNullString
    swDot = False
    '���������� ����� � ������, ���� ����������
    For j = 1 To Len(s1)
        If swDot = True Then
            If Mid(s1, j, 1) = "." Then swDot = False
            If Mid(s1, j, 1) = " " And Mid(s1, j - 1, 1) Like ("#") Then
                s = s & "."
                swDot = False
            End If
        Else
            If InStr(1, "XYZWUIKR", Mid(s1, j, 1)) <> 0 Then
                If j = 1 Then
                    If Mid(s1, j + 1, 1) Like ("#") Or Mid(s1, j + 1, 1) Like ("-") And Mid(s1, j + 2, 1) Like ("#") Then swDot = True
                ElseIf Mid(s1, j - 1, 1) = " " And Mid(s1, j + 1, 1) Like ("#") Or Mid(s1, j - 1, 1) = " " And Mid(s1, j + 1, 1) Like ("-") And Mid(s1, j + 2, 1) Like ("#") Then
                        swDot = True
                End If
            End If
        End If
    s = s & Mid(s1, j, 1)
    Next
    If Len(s1) = j - 1 And swDot = True Then s = s & "."
    '/���������� ����� � ������, ���� ����������
End If

'����� ������� ��������� �� ����1
If Not Worksheets("����2").Range(colName & i) = "z" Then
    If s = vbNullString Then
        Worksheets("����1").Range("A" & k) = Worksheets("����2").Range(colName & i)
    Else
        Worksheets("����1").Range("A" & k) = s
        s = vbNullString
    End If
    k = k + 1
End If
'����� ������� ��������� �� ����1

Next



End Sub

Function pi()
pi = Application.pi()
End Function

Sub Protochka()
Dim x01, z01, x02, z02, ugol, r1, r2 As Double
Dim x1, x2, x3, x4, z1, z2, z3, z4 As Double

x01 = Range("F4").Value
z01 = Range("F5").Value
x02 = Range("F6").Value
z02 = Range("F7").Value
ugol = Range("F8").Value * pi() / 180
r1 = Range("F9").Value
r2 = Range("F10").Value

z2 = r1 - r1 * Sin(ugol)
x2 = x01 + (z2 * Tan(ugol)) * 2
x1 = x2 - (r1 * Cos(ugol)) * 2
z1 = z01

x3 = x02 - r2 * (1 - Cos(ugol)) * 2
z3 = (x3 - x01) / (2 * Tan(ugol))
x4 = x02
z4 = z3 + r2 * Sin(ugol)

x1 = Round(x1, 3)
x2 = Round(x2, 3)
x3 = Round(x3, 3)
x4 = Round(x4, 3)

z1 = -Round(z1, 3)
z2 = -Round(z2, 3)
z3 = -Round(z3, 3)
z4 = -Round(z4, 3)

Range("A1") = "G01 X" & Replace(x1, ",", ".") & " Z" & Replace(z1, ",", ".")
Range("A2") = "G03 X" & Replace(x2, ",", ".") & " Z" & Replace(z2, ",", ".") & " R" & Replace(r1, ",", ".")
Range("A3") = "G01 X" & Replace(x3, ",", ".") & " Z" & Replace(z3, ",", ".")
Range("A4") = "G03 X" & Replace(x4, ",", ".") & " Z" & Replace(z4, ",", ".") & " R" & Replace(r2, ",", ".")


End Sub


