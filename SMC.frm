VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SMC 
   Caption         =   "SMC for Excel"
   ClientHeight    =   10410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7080
   OleObjectBlob   =   "SMC.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



' FROM COPYPASTE MODULE
Option Explicit
Public PreviousCell As Range
Public PreviouseCellSheet As String

#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hwnd As LongPtr) As Long
#Else
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
#Else
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
#Else
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As LongPtr) As Long
#Else
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function GetClipboardData Lib "user32.dll" (ByVal wFormat As LongPtr) As Long
#Else
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As Long
#Else
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As Long
#Else
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
#Else
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
#Else
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
#Else
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As Long
#Else
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
#End If


' START AUTOMERGE MODULE
Dim mafChrWid(32 To 127) As Double ' widths of printing characters
Dim msFontName As String ' font name having these widths

Function StrWidth(s As String, sFontName As String, fFontSize As Double) As Double
 ' Returns the approximate width in points of a text string
 ' in a specified font name and font size
 ' Does not account for kerning
 Dim i As Long
 Dim j As Long
 If Len(sFontName) = 0 Then Exit Function
 If sFontName <> msFontName Then
 If Not InitChrWidths(sFontName) Then Exit Function
 End If
 For i = 1 To Len(s)
 j = Asc(Mid(s, i, 1))
 If j >= 32 Then
 StrWidth = StrWidth + fFontSize * mafChrWid(j)
 End If
 Next i
End Function
Function InitChrWidths(sFontName As String) As Boolean
 Dim i As Long
 Select Case sFontName
 Case "Arial"
 For i = 32 To 127
 Select Case i
 Case 39, 106, 108
 mafChrWid(i) = 0.1902
 Case 105, 116
 mafChrWid(i) = 0.2526
 Case 32, 33, 44, 46, 47, 58, 59, 73, 91 To 93, 102, 124
 mafChrWid(i) = 0.3144
 Case 34, 40, 41, 45, 96, 114, 123, 125
 mafChrWid(i) = 0.3768
 Case 42, 94, 118, 120
 mafChrWid(i) = 0.4392
 Case 107, 115, 122
 mafChrWid(i) = 0.501
 Case 35, 36, 48 To 57, 63, 74, 76, 84, 90, 95, 97 To 101, 103, 104, 110 To 113, 117, 121
  mafChrWid(i) = 0.5634
 Case 43, 60 To 62, 70, 126
 mafChrWid(i) = 0.6252
 Case 38, 65, 66, 69, 72, 75, 78, 80, 82, 83, 85, 86, 88, 89, 119
 mafChrWid(i) = 0.6876
 Case 67, 68, 71, 79, 81
 mafChrWid(i) = 0.7494
 Case 77, 109, 127
 mafChrWid(i) = 0.8118
 Case 37
 mafChrWid(i) = 0.936
 Case 64, 87
 mafChrWid(i) = 1.0602
 End Select
 Next i
 Case "Consolas"
 For i = 32 To 127
 Select Case i
 Case 32 To 127
 mafChrWid(i) = 0.5634
 End Select
 Next i
 Case "Calibri"
 For i = 32 To 127
 Select Case i
 Case 32, 39, 44, 46, 73, 105, 106, 108
 mafChrWid(i) = 0.2526
 Case 40, 41, 45, 58, 59, 74, 91, 93, 96, 102, 123, 125
 mafChrWid(i) = 0.3144
 Case 33, 114, 116
 mafChrWid(i) = 0.3768
 Case 34, 47, 76, 92, 99, 115, 120, 122
 mafChrWid(i) = 0.4392
 Case 35, 42, 43, 60 To 63, 69, 70, 83, 84, 89, 90, 94, 95, 97, 101, 103, 107, 118, 121, 124, 126
 mafChrWid(i) = 0.501
 Case 36, 48 To 57, 66, 67, 75, 80, 82, 88, 98, 100, 104, 110 To 113, 117, 127
 mafChrWid(i) = 0.5634
 Case 65, 68, 86
 mafChrWid(i) = 0.6252
 Case 71, 72, 78, 79, 81, 85
 mafChrWid(i) = 0.6876
 Case 37, 38, 119
 mafChrWid(i) = 0.7494
 Case 109
 mafChrWid(i) = 0.8742
 Case 64, 77, 87
 mafChrWid(i) = 0.936
 End Select
 Next i
 Case "Tahoma"
 For i = 32 To 127
 Select Case i
 Case 39, 105, 108
 mafChrWid(i) = 0.2526
 Case 32, 44, 46, 102, 106
 mafChrWid(i) = 0.3144
 Case 33, 45, 58, 59, 73, 114, 116
 mafChrWid(i) = 0.3768
 Case 34, 40, 41, 47, 74, 91 To 93, 124
 mafChrWid(i) = 0.4392
 Case 63, 76, 99, 107, 115, 118, 120 To 123, 125
 mafChrWid(i) = 0.501
 Case 36, 42, 48 To 57, 70, 80, 83, 95 To 98, 100, 101, 103, 104, 110 To 113, 117
 mafChrWid(i) = 0.5634
 Case 66, 67, 69, 75, 84, 86, 88, 89, 90
 mafChrWid(i) = 0.6252
 Case 38, 65, 71, 72, 78, 82, 85
 mafChrWid(i) = 0.6876
 Case 35, 43, 60 To 62, 68, 79, 81, 94, 126
 mafChrWid(i) = 0.7494
 Case 77, 119
 mafChrWid(i) = 0.8118
 Case 109
 mafChrWid(i) = 0.8742
 Case 64, 87
 mafChrWid(i) = 0.936
 Case 37, 127
 mafChrWid(i) = 1.0602
 End Select
 Next i
 Case "Lucida Console"
 For i = 32 To 127
 Select Case i
 Case 32 To 127
 mafChrWid(i) = 0.6252
 End Select
 Next i
 
 Case "Times New Roman"
 For i = 32 To 127
 Select Case i
 Case 39, 124
 mafChrWid(i) = 0.1902
 Case 32, 44, 46, 59
 mafChrWid(i) = 0.2526
 Case 33, 34, 47, 58, 73, 91 To 93, 105, 106, 108, 116
 mafChrWid(i) = 0.3144
 Case 40, 41, 45, 96, 102, 114
 mafChrWid(i) = 0.3768
 Case 63, 74, 97, 115, 118, 122
 mafChrWid(i) = 0.4392
 Case 94, 98 To 101, 103, 104, 107, 110, 112, 113, 117, 120, 121, 123, 125
 mafChrWid(i) = 0.501
 Case 35, 36, 42, 48 To 57, 70, 83, 84, 95, 111, 126
 mafChrWid(i) = 0.5634
 Case 43, 60 To 62, 69, 76, 80, 90
 mafChrWid(i) = 0.6252
 Case 65 To 67, 82, 86, 89, 119
 mafChrWid(i) = 0.6876
 Case 68, 71, 72, 75, 78, 79, 81, 85, 88
 mafChrWid(i) = 0.7494
 Case 38, 109, 127
 mafChrWid(i) = 0.8118
 Case 37
 mafChrWid(i) = 0.8742
 Case 64, 77
 mafChrWid(i) = 0.936
 Case 87
 mafChrWid(i) = 0.9984
 End Select
 Next i
 
 Case Else
 MsgBox "Font name """ & sFontName & """ not available!", vbCritical, "StrWidth"
 Exit Function
 End Select
 msFontName = sFontName
 InitChrWidths = True
End Function
' END AUTOMERGE MODULE

'RESUME COPYPASTE MODULE

Public Sub SetClipboard(sUniText As String)
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

Public Function GetClipboard() As String
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Dim sUniText As String
    Const CF_UNICODETEXT As Long = 13&
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
            lstrcpy StrPtr(sUniText), iLock
            GlobalUnlock iStrPtr
        End If
        GetClipboard = sUniText
    End If
    CloseClipboard
End Function
'END COPYPASTE MODULE

'APPLIED TO ALL SCRIPTS
Function makeNice()
'Selection.Merge
Selection.HorizontalAlignment = xlLeft
Selection.VerticalAlignment = xlVAlignCenter
Selection.Font.Name = "Arial"
Selection.Font.Size = 12
'Selection.Font.Color = _
'RGB(0, 0, 0)
'Selection.NumberFormat = 0
End Function

Function IfThereIsLetterInStr(x As String) As Boolean
If InStr(x, "A") = 0 _
And InStr(x, "B") = 0 _
And InStr(x, "C") = 0 _
And InStr(x, "D") = 0 _
And InStr(x, "E") = 0 _
And InStr(x, "F") = 0 _
And InStr(x, "G") = 0 _
And InStr(x, "H") = 0 _
And InStr(x, "I") = 0 _
And InStr(x, "J") = 0 _
And InStr(x, "K") = 0 _
And InStr(x, "L") = 0 _
And InStr(x, "M") = 0 _
And InStr(x, "N") = 0 _
And InStr(x, "O") = 0 _
And InStr(x, "P") = 0 _
And InStr(x, "Q") = 0 _
And InStr(x, "R") = 0 _
And InStr(x, "S") = 0 _
And InStr(x, "T") = 0 _
And InStr(x, "U") = 0 _
And InStr(x, "V") = 0 _
And InStr(x, "W") = 0 _
And InStr(x, "X") = 0 And InStr(x, "Y") = 0 And InStr(x, "Z") = 0 _
Then
    IfThereIsLetterInStr = False
    'MsgBox EndProgramIfThereIsLetterIn
Else
    IfThereIsLetterInStr = True
    'MsgBox EndProgramIfThereIsLetterIn
End If
End Function


'SHOULD BE APPLIED TO ALL SCRIPTS
Function autoMerge(z As Double, g As Double)
'z to accomodate kN/m*m/m etc format length
'g to round the value to get correct length instead of 0.333333 for example
On Error GoTo eh


Dim counter_a, counter_b, t, u, j, i, cells As Double
Dim total As Double
Dim a, b As Boolean
Dim x, yy, nn As Double

'get the active cell length

t = Range(ActiveCell.Address).Column
j = ActiveCell.Width
cells = 1
If ActiveCell.MergeCells = True Then
    ActiveCell.Offset(0, 1).Activate
    u = Range(ActiveCell.Address).Column
    cells = u - t
'''''''''MsgBox ("merged cells " & cells)
    ActiveCell.Offset(0, -1).Activate
End If
total = j * cells


'MsgBox (ActiveCell.Value2)
'MsgBox (StrWidth(Round(ActiveCell.Value2, g), "Arial", 12))
x = StrWidth(Round(ActiveCell.Value2, g), "Arial", 12) + z + 1  'text length and is cell boundary

'''''''''MsgBox (cells & " cells with total legth = " & total & ", text length = " & X & " (increase Z value if longer X required)")

If total > x And total < x + j Then
    ActiveCell = ActiveCell.Formula
ElseIf total > x Then
    ActiveCell = ActiveCell.Formula 'leave as it is
    ActiveCell.UnMerge ' trying to make cell smaller
    yy = Round(x / ActiveCell.Width + 0.5)
    ActiveCell.Resize(1, yy).Merge
    If yy <> 1 Then 'for one digit issues
        Range(ActiveCell.Offset(0, yy - 1), ActiveCell.Offset(0, cells - yy)).Select
    Else
        Range(ActiveCell.Offset(0, 1), ActiveCell.Offset(0, cells - 1)).Select
    End If
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Offset(0, -1).Activate
    
    
    
    
    
Else
    'get how many cells we need to merge
    yy = Round(x / ActiveCell.Width + 0.5) '0.5 is a trick because VBA has no roundup function
'''''''''MsgBox ("cells required " & yy)
        
        'now check is there anything on the right interrupting merge
    counter_a = 1
    For i = 1 To yy Step 1
        If ActiveCell.Offset(0, counter_a).Formula = "" Then
            counter_a = counter_a + 1
        End If
    Next
        
'''''''''MsgBox ("counter_a to the right to match cells required = " & counter_a)
       
    If counter_a >= yy And counter_a > 1 Then
''MsgBox ("Let's go to the right")
        ActiveCell.UnMerge 'to avoid excessive cell merging
        Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(0, yy - 1)).Merge 'merge cells


'>>>>>>>>>>>>>>>>
    Else
''MsgBox ("Let's go to the left")
           
        counter_b = -yy
''MsgBox ("counter_b end " & counter_b)
''MsgBox ("t = " & t)
        If -counter_b >= t Then
            counter_b = -t
            nn = t - 2 'small cell in the beginning
''MsgBox ("counter_b end " & counter_b)
''MsgBox ("yy = " & yy)
        Else
            nn = yy
        End If

''MsgBox ("counter_b to the left match cells required = " & counter_b)
        If nn > 1 Then
            For i = -yy To -1 Step 1
                If ActiveCell.Offset(0, counter_b).Formula = "" Then
                counter_b = counter_b + 1
            End If
            Next
        Else
            counter_b = -t
        End If
''MsgBox ("counter_b end " & counter_b)

        If counter_b > -yy And counter_b < 0 Or counter_b = -yy Then
''MsgBox ("not enough space to merge")
'''''''''MsgBox (yy & "-" & cells & "=" & yy - cells & "required")
            Range(ActiveCell.Offset(0, 1).Address, ActiveCell.Offset(0, yy - cells).Address).Insert xlShiftToRight

'''''''''MsgBox ("go back to the beginning")
'''''''''MsgBox ("merge" & yy - cells + 1 & "cells")
            Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(0, yy - cells)).Merge
        Else
            ActiveCell.UnMerge 'to avoid excessive cell merging
            ActiveCell.Offset(0, 1 - yy + cells - 1 + counter_a - 1).Activate
                                                '+cells -1 to cath the situation when cells are merged
                                                '+counter_a -1 to catch situations where there is still some space on the right
            Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(0, yy - 1)).Merge 'merge cells
        End If
'<<<<<<<<<<<<<<<<<<<<<<<<<
        

    End If
'''''MsgBox ("end merge")
End If


Done:
    Exit Function
eh:
    MsgBox ("Can't automerge")
    End
End Function

Function GetUnitsFromString(unitz As String)
'MsgBox ("start GetUnitsFromString")
unitz = UCase(unitz) 'to avoid caps issue
If InStr(unitz, "KN/M") > 0 Then
    CommandButton_kNperm_Click
ElseIf InStr(unitz, "KNM") > 0 Then
    CommandButton_kNm_Click
ElseIf InStr(unitz, "KN") > 0 Then
    CommandButton_kN_Click
ElseIf InStr(unitz, "KPA") > 0 Then
    CommandButton_kPa_Click
ElseIf InStr(unitz, "MPA") > 0 Then
    CommandButton_MPa_Click
ElseIf InStr(unitz, "GPA") > 0 Then
    CommandButton_GPa_Click
ElseIf InStr(unitz, "M2") > 0 Then
    CommandButton_m2_Click
ElseIf InStr(unitz, "MM") > 0 Then
    CommandButton_0mm_Click
ElseIf InStr(unitz, "M") > 0 Then
    CommandButton_m_Click
Else
    CommandButton_0super_Click
End If

End Function




Private Sub ComplexCalculation_Click()
On Error GoTo eh

Dim m As Double
Dim x, donothing, unitz, vv As String
Dim acell, cell, cellx, startcell As Range
Dim Counter, ctr, DollarCount_kN, DollarCount_kNm, DollarCount_m, DollarCount_kPa, DollarCount_kNperm, DollarCount_m2 As Integer
Dim ccheck, double_element_equation  As Boolean


'start building equiation
x = "="
'set where to write the equiation because R[]C[] format needs that first
If ActiveCell.Value = "" And ActiveCell.Offset(0, -1) <> "=" Then
    ActiveCell.Formula = "="
    makeNice
    ActiveCell.Offset(0, 1).Activate
ElseIf ActiveCell.Value = "=" Then
    ActiveCell.Offset(0, 1).Activate
End If

Set startcell = ActiveCell
Set acell = ActiveCell.Offset(0, -1)

'check the latest number first
emptycellfound: 'need after empty cell check down below

If acell.Column = 1 Then GoTo edasi

If acell.MergeCells = True Then

    Do Until acell.MergeCells = False
        Set acell = acell.Offset(0, -1)
    Loop
    Set acell = acell.Offset(0, 1)
Else

    'check if there any empty cells (which should NOT be here thought)
    If acell.Value = "" Then 'look for an empty cell
''''''''MsgBox ("emptycell is found")
        Set acell = acell.Offset(0, -1)
        GoTo emptycellfound
    End If
    'check if there any non digital symbols
''''''MsgBox ("now check for a letter")
    If Asc(acell.Value) > 47 And Asc(acell.Value) < 58 Then 'it's a number
        donothing = "do nothing"
    Else
''''''''MsgBox ("letter is found")
        Set acell = acell.Offset(0, -1)
        GoTo emptycellfound
    End If

End If

edasi:

'how many cells to the left we need to offset
m = Range(ActiveCell.Address).Column - 1
Range(ActiveCell, ActiveCell.Offset(0, -m)).Select

For Each cell In Selection
    
    If cell.Value = "/" Then
        Set cellx = cell.Offset(0, -1)
        If cellx.MergeCells = True Then
            Do Until cellx.MergeCells = False
                Set cellx = cellx.Offset(0, -1)
            Loop
            'x = x & cellx.Offset(0, 1).Value & "/"
            x = x & "R[0]C[" & -startcell.Column + cellx.Column + 1 & "]" & "/"
        Else
            'x = x & cell.Offset(0, -1).Value & "/"
            x = x & "R[0]C[" & -startcell.Column + cellx.Column & "]" & "/"
            
        End If

    
     ElseIf cell.Value = "-" Then
        Set cellx = cell.Offset(0, -1)
        If cellx.MergeCells = True Then
            Do Until cellx.MergeCells = False
                Set cellx = cellx.Offset(0, -1)
            Loop
            x = x & "R[0]C[" & -startcell.Column + cellx.Column + 1 & "]" & "-"
            unitz = unitz + cellx.Offset(0, 1).NumberFormat 'guess units for simple situations like 1kN+1kN+1kN
            Counter = Counter + 1
        Else
            x = x & "R[0]C[" & -startcell.Column + cellx.Column & "]" & "-"
            unitz = unitz + cellx.Offset(0, 1).NumberFormat
            Counter = Counter + 1
        End If
        
    
    ElseIf cell.Value = "+" Then
        Set cellx = cell.Offset(0, -1)
        If cellx.MergeCells = True Then
            Do Until cellx.MergeCells = False
                Set cellx = cellx.Offset(0, -1)
            Loop
            x = x & "R[0]C[" & -startcell.Column + cellx.Column + 1 & "]" & "+"
            unitz = unitz + cellx.Offset(0, 1).NumberFormat 'guess units for simple situations like 1kN-1kN-1kN
            Counter = Counter + 1
        Else
            x = x & "R[0]C[" & -startcell.Column + cellx.Column & "]" & "+"
            unitz = unitz + cellx.NumberFormat
            Counter = Counter + 1
            
        End If
        
    ElseIf cell.Value = "×" Then
        Set cellx = cell.Offset(0, -1)
        If cellx.MergeCells = True Then
            Do Until cellx.MergeCells = False
                Set cellx = cellx.Offset(0, -1)
            Loop
            x = x & "R[0]C[" & -startcell.Column + cellx.Column + 1 & "]" & "×"
        Else
            x = x & "R[0]C[" & -startcell.Column + cellx.Column & "]" & "×"
        End If
    
    End If

'meas = meas + cellx.NumberFormat
'MsgBox ("meas = " & meas)
'MsgBox (cellx.NumberFormat)

Next cell

'LAST
'x = x & acell.Value
x = x & "R[0]C[" & -startcell.Column + acell.Column & "]"
unitz = unitz + acell.NumberFormat
Counter = Counter + 1

startcell.FormulaR1C1 = x
startcell.Select

'> START guess units for simple situations like 1kN+1kN+1kN

'check if we can send it back to the main module with 2 elements
vv = ActiveCell.Formula
check_restart:
If InStr(vv, "*") > 1 Then
    ctr = ctr + 1
    vv = Replace(vv, "*", "", , 1)
    GoTo check_restart
ElseIf InStr(vv, "/") > 1 Then
    ctr = ctr + 1
    vv = Replace(vv, "/", "", , 1)
    GoTo check_restart
ElseIf InStr(vv, "+") > 1 Then
    ctr = ctr + 1
    vv = Replace(vv, "+", "", , 1)
    GoTo check_restart
ElseIf InStr(vv, "-") > 1 Then
    ctr = ctr + 1
    vv = Replace(vv, "-", "", , 1)
    GoTo check_restart
End If
      
If ctr = 1 Then
    double_element_equation = True
End If


If double_element_equation <> True Then
    'Count how many occurrences there are
    DollarCount_kN = (Len(unitz) - Len(Replace(unitz, "kN" & Chr(34), ""))) / Len("kN" & Chr(34))
    DollarCount_kNm = (Len(unitz) - Len(Replace(unitz, "kNm" & Chr(34), ""))) / Len("kNm" & Chr(34))
    DollarCount_m = (Len(unitz) - Len(Replace(unitz, " m" & Chr(34), ""))) / Len(" m" & Chr(34))
    DollarCount_kPa = (Len(unitz) - Len(Replace(unitz, "kPa" & Chr(34), ""))) / Len("kPa" & Chr(34))
    DollarCount_kNperm = (Len(unitz) - Len(Replace(unitz, "kN/m" & Chr(34), ""))) / Len("kN/m" & Chr(34))
    DollarCount_m2 = (Len(unitz) - Len(Replace(unitz, "m²" & Chr(34), ""))) / Len("m²" & Chr(34))
    
    'MsgBox ("unitz = " & unitz & " dollar count = " & DollarCount_kN & " counter = " & counter)
    
    If DollarCount_kN = Counter Then
        Call CommandButton_kN_Click
    End If
    
    If DollarCount_kNm = Counter Then
        Call CommandButton_kNm_Click
    End If
    
    If DollarCount_m = Counter Then
        Call CommandButton_m_Click
    End If
    
    If DollarCount_kPa = Counter Then
        Call CommandButton_kPa_Click
    End If
    
    If DollarCount_kNperm = Counter Then
        Call CommandButton_kNperm_Click
    End If
    
    If DollarCount_m2 = Counter Then
        Call CommandButton_m2_Click
    End If
    '>END guess
End If

makeNice
AppActivate Application.Caption


Done:
    Exit Sub

eh: MsgBox ("ComplexCalculation sub failed")
End

End Sub



Private Sub CommandButton_0_bars_Click()
makeNice
Selection.NumberFormat = "0"" bars"""
autoMerge 30, 0

AppActivate Application.Caption
End Sub

Private Sub CommandButton_000106mm4_Click()
makeNice
Selection.NumberFormat = "0.00"" × 10^6 mm""" & ChrW(&H2074)
autoMerge 90, 2

AppActivate Application.Caption


'm unicode is not working in numberFormat
'ActiveCell.FormulaR1C1 = ChrW(109)
End Sub

Private Sub CommandButton_0_Click()
On Error GoTo eh
Dim x, y As String

'makeNice not working

'x = Selection.NumberFormat
'MsgBox (x)
'y = Replace(x, "0.00", "0")
'MsgBox (y)
'y = Replace(y, "0.0", "0")
'MsgBox (y)

'Selection.NumberFormat = y


If Len(ActiveCell.Formula) > 0 Then
    If Left(ActiveCell.Formula, 1) = "=" Then
        If OptionButton_round = True Then
            ActiveCell.Formula = "=round(" & Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1) & ",0)"
        Else
            If OptionButton_up = True Then
                ActiveCell.Formula = "=roundup(" & Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1) & ",0)"
            Else
                If OptionButton_down = True Then
                    ActiveCell.Formula = "=rounddown(" & Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1) & ",0)"
                End If
            End If
        End If
    End If
End If

Done:
    Exit Sub
eh:
    MsgBox ("check code")
    End

End Sub

Private Sub CommandButton_00_Click()
On Error GoTo eh
Dim x, y, z As String
'makeNice not working

x = Selection.NumberFormat

If InStr(x, "0.0" & Chr(34)) > 0 Then
    z = "do nothing"
ElseIf InStr(x, "0.00" & Chr(34)) > 0 Then
    y = Replace(x, "0.00" & Chr(34), "0.0" & Chr(34))

Else
    y = Replace(x, "0" & Chr(34), "0.0" & Chr(34))

End If


Selection.NumberFormat = y

If Len(ActiveCell.Formula) > 0 Then
    If Left(ActiveCell.Formula, 1) = "=" Then
        If OptionButton_round = True Then
            ActiveCell.Formula = "=round(" & Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1) & ",1)"
        Else
            If OptionButton_up = True Then
                ActiveCell.Formula = "=roundup(" & Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1) & ",1)"
            Else
                If OptionButton_down = True Then
                    ActiveCell.Formula = "=rounddown(" & Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1) & ",1)"
                End If
            End If
        End If
    End If
End If

Done:
    Exit Sub
eh:
    MsgBox ("check code")
    End
End Sub

Private Sub CommandButton_000_Click()
On Error GoTo eh
Dim x, y, z As String
'makeNice not working

x = Selection.NumberFormat

If InStr(x, "0.00" & Chr(34)) > 0 Then
    z = "do nothing"
ElseIf InStr(x, "0.0" & Chr(34)) > 0 Then
    y = Replace(x, "0.0" & Chr(34), "0.00" & Chr(34))

Else
    y = Replace(x, "0" & Chr(34), "0.00" & Chr(34))

End If

Selection.NumberFormat = y

If Len(ActiveCell.Formula) > 0 Then
    If Left(ActiveCell.Formula, 1) = "=" Then
        If OptionButton_round = True Then
            ActiveCell.Formula = "=round(" & Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1) & ",2)"
        Else
            If OptionButton_up = True Then
                ActiveCell.Formula = "=roundup(" & Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1) & ",2)"
            Else
                If OptionButton_down = True Then
                    ActiveCell.Formula = "=rounddown(" & Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1) & ",2)"
                End If
            End If
        End If
    End If
End If

Done:
    Exit Sub
eh:
    MsgBox ("check code")
    End
End Sub

Private Sub CommandButton_000GPa_Click()
makeNice
Selection.NumberFormat = "0.00"" GPa"""
autoMerge 30, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_000kN_Click()
makeNice
Selection.NumberFormat = "0.00"" kN"""
autoMerge 40, 2

AppActivate Application.Caption

End Sub

Private Sub CommandButton_000kNm_Click()
makeNice
Selection.NumberFormat = "0.00"" kNm"""
autoMerge 38, 2

AppActivate Application.Caption

End Sub

Private Sub CommandButton_000knperm_Click()
makeNice
Selection.NumberFormat = "0.00"" kN/m"""
autoMerge 42, 2 'tested with 9.13

AppActivate Application.Caption

End Sub

Private Sub CommandButton_000kPa_Click()
makeNice
Selection.NumberFormat = "0.00"" kPa"""
autoMerge 38, 2

AppActivate Application.Caption

End Sub

Private Sub CommandButton_000m_Click()
makeNice
Selection.NumberFormat = "0.00"" m"""
autoMerge 10, 2 'checked on 0.15

AppActivate Application.Caption

End Sub

Private Sub CommandButton_000m2()
makeNice
Selection.NumberFormat = "0.00"" m²"""
autoMerge 30, 2

AppActivate Application.Caption

End Sub

Private Sub CommandButton_000MPa_Click()
makeNice
Selection.NumberFormat = "0.00"" MPa"""
autoMerge 38, 2

AppActivate Application.Caption

End Sub

Private Sub CommandButton_00GPa_Click()
makeNice
Selection.NumberFormat = "0.0"" GPa"""
autoMerge 40, 0 'checked on 1.0 10.0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_00kNm_Click()
makeNice
Selection.NumberFormat = "0.0"" kNm"""
autoMerge 40, 1

AppActivate Application.Caption

End Sub

Private Sub CommandButton_00knperm_Click()
makeNice
Selection.NumberFormat = "0.0"" kN/m"""
autoMerge 40, 1

AppActivate Application.Caption

End Sub

Private Sub CommandButton_00kPa_Click()
makeNice
Selection.NumberFormat = "0.0"" kPa"""
autoMerge 38, 1

AppActivate Application.Caption

End Sub

Private Sub CommandButton_00MPa_Click()
makeNice
Selection.NumberFormat = "0.0"" MPa"""
autoMerge 38, 1

AppActivate Application.Caption
End Sub

Private Sub CommandButton_00s_Click()
makeNice
Selection.NumberFormat = "0.0"" s"""
autoMerge 16, 1

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0deg_Click()
makeNice
Selection.NumberFormat = "0""°"""
autoMerge 14, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0GPa_Click()
makeNice
Selection.NumberFormat = "0"" GPa"""
autoMerge 31, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0knperm_Click()
makeNice
Selection.NumberFormat = "0"" kN/m"""
autoMerge 30, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0m2_Click()
makeNice
Selection.NumberFormat = "0"" m²"""
autoMerge 17, 0 'checked on 1150

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0mm_Click()
makeNice
Selection.NumberFormat = "0"" mm"""
autoMerge 25, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0mmcrs_Click()
makeNice
Selection.NumberFormat = "0"" mm crs"""
autoMerge 40, 0

AppActivate Application.Caption

End Sub



Private Sub CommandButton_0super_Click()

Dim x, y As String

x = ActiveCell.Value

If x = "" Then
    y = "do nothing"
Else
    If InStr(x, ".") > 1 Then
        If Len(Split(x, ".")(0)) >= 3 Then
           CommandButton_zero_Click
        ElseIf Len(Split(x, ".")(0)) = 2 Then
            CommandButton_zerozero_Click
        ElseIf Len(Split(x, ".")(0)) = 1 And Len(Split(x, ".")(1)) = 1 Then 'like 1.1
            CommandButton_zerozero_nomerge_Click
        ElseIf Split(x, ".")(0) = 0 And Len(Split(x, ".")(1)) = 2 Then 'like 0.28
            CommandButton_zerozerozero_Click
        ElseIf Len(Split(x, ".")(0)) = 1 And Split(x, ".")(0) = 0 Then 'like 0.5
            CommandButton_zerozero_nomerge_Click
        ElseIf Len(Split(x, ".")(0)) = 1 Then
            CommandButton_zerozero_Click
        End If
    Else
        CommandButton_zero_Click
    End If

End If


End Sub

Private Sub CommandButton_beta_Click()
Dim txt As String
txt = ChrW(&H3B2)
SetClipboard (txt)
End Sub

Private Sub CommandButton_boltV_Click()
'make all nice
makeNice
ActiveCell.Resize(7, 9).Merge
ActiveCell.UnMerge

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "Try"
ActiveCell.Offset(0, 1).Activate
'ActiveCell.Resize(1, 2).Merge

ActiveCell.Characters.Font.Color = _
RGB(0, 0, 255)
ActiveCell.NumberFormat = "0"
ActiveCell.FormulaR1C1 = "2"
ActiveCell.HorizontalAlignment = xlRight
ActiveCell.Offset(0, 1).Activate

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "M"
ActiveCell.HorizontalAlignment = xlRight
ActiveCell.Offset(0, 1).Activate

ActiveCell.Characters.Font.Color = _
RGB(0, 0, 255)
ActiveCell.NumberFormat = 0
ActiveCell.FormulaR1C1 = "20"
ActiveCell.Offset(2, -3).Activate

ActiveCell.FormulaR1C1 = ChrW(&H3C6) & "Vfn"
ActiveCell.Characters(start:=3, Length:=2).Font.Subscript = True
ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 4).Merge

ActiveCell.NumberFormat = "##0.###"" kN"""
'print something useful
ActiveCell.FormulaR1C1 = "=IF(R[-2]C[0]= 16 ,59.3,IF(R[-2]C[0]=20,92.6,IF(R[-2]C[0]=24,133,IF(R[-2]C[0]=30,214,IF(R[-2]C[0]=36,313," & Chr(34) & "err" & Chr(34) & ")))))"
ActiveCell.Offset(2, -3).Activate
ActiveCell.Resize(1, 3).Merge

ActiveCell.FormulaR1C1 = "=R[-2]C[3]"
ActiveCell.Resize(1, 3).Merge
ActiveCell.NumberFormat = "##0.0###"" kN"""
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=R[-4]C[-3]"
ActiveCell.NumberFormat = 0
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "'="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=rounddown(R[0]C[-6]*R[0]C[-2],0)"
ActiveCell.Resize(1, 3).Merge
ActiveCell.NumberFormat = "0"" kN"""
ActiveCell.Offset(2, -6).Activate
ActiveCell.Resize(1, 3).Merge
ActiveCell.NumberFormat = "0"" kN"""
ActiveCell.FormulaR1C1 = "=" 'capacity value
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=IF(RC[-3]<RC[1]," & Chr(34) & "< " & Chr(34) & "," & Chr(34) & "!" & Chr(34) & ")"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 3).Merge
ActiveCell.NumberFormat = "0"" kN"""
ActiveCell.FormulaR1C1 = "=R[-2]C[2]" 'design value
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(" & Chr(34) & "<" & Chr(34) & ",RC[-4]))," & Chr(34) & ChrW(&H2192) & " OK" & Chr(34) & ", " & Chr(34) & "!!!!!!!!!!!!!!!!!!!!" & Chr(34) & ")"
ActiveCell.Offset(0, -7).Activate

AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_BU_Click()
makeNice
Dim xxx, yyy, zzz As String
Selection.NumberFormat = "0"" BU"""
autoMerge 16, 0

AppActivate Application.Caption
End Sub

Private Sub CommandButton_cantilever_defl_Click()
makeNice
ActiveCell.Resize(3, 21).Merge
ActiveCell.UnMerge

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "f"

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "q × L4 / 8 E × I"
ActiveCell.Characters(start:=6, Length:=1).Font.Superscript = True
ActiveCell.Offset(0, 5).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate


ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.00"" kN/m"""
ActiveCell.Resize(1, 3).Merge

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
'ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.0"" m"""
ActiveCell.Resize(1, 2).Merge
ActiveCell.HorizontalAlignment = xlRight

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = ChrW(&H2074)
ActiveCell.Offset(2, -12).Activate
ActiveCell.FormulaR1C1 = "/ 8 ×"
ActiveCell.Offset(0, 2).Activate
'E
ActiveCell.Resize(1, 3).Merge
ActiveCell.FormulaR1C1 = "200"
Selection.NumberFormat = "0"" GPa"""
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
'I
ActiveCell.Resize(1, 6).Merge
ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.00"" × 10^6 mm""" & ChrW(&H2074)

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=R[-2]C[-7]*R[-2]C[-3]^4/(8*R[0]C[-11]*R[0]C[-7])*1000"
ActiveCell.Resize(1, 4).Merge
Selection.NumberFormat = "0"" mm"""
ActiveCell.Offset(-2, -3).Activate
ActiveCell.FormulaR1C1 = "="
AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_cantilever_deflF_Click()
makeNice
ActiveCell.Resize(3, 21).Merge
ActiveCell.UnMerge

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "f"

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "F × L3 / 3 E × I"
ActiveCell.Characters(start:=6, Length:=1).Font.Superscript = True
ActiveCell.Offset(0, 5).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate


ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.00"" kN"""
ActiveCell.Resize(1, 3).Merge

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
'ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.0"" m"""
ActiveCell.Resize(1, 2).Merge
ActiveCell.HorizontalAlignment = xlRight

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = Chr(179)
ActiveCell.Offset(2, -12).Activate
ActiveCell.FormulaR1C1 = "/ 3 ×"
ActiveCell.Offset(0, 2).Activate
'E
ActiveCell.Resize(1, 3).Merge
ActiveCell.FormulaR1C1 = "200"
Selection.NumberFormat = "0"" GPa"""
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
'I
ActiveCell.Resize(1, 6).Merge
ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.00"" × 10^6 mm""" & ChrW(&H2074)

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=R[-2]C[-7]*R[-2]C[-3]^3/(3*R[0]C[-11]*R[0]C[-7])*1000"
ActiveCell.Resize(1, 4).Merge
Selection.NumberFormat = "0"" mm"""
ActiveCell.Offset(-2, -3).Activate
ActiveCell.FormulaR1C1 = "="
AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_cantilever_M_Click()
makeNice
ActiveCell.Resize(1, 21).Merge
ActiveCell.UnMerge

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "M*x"
ActiveCell.Characters(start:=3, Length:=3).Font.Subscript = True

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "q × L2 / 2"
ActiveCell.Characters(start:=6, Length:=1).Font.Superscript = True
ActiveCell.Offset(0, 3).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.0"" kN/m"""
ActiveCell.Resize(1, 3).Merge

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 1).Activate
'ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.0"" m"""
ActiveCell.Resize(1, 2).Merge
ActiveCell.HorizontalAlignment = xlRight

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = Chr(178)
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "/ 2"
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=R[0]C[-9]*R[0]C[-5]^2/2"
Selection.NumberFormat = "0.0"" kNm"""
autoMerge 38, 2
ActiveCell.Offset(0, -5).Activate
ActiveCell.FormulaR1C1 = "="
AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_check_kN_Click()
'Does 3 things:
'1. guess the units of the formula in the cell (works only for A+B scenario)
'2. if cell is empty - calculates equiation (the code in other sub!)
'3. if cell contains a link to other cell - gets its units
ActiveCell.NumberFormat = "General"

On Error GoTo eh
start_check_kN:

Dim y, z, tt, zz, xx, vv, ttx As String
Dim ctr As Integer
Dim double_element_equation, link_equation As Boolean


xx = ActiveCell.Formula 'not Text(because it shows ###), not Value(because it shows result)
tt = Replace(xx, "=", "")

If InStr(tt, "cos") Then GoTo Done
If InStr(tt, "sin") Then GoTo Done
If InStr(tt, "tan") Then GoTo Done
If InStr(tt, "NAME?") Then GoTo Done


'check for range input
If InStr(tt, ":") Then
    MsgBox ("it cannot be range")
    GoTo Done
End If

'check what inside of cell to call proper part of the script
vv = ActiveCell.Formula

check_restart:
If InStr(vv, "*") > 1 Then
    ctr = ctr + 1
    vv = Replace(vv, "*", "", , 1)
    GoTo check_restart
ElseIf InStr(vv, "/") > 1 Then
    ctr = ctr + 1
    vv = Replace(vv, "/", "", , 1)
    GoTo check_restart
ElseIf InStr(vv, "+") > 1 Then
    ctr = ctr + 1
    vv = Replace(vv, "+", "", , 1)
    GoTo check_restart
ElseIf InStr(vv, "-") > 1 Then
    ctr = ctr + 1
    vv = Replace(vv, "-", "", , 1)
    GoTo check_restart
End If
      
If ctr = 1 Then
    double_element_equation = True
End If

If ctr = 0 And vv <> "" Then
    link_equation = True
End If
    

'>>>>1.START guess units for equiations with two objects only like A+B


If InStr(xx, "*") > 1 And double_element_equation = True Then
    zz = Split(tt, "*")

    '>>NB check if it's not two units but A1/2 for example
    If IfThereIsLetterInStr(CStr(zz(0))) = False And IfThereIsLetterInStr(CStr(zz(1))) = True Then
        link_equation = True
        ttx = zz(1)
        tt = zz(1)
        GoTo link_check
    End If
    
    If IfThereIsLetterInStr(CStr(zz(0))) = True And IfThereIsLetterInStr(CStr(zz(1))) = False Then
        link_equation = True
        ttx = zz(0)
        tt = zz(0)
        GoTo link_check
    End If
    '<<end NB check
        
    y = Range(zz(0)).NumberFormat
    z = Range(zz(1)).NumberFormat
    'order is important!
    If InStr(y + z, "kN/m") > 1 And InStr(Replace(y + z, "kN/m", "", , 1), "m") > 1 Then
        CommandButton_kN_Click
    ElseIf InStr(y + z, "kg/m²") > 1 And InStr(y + z, "m²") > 1 Then
        CommandButton_kN_Click
    ElseIf InStr(y + z, "m") > 1 And InStr(Replace(y + z, "m", "", , 1), "m") > 1 Then
        CommandButton_m2_Click
    ElseIf InStr(y + z, "mm²") > 1 And InStr(y + z, "Pa") > 1 Then
        CommandButton_kN_Click
    ElseIf InStr(y + z, "kN") > 1 And InStr(y + z, "m") > 1 Then
        CommandButton_kNm_Click
    ElseIf InStr(y + z, "kPa") > 1 And InStr(y + z, "m²") > 1 Then
        CommandButton_kN_Click
    ElseIf InStr(y + z, "kPa") > 1 And InStr(y + z, "m") > 1 And InStr(y + z, "kN") = 0 Then
        CommandButton_kNperm_Click
    ElseIf InStr(y + z, "kN") > 1 Then
        CommandButton_kN_Click
    ElseIf InStr(y + z, "kPa") > 1 Then
        CommandButton_kPa_Click
    ElseIf InStr(y + z, "m") > 1 Then
        CommandButton_m_Click
    Else
        CommandButton_0super_Click
    End If
 
 
ElseIf InStr(xx, "+") > 1 And double_element_equation = True Then
    zz = Split(tt, "+")
    
    '>>NB check if it's not two units but A1/2 for example
    If IfThereIsLetterInStr(CStr(zz(0))) = False And IfThereIsLetterInStr(CStr(zz(1))) = True Then
        link_equation = True
        ttx = zz(1)
        tt = zz(1)
        GoTo link_check
    End If
    
    If IfThereIsLetterInStr(CStr(zz(0))) = True And IfThereIsLetterInStr(CStr(zz(1))) = False Then
        link_equation = True
        ttx = zz(0)
        tt = zz(0)
        GoTo link_check
    End If
    '<<end NB check
    
    y = Range(zz(0)).NumberFormat
    z = Range(zz(1)).NumberFormat
    'order is important!
    If InStr(y + z, "kNm") > 1 Then
        CommandButton_kNm_Click
    ElseIf InStr(y + z, "kPa") > 1 Then
        CommandButton_kPa_Click
    ElseIf InStr(y + z, "kN") > 1 Then
        CommandButton_kN_Click
     ElseIf InStr(y + z, "m") > 1 Then
        CommandButton_m_Click
    Else
        CommandButton_0super_Click
    End If
    
ElseIf InStr(xx, "/") > 1 And double_element_equation = True Then

    zz = Split(tt, "/")
    
    '>>NB check if it's not two units but A1/2 for example
    If IfThereIsLetterInStr(CStr(zz(0))) = False And IfThereIsLetterInStr(CStr(zz(1))) = True Then
        link_equation = True
        ttx = zz(1)
        tt = zz(1)
        GoTo link_check
    End If
    
    If IfThereIsLetterInStr(CStr(zz(0))) = True And IfThereIsLetterInStr(CStr(zz(1))) = False Then
        link_equation = True
        ttx = zz(0)
        tt = zz(0)
        GoTo link_check
    End If
    '<<end NB check
    
    y = Range(zz(0)).NumberFormat
    z = Range(zz(1)).NumberFormat

    'order is important!
    If InStr(y + z, "kNm") > 1 And InStr(Replace(y + z, "kNm", "", , 1), "kNm") > 1 Then
        CommandButton_0super_Click
    ElseIf InStr(y + z, "kN") > 1 And InStr(Replace(y + z, "kN", "", , 1), "kN") > 1 Then
        CommandButton_0super_Click
    ElseIf InStr(y + z, "kN") > 1 And InStr(Replace(y + z, "kN", "", , 1), "mm²") > 1 Then
        CommandButton_kPa_Click
    ElseIf InStr(y + z, "kN") > 1 And InStr(Replace(y + z, "kN", "", , 1), "m²") > 1 Then
        CommandButton_kPa_Click

    ElseIf InStr(y + z, "kNm") > 1 And InStr(y + z, "m") > 1 Then
        CommandButton_kN_Click
    ElseIf InStr(y + z, "kPa") > 1 And InStr(y + z, "kN") > 1 Then
        CommandButton_m_Click
    ElseIf InStr(y + z, "kN") > 1 And InStr(y + z, "m") > 1 Then
        CommandButton_kNperm_Click
    ElseIf InStr(y + z, "kPa") > 1 And InStr(y + z, "m") > 1 Then
        CommandButton_kNperm_Click
    ElseIf InStr(y + z, "mm") > 1 And InStr(Replace(y + z, "mm", "", , 1), "m") > 1 And InStr(y + z, "kN") < 1 And InStr(y + z, "kPa") < 1 Then
        CommandButton_0super_Click
    ElseIf InStr(y + z, "kN") > 1 Then
        CommandButton_kN_Click
    ElseIf InStr(y + z, "m") > 1 Then
        CommandButton_m_Click
    ElseIf InStr(y + z, "MPa") > 1 Then
        CommandButton_MPa_Click
    ElseIf InStr(y + z, "kPa") > 1 Then
        CommandButton_kPa_Click
    ElseIf InStr(y + z, "GPa") > 1 Then
        CommandButton_GPa_Click
    Else

        CommandButton_0super_Click
    End If

ElseIf InStr(xx, "-") > 1 And double_element_equation = True Then
    zz = Split(tt, "-")
    
    '>>NB check if it's not two units but A1/2 for example
    If IfThereIsLetterInStr(CStr(zz(0))) = False And IfThereIsLetterInStr(CStr(zz(1))) = True Then
        link_equation = True
        ttx = zz(1)
        tt = zz(1)
        GoTo link_check
    End If
    
    If IfThereIsLetterInStr(CStr(zz(0))) = True And IfThereIsLetterInStr(CStr(zz(1))) = False Then
        link_equation = True
        ttx = zz(0)
        tt = zz(0)
        GoTo link_check
    End If
    '<<end NB check
    
    y = Range(zz(0)).NumberFormat
    z = Range(zz(1)).NumberFormat
    'order is important!
    If InStr(y + z, "kNm") > 1 Then
        CommandButton_kNm_Click
    ElseIf InStr(y + z, "kPa") > 1 Then
        CommandButton_kPa_Click
    ElseIf InStr(y + z, "kN") > 1 Then
        CommandButton_kN_Click
    ElseIf InStr(y + z, "m") > 1 Then
        CommandButton_m_Click
    Else
        CommandButton_0super_Click
    End If
    
'>>END guessing the units





'2.now if the cell is empty we're looking for equation
'>>>>>START
ElseIf xx = "" Then
    If ActiveCell.Column < 4 Then GoTo Done 'lazy check it can't be an equiation such small, this check also in the ComplexCalculation_Click module
    ComplexCalculation_Click
    GoTo start_check_kN
ElseIf xx = "=" Then
    ActiveCell.Value = ""
    ComplexCalculation_Click
    GoTo start_check_kN
'>>END


Else
'>>>>>3.START link check
''MsgBox ("check for letters to see if that's a link")
''MsgBox (tt)
'check for external link

    If InStr(tt, "!") > 0 Then
        ttx = Split(tt, "!")(1)
    Else
        ttx = tt
    End If
'' (ttx)
link_check:
    If Asc(Left(ttx, 1)) > 64 And Asc(Left(ttx, 1)) < 91 And link_equation = True Then
        'MsgBox ("start 3. link check module")
        If Range(tt).NumberFormat = "0"" kN""" Then
            Selection.UnMerge
            Call CommandButton_0kN_Click
        ElseIf Range(tt).NumberFormat = "0.0"" kN""" Then
            Selection.UnMerge
            Call CommandButton_00kN_Click
        ElseIf Range(tt).NumberFormat = "0.00"" kN""" Then
            Selection.UnMerge
            Call CommandButton_000kN_Click
                 
        ElseIf Range(tt).NumberFormat = "0"" kNm""" Then
            Selection.UnMerge
            Call CommandButton_0kNm_Click
        ElseIf Range(tt).NumberFormat = "0.0"" kNm""" Then
            Selection.UnMerge
            Call CommandButton_00kNm_Click
        ElseIf Range(tt).NumberFormat = "0.00"" kNm""" Then
            Selection.UnMerge
            Call CommandButton_000kNm_Click
                     
        ElseIf Range(tt).NumberFormat = "0"" kN/m""" Then
            Selection.UnMerge
            Call CommandButton_0knperm_Click
        ElseIf Range(tt).NumberFormat = "0.0"" kN/m""" Then
            Selection.UnMerge
            Call CommandButton_00knperm_Click
        ElseIf Range(tt).NumberFormat = "0.00"" kN/m""" Then
            Selection.UnMerge
            Call CommandButton_000knperm_Click
                  
        ElseIf Range(tt).NumberFormat = "0"" kPa""" Then
            Selection.UnMerge
            Call CommandButton_0kPa_Click
        ElseIf Range(tt).NumberFormat = "0.0"" kPa""" Then
            Selection.UnMerge
            Call CommandButton_00kPa_Click
        ElseIf Range(tt).NumberFormat = "0.00"" kPa""" Then
            Selection.UnMerge
            Call CommandButton_000kPa_Click
                 
        ElseIf Range(tt).NumberFormat = "0"" MPa""" Then
            Selection.UnMerge
            Call CommandButton_0MPa_Click
        ElseIf Range(tt).NumberFormat = "0.0"" MPa""" Then
            Selection.UnMerge
            Call CommandButton_00MPa_Click
        ElseIf Range(tt).NumberFormat = "0.00"" MPa""" Then
            Selection.UnMerge
            Call CommandButton_000MPa_Click
                     
        ElseIf Range(tt).NumberFormat = "0"" GPa""" Then
            Selection.UnMerge
            Call CommandButton_0GPa_Click
        ElseIf Range(tt).NumberFormat = "0.0"" GPa""" Then
            Selection.UnMerge
            Call CommandButton_00GPa_Click
        ElseIf Range(tt).NumberFormat = "0.00"" GPa""" Then
            Selection.UnMerge
            Call CommandButton_000GPa_Click
                 
        ElseIf Range(tt).NumberFormat = "0"" m""" Then
            Selection.UnMerge
            Call CommandButton_0m_Click
        ElseIf Range(tt).NumberFormat = "0.0"" m""" Then
            Selection.UnMerge
            Call CommandButton_00m_Click
        ElseIf Range(tt).NumberFormat = "0.00"" m""" Then
            Selection.UnMerge
            Call CommandButton_000m_Click
                            
        ElseIf Range(tt).NumberFormat = "0"" m²""" Then
            Selection.UnMerge
            Call CommandButton_0m2_Click
        ElseIf Range(tt).NumberFormat = "0.0"" m²""" Then
            Selection.UnMerge
            Call CommandButton_00m2_Click
                                 
        ElseIf Range(tt).NumberFormat = "0"" mm""" Then
            Selection.UnMerge
            Call CommandButton_0mm_Click
        ElseIf Range(tt).NumberFormat = "0"" mm²""" Then
            Selection.UnMerge
            Call CommandButton_0mm2_Click
        ElseIf Range(tt).NumberFormat = "0"" mm³""" Then
            Selection.UnMerge
            Call CommandButton_0mm3_Click

        ElseIf Range(tt).NumberFormat = "0" Then
            Selection.UnMerge
            Call CommandButton_zero_Click
        ElseIf Range(tt).NumberFormat = "0.0" Then
            Selection.UnMerge
            Call CommandButton_zerozero_Click
        ElseIf Range(tt).NumberFormat = "0.00" Then
            Selection.UnMerge
            Call CommandButton_zerozerozero_Click
                     
        ElseIf Range(tt).NumberFormat = "0.0"" kg/m²""" Then
            Selection.UnMerge
            Call CommandButton_00kgm2_Click
        ElseIf Range(tt).NumberFormat = "0.0%" Then
            Selection.UnMerge
            Call CommandButton_00percent_Click
        ElseIf Range(tt).NumberFormat = "0""°""" Then
            Selection.UnMerge
            Call CommandButton_0deg_Click
                   
        End If
    End If
'END link check

End If

Done:
    Exit Sub
eh:
    MsgBox ("Can't finish CommandButton_check_kN sub")
End

End Sub

Private Sub CommandButton_delete_row_Click()
If ActiveCell.Formula = "" Then
    ActiveCell.EntireRow.Delete
    ActiveCell.Offset(-1, 0).Activate
End If
End Sub

Private Sub CommandButton_delta_Click()
Dim txt As String
txt = ChrW(&H3B4)
SetClipboard (txt)
End Sub

Private Sub CommandButton_dia_Click()

Dim txt As String
  
txt = "Ø"
SetClipboard (txt)

End Sub

Private Sub CommandButton_00kgm2_Click()
makeNice
Selection.NumberFormat = "0.0"" kg/m²"""
autoMerge 40, 1

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0kN_Click()
makeNice
Selection.NumberFormat = "0"" kN"""
autoMerge 22, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0kNm_Click()
makeNice
Selection.NumberFormat = "0"" kNm"""
autoMerge 31, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0kPa_Click()
makeNice
Selection.NumberFormat = "0"" kPa"""
autoMerge 25, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_00m_Click()
makeNice
Selection.NumberFormat = "0.0"" m"""
autoMerge 15, 1

AppActivate Application.Caption

End Sub

Private Sub CommandButton_00m2_Click()
makeNice
Selection.NumberFormat = "0.0"" m²"""
autoMerge 15, 1 'tested on 21.8

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0mm3_Click()
makeNice
Selection.NumberFormat = "0"" mm³"""
autoMerge 20, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0MPa_Click()
makeNice
Selection.NumberFormat = "0"" MPa"""
autoMerge 32, 0 'checked on 40

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0mm2_Click()
makeNice
Selection.NumberFormat = "0"" mm²"""
autoMerge 24, 0 'checked on 100 and 1000

AppActivate Application.Caption

End Sub

Private Sub CommandButton_00kN_Click()
makeNice
Selection.NumberFormat = "0.0"" kN"""
autoMerge 20, 1 'checked 0.5 1.5

AppActivate Application.Caption

End Sub

Private Sub CommandButton_divided_Click()
On Error GoTo eh
Dim Counter As Integer


begg:
If Counter < 20 Then 'to catch the loop if error arise
   Counter = Counter + 1
        If ActiveCell.MergeCells = True Then
            ActiveCell.Offset(0, 1).Select
            GoTo begg
        Else
            ActiveCell.Select
            GoTo zzzd
        End If
    
zzzd:
    If ActiveCell.Formula = "" Then
        ActiveCell.FormulaR1C1 = "'/"
        If ActiveCell.Offset(0, 1).Formula = "" Then
            makeNice
            ActiveCell.HorizontalAlignment = xlCenter
            ActiveCell.Offset(0, 1).Activate
            makeNice
            ActiveCell.FormulaR1C1 = "="
            AppActivate Application.Caption
            SendKeys "{F2}"

        Else
            ActiveCell.Offset(0, 1).Activate
            AppActivate Application.Caption
        End If
        Counter = Counter + 1
    
    ElseIf ActiveCell.FormulaR1C1 = "/" Then
        ActiveCell.Offset(0, 1).Activate
        AppActivate Application.Caption
        SendKeys "{F2}"
        GoTo Done
    Else
        ActiveCell.Offset(0, 1).Activate
        GoTo begg
    End If

Else
    GoTo Done
End If

Done:
    Exit Sub
eh:
    MsgBox ("err")
End

End Sub


Private Sub CommandButton_equal_Click()
On Error GoTo eh
Dim Counter As Integer


begg:
If Counter < 20 Then 'to catch the loop if error arise
   Counter = Counter + 1
        If ActiveCell.MergeCells = True Then
            ActiveCell.Offset(0, 1).Select
            GoTo begg
        Else
            ActiveCell.Select
            GoTo zzzd
        End If
    
zzzd:
    If ActiveCell.Formula = "" Then
        ActiveCell.FormulaR1C1 = "'="
        If ActiveCell.Offset(0, 1).Formula = "" Then
            makeNice
            ActiveCell.HorizontalAlignment = xlCenter
            ActiveCell.Offset(0, 1).Activate
            makeNice
            ActiveCell.FormulaR1C1 = "="
            AppActivate Application.Caption
            SendKeys "{F2}"

        Else
            ActiveCell.Offset(0, 1).Activate
            AppActivate Application.Caption
        End If
        Counter = Counter + 1
    
    ElseIf ActiveCell.FormulaR1C1 = "=" Then
        ActiveCell.Offset(0, 1).Activate
        AppActivate Application.Caption
        SendKeys "{F2}"
        GoTo Done
    Else
        ActiveCell.Offset(0, 1).Activate
        GoTo begg
    End If

Else
    GoTo Done
End If

Done:
    Exit Sub
eh:
    MsgBox ("err")
End

End Sub
Private Sub CommandButton_fi_Click()

Dim txt As String
txt = ChrW(&H3C6)
SetClipboard (txt)
  
End Sub

Private Sub CommandButton_fimn_Click()
makeNice

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = ChrW(&H3C6) & "Mn"
ActiveCell.Characters(start:=3, Length:=1).Font.Subscript = True

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate


makeNice


AppActivate Application.Caption
SendKeys "{F2}"


End Sub

Private Sub CommandButton_goal_seek_Click()
Range(TextBox1).GoalSeek Goal:=0, ChangingCell:=Range(TextBox2)
End Sub

Private Sub CommandButton_fintf_Click()

'make all nice
makeNice
ActiveCell.Resize(7, 9).Merge
ActiveCell.UnMerge

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "Try"
ActiveCell.Offset(0, 1).Activate
'ActiveCell.Resize(1, 2).Merge

ActiveCell.Characters.Font.Color = _
RGB(0, 0, 255)
ActiveCell.NumberFormat = "0"
ActiveCell.FormulaR1C1 = "2"
ActiveCell.HorizontalAlignment = xlRight
ActiveCell.Offset(0, 1).Activate

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "M"
ActiveCell.HorizontalAlignment = xlRight
ActiveCell.Offset(0, 1).Activate

ActiveCell.Characters.Font.Color = _
RGB(0, 0, 255)
ActiveCell.NumberFormat = 0
ActiveCell.FormulaR1C1 = "20"
ActiveCell.Offset(2, -3).Activate

ActiveCell.FormulaR1C1 = ChrW(&H3C6) & "Ntf"
ActiveCell.Characters(start:=3, Length:=2).Font.Subscript = True
ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 4).Merge

ActiveCell.NumberFormat = "0"" kN"""
'print something useful
ActiveCell.FormulaR1C1 = "=IF(R[-2]C[0]= 16 ,104,IF(R[-2]C[0]=20,163,IF(R[-2]C[0]=24,234,IF(R[-2]C[0]=30,373,IF(R[-2]C[0]=36,541," & Chr(34) & "err" & Chr(34) & ")))))"
ActiveCell.Offset(2, -3).Activate
ActiveCell.Resize(1, 3).Merge

ActiveCell.FormulaR1C1 = "=R[-2]C[3]"
ActiveCell.Resize(1, 3).Merge
ActiveCell.NumberFormat = "0"" kN"""
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=R[-4]C[-3]"
ActiveCell.NumberFormat = 0
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "'="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=rounddown(R[0]C[-6]*R[0]C[-2],0)"
ActiveCell.Resize(1, 3).Merge
ActiveCell.NumberFormat = "0"" kN"""
ActiveCell.Offset(2, -6).Activate
ActiveCell.Resize(1, 3).Merge
ActiveCell.NumberFormat = "0"" kN"""
ActiveCell.FormulaR1C1 = "=" 'capacity value
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=IF(RC[-3]<RC[1]," & Chr(34) & "< " & Chr(34) & "," & Chr(34) & "!" & Chr(34) & ")"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 3).Merge
ActiveCell.NumberFormat = "0"" kN"""
ActiveCell.FormulaR1C1 = "=R[-2]C[2]" 'design value
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(" & Chr(34) & "<" & Chr(34) & ",RC[-4]))," & Chr(34) & ChrW(&H2192) & " OK" & Chr(34) & ", " & Chr(34) & "!!!!!!!!!!!!!!!!!!!!" & Chr(34) & ")"
ActiveCell.Offset(0, -7).Activate

AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_goto_Click()
On Error GoTo eh
Dim x, y, z, j, second, Result, yx  As String
Dim suspcount As Integer
'Dim PreviousCell As Range ---> see public var in the beginning
'we need it for Go Back
Set PreviousCell = ActiveCell
PreviouseCellSheet = ActiveSheet.Name


x = ActiveCell.Formula 'mind this is not a string but variant type

'check for a formula
If InStr(x, "*") > 0 Or InStr(x, "/") > 0 Or InStr(x, "-") > 0 Or InStr(x, "+") > 0 Then
    MsgBox ("no formulas allowed as input")
    GoTo Done
End If

'check for a letter with special function
If IfThereIsLetterInStr(CStr(x)) = False Then GoTo Done



If x <> "" Then
    y = Replace(x, "=", "")
        
    If InStr(y, "!") > 0 Then
        yx = Split(y, "!")(1)
    Else
        yx = y
    End If

    
    
    z = Range(yx).Column
    
    
    Dim i, StringLength As Integer
    StringLength = Len(yx)
    
    For i = 1 To StringLength Step 1
        If IsNumeric(Mid(yx, i, 1)) Then
            Result = Result & Mid(yx, i, 1)
        End If
    Next i
    
    
   ''''MsgBox (Result)
    If CInt(Result) - 10 > 0 Then
        second = CStr(CInt(Result) - 10)
    Else
        second = CStr(1)
    End If
   ''''MsgBox (second)
    
    If InStr(y, "!") > 0 Then
        '''MsgBox (Split(y, "!")(0) & "!a" & second)
        Application.Goto Range(Split(y, "!")(0) + "!a" + second), True
    Else
        Application.Goto Range("a" + second), True
    End If

    
    
    Range(y).Select
    AppActivate Application.Caption
    
End If

Done:
    Exit Sub
eh:
    MsgBox ("error")
End
End Sub

Private Sub CommandButton_gofrom_Click()
On Error GoTo eh
Dim x, y, z, j, second, Result, yx As String

If PreviousCell.Address <> "" Then

    y = "'" + PreviouseCellSheet + "'!" + Replace(PreviousCell.Address, "$", "")
    z = PreviousCell.Column
    
    
    
    If InStr(y, "!") > 0 Then
        yx = Split(y, "!")(1)
    Else
        yx = y
    End If
    
    Dim i, StringLength As Integer
    StringLength = Len(yx)
    
    For i = 1 To StringLength Step 1
        If IsNumeric(Mid(yx, i, 1)) Then
            Result = Result & Mid(yx, i, 1)
        End If
    Next i
    'MsgBox (z & "    " & Result)
    If CInt(Result) - 10 > 0 Then
        second = CStr(CInt(Result) - 10)
    Else
        second = CStr(1)
    End If
    
    
    second = CStr(CInt(Result) - 10)
    '''MsgBox (y)
    If InStr(y, "!") > 0 Then
        ''MsgBox (yx)
        ''MsgBox (Split(y, "!")(0) & "!a" & second)
        Application.Goto Range(Split(y, "!")(0) + "!a" + second), True
    Else
        Application.Goto Range("a" + second), True
    End If
    
    PreviousCell.Select
    AppActivate Application.Caption

End If

Done:
    Exit Sub
eh:
    MsgBox ("error")
End
End Sub

Private Sub CommandButton_kg_Click()
makeNice
Dim xxx, yyy, zzz As String
'xxx = Chr(34) & "0" & Chr(34)
'yyy = Chr(34) & " kg" & Chr(34) & Chr(34)
'zzz = xxx & yyy
'ActiveCell.FormulaR1C1 = yyy
'Selection.NumberFormat = zzz
Selection.NumberFormat = "0"" kg"""
autoMerge 16, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_kNm_Click()
Dim x, y As String

x = ActiveCell.Formula

If x = "" Then
    y = "do nothing"
Else
    If Len(Split(x, ".")(0)) >= 3 Then
       CommandButton_0kNm_Click
    ElseIf Len(Split(x, ".")(0)) = 2 Then
        CommandButton_00kNm_Click
    ElseIf Len(Split(x, ".")(0)) = 1 And Split(x, ".")(0) = 0 Then
        CommandButton_00kNm_Click
    ElseIf Len(Split(x, ".")(0)) = 1 Then
        CommandButton_00kNm_Click
    End If

End If

End Sub

Private Sub CommandButton_kNm3_Click()
makeNice
Selection.NumberFormat = "0"" kN/m³"""
autoMerge 40, 0

AppActivate Application.Caption
End Sub

Private Sub CommandButton_kNperm_Click()
Dim x, y As String

x = ActiveCell.Value

If x = "" Then
    y = "do nothing"
Else
    If Len(Split(x, ".")(0)) >= 3 Then
       CommandButton_0knperm_Click
    ElseIf Len(Split(x, ".")(0)) = 2 Then
        CommandButton_00knperm_Click
    ElseIf Len(Split(x, ".")(0)) = 1 And Split(x, ".")(0) = 0 Then
        CommandButton_00knperm_Click
    ElseIf Len(Split(x, ".")(0)) = 1 Then
        CommandButton_000knperm_Click
    End If

End If

End Sub

Private Sub CommandButton_kPa_Click()
Dim x, y As String
''''MsgBox ("kN start")
x = ActiveCell.Value2
'''MsgBox ("!!!!!" & X)
If x = "" Then
    y = "do nothing"
Else
    If InStr(x, ".") > 1 Then
        If Len(Split(x, ".")(0)) >= 3 Then
           CommandButton_0kPa_Click
        ElseIf Len(Split(x, ".")(0)) = 2 Then
            CommandButton_00kPa_Click
        ElseIf Len(Split(x, ".")(0)) = 1 And Split(x, ".")(0) = 0 Then
            CommandButton_000kPa_Click
        ElseIf Len(Split(x, ".")(0)) = 1 Then
            CommandButton_000kPa_Click
        End If
    Else
       CommandButton_0kPa_Click
    End If
End If
End Sub


Private Sub CommandButton_ksi_Click()
Dim txt As String
txt = ChrW(&H3C8)
SetClipboard (txt)
End Sub

Private Sub CommandButton_lambda_Click()
Dim txt As String
txt = ChrW(&H3BB)
SetClipboard (txt)
End Sub



Private Sub CommandButton_merge_Click()
Application.DisplayAlerts = False
On Error GoTo eh

Dim Counter, mcells, fcells As Integer
Dim cell As Range


For Each cell In Selection
    Counter = Counter + 1
    If Counter = 2000 Then GoTo Done 'safety catch
    
    If cell.MergeCells = True Then
        mcells = mcells + 1
    End If

    If cell.Formula <> "" Then
        fcells = fcells + 1
    End If
Next

If fcells > 1 Then
    Dim answer As Integer
    answer = MsgBox("Merging keep only left-right cell value, proceed?", vbQuestion + vbYesNo + vbDefaultButton1, "Warning")
End If

If answer = vbNo Then GoTo Done



If ActiveCell.MergeCells And Counter = mcells Then
    Selection.HorizontalAlignment = xlLeft
    Selection.VerticalAlignment = xlVAlignCenter
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 12
    'Selection.Font.Color = _
    'RGB(0, 0, 0)
    ActiveCell.UnMerge
    'Selection(0).Activate    'not working
    ActiveCell.Offset(0, 1).Activate
    AppActivate Application.Caption

Else
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    Selection.VerticalAlignment = xlVAlignCenter
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 12
    'Selection.Font.Color = _
    'RGB(0, 0, 0)
    AppActivate Application.Caption


End If

Done:
    Exit Sub
eh:
    MsgBox ("Can't merge")
    End
    
End Sub

Private Sub CommandButton_minus_Click()
On Error GoTo eh
Dim Counter As Integer


begg:
If Counter < 20 Then 'to catch the loop if error arise
   Counter = Counter + 1
        If ActiveCell.MergeCells = True Then
            ActiveCell.Offset(0, 1).Select
            GoTo begg
        Else
            ActiveCell.Select
            GoTo zzzd
        End If
    
zzzd:
    If ActiveCell.Formula = "" Then
        ActiveCell.FormulaR1C1 = "'-"
        If ActiveCell.Offset(0, 1).Formula = "" Then
            makeNice
            ActiveCell.HorizontalAlignment = xlCenter
            ActiveCell.Offset(0, 1).Activate
            makeNice
            ActiveCell.FormulaR1C1 = "="
            AppActivate Application.Caption
            SendKeys "{F2}"

        Else
            ActiveCell.Offset(0, 1).Activate
            AppActivate Application.Caption
        End If
        Counter = Counter + 1
    
    ElseIf ActiveCell.FormulaR1C1 = "-" Then
        ActiveCell.Offset(0, 1).Activate
        AppActivate Application.Caption
        SendKeys "{F2}"
        GoTo Done
    Else
        ActiveCell.Offset(0, 1).Activate
        GoTo begg
    End If

Else
    GoTo Done
End If

Done:
    Exit Sub
eh:
    MsgBox ("err")
End

End Sub


Private Sub CommandButton_Moments_Click()
Dim chromePath As String
On Error GoTo eh
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""

Shell (chromePath & " -url http://doctorlom.com/item173.html#sharnir")
Done:
    Exit Sub
eh:
    MsgBox ("check Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End
End Sub

Private Sub CommandButton_moreequal_Click()

On Error GoTo eh
Dim Counter As Integer

zzzd:
If Counter < 20 Then 'to catch the loop if error arise
    If ActiveCell.Formula = "" And ActiveCell.MergeCells = False Then
        ActiveCell.FormulaR1C1 = "'+"
        If ActiveCell.Offset(0, 1).Formula = "" Then
            makeNice
            ActiveCell.NumberFormat = "@"
            ActiveCell.FormulaR1C1 = ChrW(&H2265)
            ActiveCell.Offset(0, 1).Activate
            makeNice
            ActiveCell.FormulaR1C1 = "="
            AppActivate Application.Caption
            SendKeys "{F2}"

        Else
            ActiveCell.Offset(0, 1).Activate
            AppActivate Application.Caption
        End If
    Counter = Counter + 1
    Else
begg:
        Counter = Counter + 1
        If ActiveCell.MergeCells = True Then
            ActiveCell.Offset(0, 1).Activate
            GoTo begg
        Else
            ActiveCell.Activate
            GoTo zzzd
        End If
    End If

Else
    GoTo Done
End If

Done:
    Exit Sub
eh:
    MsgBox ("err")
End

End Sub


End Sub



Private Sub CommandButton_mu_Click()
Dim txt As String
txt = ChrW(&H3BC)
SetClipboard (txt)
End Sub

Private Sub CommandButton_N_Click()
makeNice

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = ChrW(&H3C6) & "N"
'ActiveCell.Characters(Start:=3, Length:=3).Font.Subscript = True

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 3).Merge
makeNice
Selection.NumberFormat = "0.0"" kN"""

AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_NNN_Click()
Dim txt As String
txt = ChrW(&H3B7)
SetClipboard (txt)
End Sub

Private Sub CommandButton_nuequal_Click()
makeNice

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = ChrW(&H3BC) & " = 1.25"
ActiveCell.Characters(start:=0, Length:=10).Font.Superscript = True
ActiveCell.Characters.Font.Color = _
RGB(0, 0, 255)
ActiveCell.HorizontalAlignment = xlCenter
AppActivate Application.Caption
SendKeys "{F2}"
SendKeys "^a"
End Sub

Private Sub CommandButton_nzs3404_Click()
On Error GoTo eh
Dim chromePath As String
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Shell (chromePath & " -url https://www.dropbox.com/s/7ng0kvd96wdup47/NZS%203404.1%262-1997%20%28Steel%20structures%29.pdf?dl=0")
Done:
    Exit Sub
eh:
    MsgBox ("check file address or Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End
End Sub

Private Sub CommandButton_OK_Click()
makeNice
ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(" & Chr(34) & "<" & Chr(34) & ",RC[-4]))," & Chr(34) & ChrW(&H2192) & " OK" & Chr(34) & ", " & Chr(34) & "!!!!!!!!!!!!!!!!!!!!" & Chr(34) & ")"
ActiveCell.Offset(0, 1).Activate
End Sub

Private Sub CommandButton_omega_Click()
Dim txt As String
txt = ChrW(&H3C9)
SetClipboard (txt)
End Sub

Private Sub CommandButton_PFC_Click()
On Error GoTo eh
Dim chromePath As String
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Shell (chromePath & " -url https://www.dropbox.com/s/znrjv9lkppc3zah/S%26T_Design_With_Steel_2013_2_Part6.pdf?dl=0")
Done:
    Exit Sub
eh:
    MsgBox ("check file address or Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End
End Sub

Private Sub CommandButton_pi_Click()
   
Dim txt As String
txt = ChrW(&H3C0)
SetClipboard (txt)

End Sub

Private Sub CommandButton_plus_Click()
On Error GoTo eh
Dim Counter As Integer


begg:
If Counter < 20 Then 'to catch the loop if error arise
   Counter = Counter + 1
        If ActiveCell.MergeCells = True Then
            ActiveCell.Offset(0, 1).Select
            GoTo begg
        Else
            ActiveCell.Select
            GoTo zzzd
        End If
    
zzzd:
    If ActiveCell.Formula = "" Then
        ActiveCell.FormulaR1C1 = "'+"
        If ActiveCell.Offset(0, 1).Formula = "" Then
            makeNice
            ActiveCell.HorizontalAlignment = xlCenter
            ActiveCell.Offset(0, 1).Activate
            makeNice
            ActiveCell.FormulaR1C1 = "="
            AppActivate Application.Caption
            SendKeys "{F2}"

        Else
            ActiveCell.Offset(0, 1).Activate
            AppActivate Application.Caption
        End If
        Counter = Counter + 1
    
    ElseIf ActiveCell.FormulaR1C1 = "+" Then
        ActiveCell.Offset(0, 1).Activate
        AppActivate Application.Caption
        SendKeys "{F2}"
        GoTo Done
    Else
        ActiveCell.Offset(0, 1).Activate
        GoTo begg
    End If

Else
    GoTo Done
End If

Done:
    Exit Sub
eh:
    MsgBox ("err")
End

End Sub


Private Sub CommandButton_plus_sheet_Click()

Dim x As Long
  For x = 1 To 39 '39 rows in a sheet
    ActiveCell.Offset(1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    ActiveCell.EntireRow.Copy
    ActiveCell.Offset(1).EntireRow.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
  Next x

End Sub

Private Sub CommandButton_RHS_Click()
On Error GoTo eh
Dim chromePath As String
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Shell (chromePath & " -url https://www.dropbox.com/s/2xcve7wqo7gewtn/S%26T_Design_With_Steel_2013_2_Part15.pdf?dl=0")
Done:
    Exit Sub
eh:
    MsgBox ("check file address or Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End
End Sub

Private Sub CommandButton_ro_Click()
Dim txt As String
txt = ChrW(&H3C1)
SetClipboard (txt)
End Sub

Private Sub CommandButton_RunConcat_Click()
On Error GoTo eh
makeNice

Dim x
Dim Length

x = Split(TextBox_concat)

Length = UBound(x) - LBound(x)
Dim i As Integer
Dim arr() As String

For i = 0 To Length
   ReDim Preserve arr(i)
   arr(i) = Chr(34) & " " & x(i) & Chr(34) & ","
Next i

ActiveCell.FormulaR1C1 = "=CONCATENATE(" & Join(arr) & ")"

Done:
    Exit Sub
eh:
    End

End Sub

Private Sub CommandButton_s2_Click()

Dim txt As String
  
txt = Chr(178)
SetClipboard (txt)

End Sub

Private Sub CommandButton_S3_Click()
Dim txt As String
  
txt = Chr(179)
SetClipboard (txt)

End Sub


Private Sub CommandButton_s4_Click()
Dim txt As String
  
txt = ChrW(&H2074)
SetClipboard (txt)

End Sub

Private Sub CommandButton_s6_Click()
Dim txt As String
  
txt = ChrW(&H2076)
SetClipboard (txt)
End Sub

Private Sub CommandButton_sesoc_Click()
On Error GoTo eh
Dim chromePath As String
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Shell (chromePath & " -url https://www.dropbox.com/s/znr7hcsijzvithb/SESOC%20-%20Simplified_Design_of_Steel_Members.pdf?dl=0")
Done:
    Exit Sub
eh:
    MsgBox ("check file address or Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End
End Sub

Private Sub CommandButton_SHS_Click()
On Error GoTo eh
Dim chromePath As String

chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""

Shell (chromePath & " -url http://sunsetpatios.com.au/beam-deflection-calculator.php")
Done:
    Exit Sub
eh:
    MsgBox ("check Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End
End Sub

Private Sub CommandButton_SHS222_Click()
On Error GoTo eh
Dim chromePath As String
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Shell (chromePath & " -url https://www.dropbox.com/s/48txhovc3ek22el/S%26T_Design_With_Steel_2013_2_Part20.pdf?dl=0")
Done:
    Exit Sub
eh:
    MsgBox ("check file address or Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End
End Sub

Private Sub CommandButton_sigma_Click()
Dim txt As String
txt = ChrW(&H3C3)
SetClipboard (txt)
End Sub

Private Sub CommandButton_simplebeam_defl_Click()
makeNice
ActiveCell.Resize(3, 21).Merge
ActiveCell.UnMerge

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "f"

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "5 q × L4 / 384 E × I"
ActiveCell.Characters(start:=8, Length:=1).Font.Superscript = True
ActiveCell.Offset(0, 5).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "5"
ActiveCell.HorizontalAlignment = xlRight
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.00"" kN/m"""
ActiveCell.Resize(1, 3).Merge

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
'ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.0"" m"""
ActiveCell.Resize(1, 2).Merge
ActiveCell.HorizontalAlignment = xlRight

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = ChrW(&H2074)
ActiveCell.Offset(2, -15).Activate
ActiveCell.FormulaR1C1 = "/ 384 ×"
ActiveCell.Offset(0, 2).Activate
'E
ActiveCell.Resize(1, 3).Merge
ActiveCell.FormulaR1C1 = "200"
Selection.NumberFormat = "0"" GPa"""
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
'I
ActiveCell.Resize(1, 6).Merge
ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.00"" × 10^6 mm""" & ChrW(&H2074)

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=5*R[-2]C[-4]*R[-2]C[0]^4/(384*R[0]C[-11]*R[0]C[-7])*1000"
ActiveCell.Resize(1, 4).Merge
Selection.NumberFormat = "0"" mm"""
ActiveCell.Offset(-2, 0).Activate
ActiveCell.FormulaR1C1 = "="
AppActivate Application.Caption
SendKeys "{F2}"

End Sub

Private Sub CommandButton_simplebeam_defl_timber_Click()
makeNice
ActiveCell.Resize(7, 21).Merge
ActiveCell.UnMerge


ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "Try"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 2).Merge
ActiveCell.FormulaR1C1 = "90"
ActiveCell.HorizontalAlignment = xlRight
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 2).Merge
ActiveCell.FormulaR1C1 = "90"
ActiveCell.NumberFormat = "0"" (h)"""
ActiveCell.HorizontalAlignment = xlLeft
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 2).Merge
ActiveCell.FormulaR1C1 = "SG8"
ActiveCell.Offset(2, -6).Activate


ActiveCell.FormulaR1C1 = "k2"
ActiveCell.Characters(start:=2, Length:=1).Font.Subscript = True
ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "2"
ActiveCell.Offset(2, -3).Activate

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "f"
ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate


ActiveCell.FormulaR1C1 = "k2 × 5 q × L4 / 384 E × I"
ActiveCell.Characters(start:=13, Length:=1).Font.Superscript = True
ActiveCell.Characters(start:=2, Length:=1).Font.Subscript = True
ActiveCell.Offset(0, 6).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=R[-2]C[-7]" 'get k2
ActiveCell.HorizontalAlignment = xlRight
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "5"
ActiveCell.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.00"" kN/m"""
ActiveCell.Resize(1, 3).Merge

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
'ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.0"" m"""
ActiveCell.Resize(1, 2).Merge
ActiveCell.HorizontalAlignment = xlRight

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = ChrW(&H2074)
ActiveCell.Offset(2, -17).Activate
ActiveCell.FormulaR1C1 = "/ 384 ×"
ActiveCell.Offset(0, 2).Activate
'E
ActiveCell.Resize(1, 3).Merge
ActiveCell.FormulaR1C1 = "=IF(R[-6]C[1]=" & Chr(34) & "SG8" & Chr(34) & ", 5.4,IF(R[-6]C[1]=" & Chr(34) & "LVL11" & Chr(34) & ", 11,IF(R[-6]C[1]=" & Chr(34) & "LVL13" & Chr(34) & ", 13,5.4)))"
Selection.NumberFormat = "0.0"" GPa"""
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate
'I
ActiveCell.Resize(1, 6).Merge

ActiveCell.FormulaR1C1 = "=(R[-6]C[-8]*R[-6]C[-5]^3)/(3*1000000)" 'get section modulus
Selection.NumberFormat = "0.00"" × 10^6 mm""" & ChrW(&H2074)

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate

ActiveCell.FormulaR1C1 = "=R[-2]C[-6]*5*R[-2]C[-2]*R[-2]C[2]^4/(384*R[0]C[-11]*R[0]C[-7])*1000"
ActiveCell.Resize(1, 3).Merge
Selection.NumberFormat = "0"" mm"""

ActiveCell.Offset(0, 1).Activate
ActiveCell.Offset(-2, -1).Activate
ActiveCell.FormulaR1C1 = "="


AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_simplebeam_M_Click()
makeNice
ActiveCell.Resize(1, 21).Merge
ActiveCell.UnMerge

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "M*x"
ActiveCell.Characters(start:=3, Length:=3).Font.Subscript = True

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "q × L2 / 8"
ActiveCell.Characters(start:=6, Length:=1).Font.Superscript = True
ActiveCell.Offset(0, 3).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.0"" kN/m"""
ActiveCell.Resize(1, 3).Merge

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.Offset(0, 1).Activate

Selection.NumberFormat = "0.0"" m"""
ActiveCell.Resize(1, 2).Merge
ActiveCell.HorizontalAlignment = xlRight

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = Chr(178)
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "/ 8"
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=R[0]C[-9]*R[0]C[-5]^2/8"
Selection.NumberFormat = "0.0"" kNm"""
autoMerge 38, 2
ActiveCell.Offset(0, -5).Activate
ActiveCell.FormulaR1C1 = "="
AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_UB_Click()
On Error GoTo eh
'Dim pat1, pat2, pat3 As String
'pat1 = """C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe"""
'pat2 = "/A ""page=3"""
'pat3 = """C:\Users\64210\Dropbox (SMC Design Studio)\SMC Design Studio Team Folder\Specifications\STEEL SECTIONS S&T_Design_with_Steel_Nov2015-01.pdf"""
'Shell pat1 & " " & pat2 & " " & pat3, vbNormalFocus

Dim chromePath As String
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Shell (chromePath & " -url https://www.dropbox.com/s/790dxmxyp4d565d/S%26T_Design_With_Steel_2013_2_Part3.pdf?dl=0")
Done:
    Exit Sub
eh:
    MsgBox ("check file address or Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End
End Sub

Private Sub CommandButton_tau_Click()
Dim txt As String
txt = ChrW(&H3C4)
SetClipboard (txt)
End Sub

Private Sub CommandButton_tet_Click()
Dim txt As String
txt = ChrW(&H3B8)
SetClipboard (txt)
End Sub

Private Sub CommandButton_L_Click()
On Error GoTo eh
Dim chromePath As String
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Shell (chromePath & " -url https://www.dropbox.com/s/zimhh8s9o2hnts3/S%26T_Design_With_Steel_2013_2_Part10.pdf?dl=0")
Done:
    Exit Sub
eh:
    MsgBox ("check file address or Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End
End Sub

Private Sub CommandButton_update_Click()
On Error GoTo eh
Dim wks As Object
For Each wks In Worksheets

   wks.PageSetup.LeftFooter = Sheet00.Range("G11").Text + ", " + Sheet00.Range("G12").Text
   wks.PageSetup.RightFooter = "Print Date: &D    " & Replace(wks.Range("B2").Text, ".", "") & "-&P" 'removing dot symbol


Next wks

Sheet00.PageSetup.RightFooter = "Print Date: &D" 'remove frome the title


Done:
    Exit Sub
eh:
    MsgBox ("something went wrong")
    End

End Sub

Private Sub CommandButton_V_Click()
makeNice

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = ChrW(&H3C6) & "V"
'ActiveCell.Characters(Start:=3, Length:=3).Font.Subscript = True

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 3).Merge

makeNice
Selection.NumberFormat = "0.0"" kN"""

AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_weld_Click()
'make all nice
makeNice
ActiveCell.Resize(5, 12).Merge
ActiveCell.UnMerge

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "Try"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 2).Merge

ActiveCell.Characters.Font.Color = _
RGB(0, 0, 255)
ActiveCell.NumberFormat = "0"" mm"""
ActiveCell.FormulaR1C1 = "5"
ActiveCell.Offset(0, 1).Activate

ActiveCell.Characters.Font.Color = _
RGB(0, 0, 255)
ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "SP weld"
ActiveCell.Offset(2, -3).Activate

ActiveCell.FormulaR1C1 = ChrW(&H3C6) & "Vw"
ActiveCell.Characters(start:=3, Length:=1).Font.Subscript = True
ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 4).Merge

ActiveCell.NumberFormat = "0.###"" kN/mm"""
'print something useful
ActiveCell.FormulaR1C1 = "=IF(R[-2]C[0]=" & Chr(34) & "GP weld" & Chr(34) & ",IF(R[-2]C[-2]=5,0.522,IF(R[-2]C[-2]=6,0.626,IF(R[-2]C[-2]=8,0.835,IF(R[-2]C[-2]=10,1.04,IF(R[-2]C[-2]=12,1.25," & Chr(34) & "err1" & Chr(34) & "))))),IF(R[-2]C[0]=" & Chr(34) & "SP weld" & Chr(34) & ", IF(R[-2]C[-2]=5,0.696,IF(R[-2]C[-2]=6,0.835,IF(R[-2]C[-2]=8,1.11,IF(R[-2]C[-2]=10,1.39,IF(R[-2]C[-2]=12,1.67," & Chr(34) & "err2" & Chr(34) & ")))))," & Chr(34) & "err3" & Chr(34) & "))"
ActiveCell.Offset(2, -3).Activate
'added

ActiveCell.Resize(1, 4).Merge
ActiveCell.NumberFormat = "0.000"" kN/mm"""
ActiveCell.FormulaR1C1 = "=" 'capacity value
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=IF(RC[-4]<RC[1]," & Chr(34) & "< " & Chr(34) & "," & Chr(34) & "!" & Chr(34) & ")"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 4).Merge
ActiveCell.NumberFormat = "0.000"" kN/mm"""
ActiveCell.FormulaR1C1 = "=R[-2]C[-2]" 'design value
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(" & Chr(34) & "<" & Chr(34) & ",RC[-5]))," & Chr(34) & ChrW(&H2192) & " OK" & Chr(34) & ", " & Chr(34) & "!!!!!!!!!!!!!!!!!!!!" & Chr(34) & ")"
ActiveCell.Offset(0, -7).Activate

AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_x_Click()

Dim txt As String
  
txt = "×"
SetClipboard (txt)
  
End Sub

Private Sub CommandButton_comp_Click()

ActiveCell.Resize(1, 3).Merge
makeNice
ActiveCell.FormulaR1C1 = "=1"
Selection.NumberFormat = "0.0"" kNm"""
ActiveCell.Offset(0, 1).Activate

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=IF(RC[-4]<RC[1]," & Chr(34) & "< " & Chr(34) & "," & Chr(34) & ">" & Chr(34) & ")"
makeNice
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "="
makeNice
ActiveCell.Resize(1, 3).Merge
makeNice
Selection.NumberFormat = "0.0"" kNm"""
ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(" & Chr(34) & "<" & Chr(34) & ",RC[-5]))," & Chr(34) & ChrW(&H2192) & " OK" & Chr(34) & ", " & Chr(34) & "NOT OK, PLEASE REFER TO ENGINEER" & Chr(34) & ")"

makeNice

'conditional format

ActiveCell.FormatConditions.Delete
ActiveCell.FormatConditions.Add Type:=xlTextString, TextOperator:=xlContains, String:="NOT OK"
ActiveCell.FormatConditions(1).Interior.Color = RGB(255, 0, 0)



ActiveCell.Offset(0, -4).Activate
AppActivate Application.Caption
SendKeys "{F2}"

End Sub

Private Sub CommandButton_0m_Click()
makeNice
Selection.NumberFormat = "0"" m"""
autoMerge 15, 0 'checked on 3

AppActivate Application.Caption

End Sub

Private Sub CommandButton_0mm4_Click()
makeNice
Selection.NumberFormat = "0"" mm""" & ChrW(&H2074)
autoMerge 16, 0

AppActivate Application.Caption

End Sub

Private Sub CommandButton_RR_Click()

Dim txt As String
  
txt = "ROUND(,0)"

SetClipboard (txt)

End Sub

Private Sub CommandButton15_Click()

    If Len(ActiveCell.Formula) > 0 Then
       If Left(ActiveCell.Formula, 1) = "=" Then
           ActiveCell.Formula = "=round(" & Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1) & ",1)"
        End If
    End If

End Sub

Private Sub CommandButton_zero_Click()
makeNice
Selection.NumberFormat = "0"
autoMerge 3, 0


AppActivate Application.Caption
End Sub

Private Sub CommandButton_zerozero_Click()
makeNice
Selection.NumberFormat = "0.0"
autoMerge 8, 1
'8 for 20


AppActivate Application.Caption

End Sub

Private Sub CommandButton_zerozero_nomerge_Click()
makeNice
Selection.NumberFormat = "0.0"

AppActivate Application.Caption

End Sub

Private Sub CommandButton_zerozerozero_Click()
makeNice
Selection.NumberFormat = "0.00"
autoMerge 17, 2


AppActivate Application.Caption

End Sub

Private Sub CommandButton_zeta_Click()
Dim txt As String
txt = ChrW(&H3B6)
SetClipboard (txt)
End Sub

Private Sub CommandButton_gamma_Click()
Dim txt As String
txt = ChrW(&H3B3)
SetClipboard (txt)
End Sub

Private Sub CommandButton_zz_Click()
Dim txt As String
txt = ChrW(&H3BE)
SetClipboard (txt)
End Sub

Private Sub CommandButton_xx_Click()
Dim txt As String
txt = ChrW(&H3C7)
SetClipboard (txt)
End Sub


Private Sub CommandButton_mult_Click()

makeNice

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "M*x"
ActiveCell.Characters(start:=3, Length:=3).Font.Subscript = True

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate


makeNice


AppActivate Application.Caption
SendKeys "{F2}"

End Sub

Private Sub CommandButton_insert_row_Click()
Application.CutCopyMode = False ' clear clipboard
ActiveCell.Offset(0).EntireRow.Insert Shift:=xlUp, CopyOrigin:=xlFormatFromRightOrBelow

End Sub


Private Sub CommandButton_00percent_Click()
makeNice

If ActiveCell.Value = 1 Then 'for 100%
    Selection.NumberFormat = "0%"
Else
    Selection.NumberFormat = "0.0%"
End If

autoMerge 30, 0
AppActivate Application.Caption
End Sub

Private Sub CommandButton_lessequal_Click()
On Error GoTo eh
Dim Counter As Integer

zzzd:
If Counter < 20 Then 'to catch the loop if error arise
    If ActiveCell.Formula = "" And ActiveCell.MergeCells = False Then
        ActiveCell.FormulaR1C1 = "'+"
        If ActiveCell.Offset(0, 1).Formula = "" Then
            makeNice
            ActiveCell.NumberFormat = "@"
            ActiveCell.FormulaR1C1 = ChrW(&H2264)
            ActiveCell.Offset(0, 1).Activate
            makeNice
            ActiveCell.FormulaR1C1 = "="
            AppActivate Application.Caption
            SendKeys "{F2}"
        Else
            ActiveCell.Offset(0, 1).Activate
            AppActivate Application.Caption
        End If
    Counter = Counter + 1
    Else
begg:
        Counter = Counter + 1
        If ActiveCell.MergeCells = True Then
            ActiveCell.Offset(0, 1).Activate
            GoTo begg
        Else
            ActiveCell.Activate
            GoTo zzzd
        End If
    End If

Else
    GoTo Done
End If

Done:
    Exit Sub
eh:
    MsgBox ("err")
End

End Sub



Private Sub CommandButton_ix_Click()
makeNice
ActiveCell.Resize(7, 9).Merge
ActiveCell.UnMerge

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "Try"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 3).Merge

ActiveCell.Characters.Font.Color = _
RGB(0, 0, 255)
ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "200PFC"
ActiveCell.HorizontalAlignment = xlRight
ActiveCell.Offset(1, -1).Activate

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "Ix"
ActiveCell.Characters(start:=2, Length:=1).Font.Subscript = True

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.Resize(1, 4).Merge

Selection.NumberFormat = "0.00"" × 10^6 mm""" & ChrW(&H2074)
'print something useful
ActiveCell.FormulaR1C1 = "=IF(R[-1]C[-2]=" & Chr(34) & "200PFC" & Chr(34) & ",19.1,IF(R[-1]C[-2]= " & Chr(34) & "230PFC" & Chr(34) & ",26.8,IF(R[-1]C[-2]=" & Chr(34) & "250PFC" & Chr(34) & ",45.1,IF(R[-1]C[-2]=" & Chr(34) & "300PFC" & Chr(34) & ",72.4,IF(R[-1]C[-2]=" & Chr(34) & "200UB" & Chr(34) & ",15.8,IF(R[-1]C[-2]=" & Chr(34) & "250UB" & Chr(34) & ",35.4,IF(R[-1]C[-2]=" & Chr(34) & "310UB" & Chr(34) & ",63.2,IF(R[-1]C[-2]=" & Chr(34) & "360UB" & Chr(34) & ",121,IF(R[-1]C[-2]=" & Chr(34) & "410UB" & Chr(34) & ",188,IF(R[-1]C[-2]=" & Chr(34) & "460UB" & Chr(34) & ",296,IF(R[-1]C[-2]=" & Chr(34) & "530UB" & Chr(34) & ",477,IF(R[-1]C[-2]=" & Chr(34) & "610UB" & Chr(34) & ",761,IF(R[-1]C[-2]=" & Chr(34) & "200UC" & Chr(34) & ",45.9,IF(R[-1]C[-2]=" & Chr(34) & "250UC" & Chr(34) & ",114,IF(R[-1]C[-2]=" & Chr(34) & "310UC" & Chr(34) & ",223,0)))))))))))))))"

ActiveCell.Resize(1, 6).Merge

ActiveCell.Offset(-1, -2).Activate
AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_arrow_Click()
makeNice
ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = ChrW(&H2192)
ActiveCell.Offset(0, 1).Activate
makeNice
AppActivate Application.Caption
SendKeys "{F2}"
End Sub




Private Sub CommandButton_startup_Click()
On Error GoTo eh
'if variable not defined - in VBA Click on Tools-References in the VBE,
'and scroll down and tick the entry for Microsoft Visual Basic for Applications Extensibility 5.3.

'if access error 1004 - turn on in excel - Trust Center- macros - vba

''Create on load sub
With ActiveWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
   .InsertLines 1, "Private Sub Workbook_Open()"
   .InsertLines 2, "   SMC.Show"
   .InsertLines 3, "End Sub"
End With


''Add new module with code

Dim vbp As VBProject
Dim vbc As VBComponent
Dim strCode
Set vbp = Application.VBE.ActiveVBProject
Set vbc = vbp.VBComponents.Add(vbext_ct_StdModule)
vbc.Name = "Module1"
strCode = "Private Sub Workbook_Open()" & vbNewLine & "    SMC.Show" & vbNewLine & "End Sub"
vbc.CodeModule.AddFromString strCode

Done:
    Exit Sub
eh:
    MsgBox ("check recommendations in code")
    End
    
End Sub




Private Sub CommandButton_cantilever_MF_Click()
makeNice
ActiveCell.Resize(1, 21).Merge
ActiveCell.UnMerge

ActiveCell.NumberFormat = "@"
ActiveCell.FormulaR1C1 = "M*x"
ActiveCell.Characters(start:=3, Length:=3).Font.Subscript = True

ActiveCell.Offset(0, 2).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate

ActiveCell.FormulaR1C1 = "1"
Selection.NumberFormat = "0.0"" kN"""
ActiveCell.Resize(1, 3).Merge

ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "×"
ActiveCell.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 1).Activate

Selection.NumberFormat = "0.0"" m"""
ActiveCell.Resize(1, 2).Merge
ActiveCell.HorizontalAlignment = xlRight


ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "="
ActiveCell.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=R[0]C[-7]*R[0]C[-3]"
Selection.NumberFormat = "0.0"" kNm"""
autoMerge 38, 2
ActiveCell.Offset(0, -3).Activate
ActiveCell.FormulaR1C1 = "="
AppActivate Application.Caption
SendKeys "{F2}"
End Sub

Private Sub CommandButton_check_Click()
On Error GoTo eh

Dim x, y, z, tt, ff, zz, xx, vv As String
Dim f, mc As Range
Dim ss, gg, mm
Dim i As Integer

xx = ActiveCell.Formula2

If InStr(ActiveCell.Value, "NAME?") Then GoTo Done

ff = Replace(xx, "+", " ")
tt = Replace(ff, "=", " ")
ff = Replace(tt, "*", " ")
tt = Replace(ff, "/", " ")
ff = Replace(tt, "^", " ")
tt = Replace(ff, "-", " ")

ss = Split(tt, " ")

gg = UBound(ss) - LBound(ss)  'should be +1 but not in this case


'check for any mm
For i = 1 To gg:
    If Asc(Left(ss(i), 1)) > 64 And Asc(Left(ss(i), 1)) < 91 Then 'check is it link to a cell
        Set f = Application.Range(cell1:=ss(i))
        z = f.NumberFormat
        If InStr(z, " mm") > 0 And InStr(z, " mm²") = 0 And InStr(z, " mm³") = 0 And InStr(z, "0"" mm""" & ChrW(&H2074)) = 0 And InStr(z, "0.00"" × 10^6 mm""" & ChrW(&H2074)) = 0 Then  'position of substring
            xx = Replace(xx, ss(i), "(" & ss(i) + "/1000)")
        End If
      End If
ActiveCell.Value = xx
Next

'check for any kg
For i = 1 To gg:
    If Asc(Left(ss(i), 1)) > 64 And Asc(Left(ss(i), 1)) < 91 Then 'check is it link to a cell
        Set f = Application.Range(cell1:=ss(i))
        z = f.NumberFormat
        If InStr(z, " kg") > 0 Then   'position of substring
            xx = Replace(xx, ss(i), "(" & ss(i) + "/100)")
        End If
      End If
ActiveCell.Value = xx
Next

'check for any mm²
For i = 1 To gg:
    If Asc(Left(ss(i), 1)) > 64 And Asc(Left(ss(i), 1)) < 91 Then 'check is it link to a cell
        Set f = Application.Range(cell1:=ss(i))
        z = f.NumberFormat
        If InStr(z, " mm²") > 0 Then   'position of substring
            xx = Replace(xx, ss(i), "(" & ss(i) + "/1000000)")
        End If
      End If
ActiveCell.Value = xx
Next

'check for any mm³
For i = 1 To gg:
    If Asc(Left(ss(i), 1)) > 64 And Asc(Left(ss(i), 1)) < 91 Then 'check is it link to a cell
        Set f = Application.Range(cell1:=ss(i))
        z = f.NumberFormat
        If InStr(z, " mm³") > 0 Then   'position of substring
            xx = Replace(xx, ss(i), "(" & ss(i) + "/1000000000)")
        End If
      End If
ActiveCell.Value = xx
Next


'check for any mm^4
'"0"" mm""" & ChrW(&H2074)
For i = 1 To gg:
    If Asc(Left(ss(i), 1)) > 64 And Asc(Left(ss(i), 1)) < 91 Then 'check is it link to a cell
        Set f = Application.Range(cell1:=ss(i))
        z = f.NumberFormat
        'If InStr(z, ChrW(&H2074)) > 0 Then   'position of substring
        If z = "0"" mm""" & ChrW(&H2074) Then
            xx = Replace(xx, ss(i), "(" & ss(i) + "/1000000000000)")
        End If
      End If
ActiveCell.Value = xx
Next

'check for any
'"0.00"" × 10^6 mm""" & ChrW(&H2074)
For i = 1 To gg:
    If Asc(Left(ss(i), 1)) > 64 And Asc(Left(ss(i), 1)) < 91 Then 'check is it link to a cell
        Set f = Application.Range(cell1:=ss(i))
        z = f.NumberFormat
        'If InStr(z, ChrW(&H2074)) > 0 Then   'position of substring
        If z = "0.00"" × 10^6 mm""" & ChrW(&H2074) Then
            xx = Replace(xx, ss(i), "(" & ss(i) + "/1000000)")
        End If
      End If
ActiveCell.Value = xx
Next

'check for any MPa
For i = 1 To gg:
    If Asc(Left(ss(i), 1)) > 64 And Asc(Left(ss(i), 1)) < 91 Then 'check is it link to a cell
        Set f = Application.Range(cell1:=ss(i))
        z = f.NumberFormat
        If InStr(z, " MPa") > 0 Then   'position of substring
            xx = Replace(xx, ss(i), "(" & ss(i) + "*1000)")
        End If
      End If
   ActiveCell.Value = xx
Next

'check for any GPa
For i = 1 To gg:
    If Asc(Left(ss(i), 1)) > 64 And Asc(Left(ss(i), 1)) < 91 Then 'check is it link to a cell
        Set f = Application.Range(cell1:=ss(i))
        z = f.NumberFormat
        If InStr(z, " GPa") > 0 Then   'position of substring
            xx = Replace(xx, ss(i), "(" & ss(i) + "*1000000)")
        End If
      End If
   ActiveCell.Value = xx
Next


Done:
    Exit Sub
eh:
    End

End Sub



Private Sub CommandButton_kN_Click()
On Error GoTo eh

Dim x, y As String

''''MsgBox ("kN start")
x = ActiveCell.Value2
'''MsgBox ("!!!!!" & X)
If x = "" Then
    y = "do nothing"
Else
    If InStr(x, ".") > 1 Then
        If Len(Split(x, ".")(0)) >= 3 Then
           CommandButton_0kN_Click
        ElseIf Len(Split(x, ".")(0)) = 2 Then
            makeNice
            Selection.NumberFormat = "0.0"" kN"""
            autoMerge 20, 1 'checked 10
            AppActivate Application.Caption
        ElseIf Len(Split(x, ".")(0)) = 1 And Len(Split(x, ".")(0)) = 1 Then
            makeNice
            Selection.NumberFormat = "0.0"" kN"""
            autoMerge 20, 1 'checked 0.5 1.5
            AppActivate Application.Caption
        'ElseIf Len(Split(x, ".")(0)) = 1 Then
        '    CommandButton_00kN_Click
        End If
    Else
        ''''MsgBox ("else")
        CommandButton_0kN_Click
    End If
End If


Done:
    Exit Sub
eh:
    MsgBox ("Check cells")

End Sub

Private Sub CommandButton_MPa_Click()
Dim x, y As String
''''MsgBox ("kN start")
x = ActiveCell.Value2
'''MsgBox ("!!!!!" & X)
If x = "" Then
    y = "do nothing"
Else
    If InStr(x, ".") > 1 Then
        If Len(Split(x, ".")(0)) >= 3 Then
            CommandButton_0MPa_Click
        ElseIf Len(Split(x, ".")(0)) = 2 Then
            CommandButton_0MPa_Click
        ElseIf Len(Split(x, ".")(0)) = 1 And Split(x, ".")(0) = 0 Then
            CommandButton_00MPa_Click
        ElseIf Len(Split(x, ".")(0)) = 1 Then
            CommandButton_000MPa_Click
        End If
    Else
       CommandButton_0MPa_Click
    End If
End If
End Sub

Private Sub CommandButton_GPa_Click()
Dim x, y As String
''''MsgBox ("kN start")
x = ActiveCell.Value2
'''MsgBox ("!!!!!" & X)
If x = "" Then
    y = "do nothing"
Else
    If InStr(x, ".") > 1 Then
        If Len(Split(x, ".")(0)) >= 3 Then
           CommandButton_0GPa_Click
        ElseIf Len(Split(x, ".")(0)) = 2 Then
            CommandButton_00GPa_Click
        ElseIf Len(Split(x, ".")(0)) = 1 And Split(x, ".")(0) = 0 Then
            CommandButton_00GPa_Click
        ElseIf Len(Split(x, ".")(0)) = 1 Then
            CommandButton_00GPa_Click
        End If
    Else
       CommandButton_0GPa_Click
    End If
End If
End Sub

Private Sub CommandButton_m_Click()
Dim x, y As String

x = ActiveCell.Value

If x = "" Then
    y = "do nothing"
Else
    If InStr(x, ".") > 1 Then
        If Len(Split(x, ".")(0)) >= 3 Then
           CommandButton_0m_Click
        ElseIf Len(Split(x, ".")(0)) = 2 Then
            CommandButton_00m_Click
        ElseIf Len(Split(x, ".")(0)) = 1 And Split(x, ".")(0) = 0 And Len(Split(x, ".")(1)) = 2 Then
            CommandButton_000m_Click
        ElseIf Len(Split(x, ".")(0)) = 1 And Split(x, ".")(0) = 0 Then
            CommandButton_00m_Click
        ElseIf Len(Split(x, ".")(0)) = 1 Then
            CommandButton_00m_Click
        End If
    Else
        CommandButton_0m_Click
    End If
End If

End Sub

    Private Sub CommandButton_m2_Click()
Dim x, y As String
x = ActiveCell.Value
If x = "" Then
    y = "do nothing"
Else
    If InStr(x, ".") > 1 Then
        If Len(Split(x, ".")(0)) >= 3 Then
           CommandButton_0m2_Click
        ElseIf Len(Split(x, ".")(0)) = 2 Then
            CommandButton_00m2_Click
        ElseIf Len(Split(x, ".")(0)) = 1 And Split(x, ".")(0) = 0 Then
            CommandButton_00m2_Click
        ElseIf Len(Split(x, ".")(0)) = 1 Then
            CommandButton_00m2_Click
        End If
    Else
        CommandButton_0m2_Click
    End If
End If
End Sub


Private Sub CommandButton2_Click()
On Error GoTo eh
Dim cell As Range

For Each cell In Selection
    cell.Formula2 = Replace(cell.Formula2, "$", "")
Next cell

Done:
    Exit Sub
eh:
    End

End Sub





Private Sub CommandButton3_Click()

On Error GoTo eh

ActiveCell.Delete Shift:=xlToLeft
AppActivate Application.Caption

Done:
    Exit Sub
eh:
    End
End Sub

Private Sub CommandButton4_Click()
On Error GoTo eh
Dim chromePath As String
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Shell (chromePath & " -url https://www.dropbox.com/s/4qtq6v2bqvbi0z2/SESOC%20-%20Simplified_Design_of_Steel_Members_page63-67.pdf?dl=0")
Done:
    Exit Sub
eh:
    MsgBox ("check file address or Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End
    
End Sub

Private Sub CommandButton5_Click()
On Error GoTo eh
Dim chromePath As String
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Shell (chromePath & " -url https://www.dropbox.com/s/hm82e3auheiakxb/SESOC%20-%20Simplified_Design_of_Steel_Members_page69-76.pdf?dl=0")

Done:
    Exit Sub
eh:
    MsgBox ("check file address or Chrome address C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    End

End Sub

Private Sub CommandButton6_Click()
ActiveCell.Insert xlShiftToRight
AppActivate Application.Caption
End Sub

Private Sub CommandButton7_Click()
'Allows you to write formula like 5kn*6
Dim xv, yg, zt, tg, fd, zj, xm, vm, Result, Result2 As String
Dim Counter As Integer

'On Error GoTo Done

xv = ActiveCell.Formula

xv = Replace(xv, "=", "")

If xv = "" Then GoTo Done
ActiveCell.Offset(1, 0).Select

start:
'''''''MsgBox ("xv in beginning = " & xv)
If Counter > 20 Then GoTo Done

Counter = Counter + 1

Dim i, StringLength As Integer
StringLength = Len(xv)
Result = ""
'MsgBox (StringLength)

For i = 1 To StringLength Step 1
    If IsNumeric(Mid(xv, i, 1)) Or Mid(xv, i, 1) = "." Then
        Result = Result & Mid(xv, i, 1)
        'MsgBox (Result)
    Else
        GoTo NextStep
    End If
Next i

NextStep:
'''''''MsgBox ("numbers = " & Result)

ActiveCell.Formula = Result

xv = Split(xv, Result, 2)(1)
Result2 = ""
StringLength = Len(xv)
For i = 1 To StringLength Step 1
    If Asc(Mid(xv, i, 1)) > 64 And Asc(Mid(xv, i, 1)) < 91 Or Asc(Mid(xv, i, 1)) > 96 And Asc(Mid(xv, i, 1)) < 123 Then  'upper and lowercase
        Result2 = Result2 & Mid(xv, i, 1)
    Else
        GoTo NextStep2
    End If
Next i

NextStep2:
'''''''MsgBox ("units = " & Result2)

GetUnitsFromString (Result2)

If xv = Result2 Then
   GoTo ToExecution
Else
    If Result2 <> "" Then
        xv = Split(xv, Result2, 2)(1)
'''''''MsgBox ("xv = " & xv & " Result = " & Result & " Result2 = " & Result2)
    End If
End If

'''''''MsgBox ("Sign is " & Mid(xv, 1, 1))
If Mid(xv, 1, 1) = "*" Then
    xv = Split(xv, "*", 2)(1)
    CommandButtonx_Click
ElseIf Mid(xv, 1, 1) = "/" Then
    xv = Split(xv, "/", 2)(1)
    CommandButton_divided_Click
ElseIf Mid(xv, 1, 1) = "-" Then
    xv = Split(xv, "-", 2)(1)
    CommandButton_minus_Click
ElseIf Mid(xv, 1, 1) = "+" Then
    xv = Split(xv, "+", 2)(1)
    CommandButton_plus_Click
End If
''and LOOP
GoTo start

ToExecution:
'''''''MsgBox ("=")
Call CommandButton_equal_Click
ActiveCell.Formula = ""
'''''''MsgBox ("aaand")
Call CommandButton_check_kN_Click


Done:

End Sub

Private Sub CommandButton8_Click()
'get 3 actual numbers like 1/3=0.333
Dim x, yn, gn As String
Dim Counter, Counter2 As Integer
Dim i, StringLength As Integer
Dim trigger As Boolean

makeNice
If ActiveCell.NumberFormat = "General" Then
'check to avoid 1E+10 format
     ActiveCell.NumberFormat = "0.00"
End If


start:
'reset start

Counter2 = 0
While Counter2 < 15 ' just to be sure we have no decimals for the start
    Application.CommandBars.FindControl(ID:=399).Execute
    Counter2 = Counter2 + 1
Wend
'reset end

x = ActiveCell.Value()
If x = "" Then GoTo Done

If InStr(x, "E") > 0 And InStr(x, "E+") = 0 Then
    yn = Split(x, "E-")(0)
    gn = CInt(Split(x, "E-")(1))
    Counter = 2 + gn
    While Counter > 0
        Application.CommandBars.FindControl(ID:=398).Execute
        Counter = Counter - 1
    Wend
End If


If InStr(x, "E") = 0 And Len(x) <> 1 Then
    StringLength = Len(x)
    Counter = 0
    trigger = False
    'trigger means should we count 0 as actual number or not
    'it's True for 0.203 for exaple
    
    Counter = Len(Split(x, ".")(0))
        
    For i = 1 To StringLength Step 1
        If Mid(x, i, 1) <> "." And Counter < 3 And Mid(x, i, 1) <> "0" Then
            trigger = True
            Application.CommandBars.FindControl(ID:=398).Execute
            Counter = Counter + 1
        ElseIf Mid(x, i, 1) = "0" And Counter < 3 And trigger = False Then
            Application.CommandBars.FindControl(ID:=398).Execute
        End If
    'MsgBox ("number is " & Mid(x, i, 1) & " , trigger= " & trigger & " , counter = " & Counter)
    Next i
End If


If Len(x) = 1 Then 'for simple cases
    Application.CommandBars.FindControl(ID:=398).Execute
    Application.CommandBars.FindControl(ID:=398).Execute
End If


endd: 'merging
Dim zz, xx As Integer
Dim pin2 As Boolean
pin2 = False
If InStr(ActiveCell.Text, "#") > 0 Then
pam:
    If ActiveCell.MergeCells = True Then
        pin2 = True
        zz = zz + 1
        ActiveCell.Offset(0, 1).Select
        GoTo pam
    Else
        If zz > 0 Then
            ActiveCell.Offset(0, -zz).Select
            'MsgBox ("a")
        End If
        
        If ActiveCell.Offset(0, 1) <> "" Then
            ActiveCell.Offset(0, 1).Insert xlShiftToRight
            
            If pin2 = True Then
                ActiveCell.Offset(0, 1).Select
            End If
            
            'MsgBox ("b")
            GoTo pam
        End If
            
        If zz > 0 Then 'because of GoTo pam and because we use zz in other places
            xx = zz - 1
        Else
            xx = zz
        End If
        
        
        'MsgBox ("merge")
        ActiveCell.Resize(1, 2 + xx).Merge
        
    End If
    GoTo start
End If

Done:
End Sub



Function GetTitle(sheet_ As Worksheet, cell_address As String, numb As String) As Worksheet
'first is where to start
'adress is for title sheet
'numb is the number on the left
Dim tst, mmt As String

exwhile6:
If sheet_.Index + 1 <= Application.Sheets.Count Then
    If Worksheets(sheet_.Index + 1).Visible = 0 Then
        Set sheet_ = Worksheets(sheet_.Index + 1)
        'MsgBox (sheet_.Name)
        GoTo exwhile6
    Else:
        Set sheet_ = Worksheets(sheet_.Index + 1)
        'MsgBox (sheet_.Name)
        tst = "=" & "'" & sheet_.Range("D2").Parent.Name & "'" & "!" & sheet_.Range("D2").Address(0)
        tst = Replace(tst, "$", "")
        Sheet00.Range(cell_address).Formula = tst
        Sheet00.Range(cell_address).Offset(0, -1).Formula = "–"

    End If
End If

If Sheet00.Range(cell_address).Text <> "" Then
    Sheet00.Range(cell_address).Offset(0, -3).Formula = numb
End If

'and get back numbers
mmt = "=" & "'" & Sheet00.Range(cell_address).Offset(0, -3).Parent.Name & "'" & "!" & Sheet00.Range(cell_address).Offset(0, -3).Address(0)
mmt = Replace(mmt, "$", "")
mmt = Replace(mmt, "=", "")
sheet_.Range("B2").NumberFormat = "General"

sheet_.Range("B2").Formula = "=CONCATENATE(" & mmt & "," & Chr(34) & "." & Chr(34) & ")" 'to get a dot after the number

sheet_.Range("B2").HorizontalAlignment = xlLeft
sheet_.Range("B2").VerticalAlignment = xlVAlignCenter
sheet_.Range("B2").Font.Name = "Arial"
sheet_.Range("B2").Font.Size = 12
sheet_.Range("B2").Font.Color = _
RGB(0, 0, 0)
sheet_.Range("B2").Font.Bold = True

Set GetTitle = sheet_

End Function



Private Sub CommandButton9_Click()

On Error GoTo eh

Dim sheet_ As Worksheet

Set sheet_ = Sheet00

Dim x As Worksheet
Dim m As String

Set x = GetTitle(sheet_, "E35", "1")
Set x = GetTitle(x, "E36", "2")
Set x = GetTitle(x, "E37", "3")
Set x = GetTitle(x, "E38", "4")
Set x = GetTitle(x, "E39", "5")
Set x = GetTitle(x, "E40", "6")
Set x = GetTitle(x, "E41", "7")
Set x = GetTitle(x, "E42", "8")
Set x = GetTitle(x, "E43", "9")
Set x = GetTitle(x, "E44", "10")
Set x = GetTitle(x, "E45", "11")
Set x = GetTitle(x, "E46", "12")
Set x = GetTitle(x, "E47", "13")
Set x = GetTitle(x, "E48", "14")
Set x = GetTitle(x, "E49", "15")
Set x = GetTitle(x, "E50", "16")
Set x = GetTitle(x, "E51", "17")
Set x = GetTitle(x, "E52", "18")
Set x = GetTitle(x, "E53", "19")
Set x = GetTitle(x, "E54", "20")
Set x = GetTitle(x, "E55", "21")

Done:
    Exit Sub

eh:
MsgBox ("check merged cells at the bottom")

End Sub

Private Sub concrete_beam_Click()
On Error GoTo eh
Sheets("MISC").Range("B1054:AQ1131").Copy

ActiveSheet.Paste

Application.CutCopyMode = False

Done:
    Exit Sub
eh:
    MsgBox ("Can't find sheet MISC")
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label10_Click()

End Sub

Private Sub steel_beam_Click()
On Error GoTo eh
Sheets("MISC").Range("B1717:V1794").Copy 'AQ instead of V if required

ActiveSheet.Paste

Application.CutCopyMode = False

Done:
    Exit Sub
eh:
    MsgBox ("Can't find sheet MISC")
End Sub

Private Sub timber_beam_Click()
On Error GoTo eh

Sheets("MISC").Range("B391:AQ507").Copy

ActiveSheet.Paste

Application.CutCopyMode = False

Done:
    Exit Sub
eh:
    MsgBox ("Can't find sheet MISC")

End Sub


Private Sub CommandButtonx_Click()
On Error GoTo eh
Dim Counter As Integer

begg:
If Counter < 20 Then 'to catch the loop if error arise
   Counter = Counter + 1
        If ActiveCell.MergeCells = True Then
            ActiveCell.Offset(0, 1).Select
            GoTo begg
        Else
            ActiveCell.Select
            GoTo zzzd
        End If
    
zzzd:
    If ActiveCell.Formula = "" Then
        ActiveCell.FormulaR1C1 = "'×"
        If ActiveCell.Offset(0, 1).Formula = "" Then
            makeNice
            ActiveCell.HorizontalAlignment = xlCenter
            ActiveCell.Offset(0, 1).Activate
            makeNice
            ActiveCell.FormulaR1C1 = "="
            AppActivate Application.Caption
            SendKeys "{F2}"

        Else
            ActiveCell.Offset(0, 1).Activate
            AppActivate Application.Caption
        End If
        Counter = Counter + 1
    
    ElseIf ActiveCell.FormulaR1C1 = "×" Then
        ActiveCell.Offset(0, 1).Activate
        AppActivate Application.Caption
        SendKeys "{F2}"
        GoTo Done
    Else
        ActiveCell.Offset(0, 1).Activate
        GoTo begg
    End If

Else
    GoTo Done
End If

Done:
    Exit Sub
eh:
    MsgBox ("err")
End

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label9_Click()

End Sub


Private Sub timber_column_Click()
On Error GoTo eh
Sheets("MISC").Range("B547:V624").Copy 'AQ instead of V if needed

ActiveSheet.Paste

Application.CutCopyMode = False
Done:
    Exit Sub
eh:
    MsgBox ("Can't find sheet MISC")

End Sub

Private Sub UserForm_Click()

End Sub
