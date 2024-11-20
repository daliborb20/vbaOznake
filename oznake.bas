Attribute VB_Name = "oznake"
Sub Zuto()
Attribute Zuto.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+Shift+S
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub Crveno()
Attribute Crveno.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' Macro3 Macro
'
' Keyboard Shortcut: Ctrl+Shift+A
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Sub redZuti()
Attribute redZuti.VB_ProcData.VB_Invoke_Func = "L\n14"
Dim oblast As Range
With Selection
    For Each cell In Selection
        Selection.EntireRow.Interior.Color = RGB(255, 255, 102)
    Next cell
End With


End Sub




Sub redCrveni()
Attribute redCrveni.VB_ProcData.VB_Invoke_Func = "K\n14"
Dim oblast As Range
With Selection
    For Each cell In Selection
        Selection.EntireRow.Interior.Color = RGB(255, 102, 153)
    Next cell
End With
End Sub

Sub redZeleni()
Attribute redZeleni.VB_ProcData.VB_Invoke_Func = "O\n14"
Dim oblast As Range
With Selection
    For Each cell In Selection
        Selection.EntireRow.Interior.Color = RGB(204, 255, 153)
    Next cell
End With

End Sub
