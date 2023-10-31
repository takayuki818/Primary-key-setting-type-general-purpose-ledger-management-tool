Attribute VB_Name = "その他"
Option Explicit
Sub 保護切替()
    With ActiveSheet
        Select Case .ProtectContents
            Case True: .Unprotect: MsgBox "シート保護を解除しました"
            Case False: .Protect: MsgBox "シートを保護しました"
        End Select
    End With
End Sub
Sub 全シート展開()
    Dim シート As Worksheet
    Application.ScreenUpdating = False
    For Each シート In Sheets
        シート.Visible = True
    Next
    Application.ScreenUpdating = True
End Sub
Sub 書式再設定()
    Dim 最下行 As Long, 最右列 As Long, 列 As Long
    With Sheets("管理台帳")
        最右列 = .Cells(1, 1).End(xlToRight).Column
        最下行 = 2
        For 列 = 1 To 最右列
            If 最下行 < .Cells(Rows.Count, 列).End(xlUp).Row Then
                最下行 = .Cells(Rows.Count, 列).End(xlUp).Row
            End If
        Next
        Call 台帳書式設定(最下行 + 100, 最右列)
        Application.EnableEvents = False
        .Activate
        Application.EnableEvents = True
        MsgBox "「台帳」シートの書式を再設定しました" & vbCrLf & "※最下行+100行まで書式設定行延長"
    End With
End Sub
Sub 台帳書式設定(最下行 As Long, 最右列 As Long)
    Dim 条件 As FormatCondition
    With Sheets("管理台帳")
        .Range("A2").Resize(Rows.Count - 1, Columns.Count).Borders.LineStyle = False
        .Range("A2").Resize(最下行 - 1, 最右列).Borders.LineStyle = True
        .Cells.FormatConditions.Delete
        Set 条件 = .Range("A2").Resize(最下行 - 1, 最右列).FormatConditions.Add(Type:=xlExpression, Formula1:="=ROW()=_選択行")
        条件.Interior.Color = RGB(217, 225, 242)
    End With
End Sub
