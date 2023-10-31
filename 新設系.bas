Attribute VB_Name = "新設系"
Option Explicit
Sub ツール新規作成()
    Dim 終行 As Long, 行 As Long, 最右列 As Long, 列 As Long
    Dim 定義名 As String
    If MsgBox("項目設定を元に新しいツールベースを作成します" & vbCrLf & vbCrLf & "よろしいですか？", vbYesNo) <> vbYes Then Exit Sub
    Call 全シート展開
    Call 定義名全削除
    With Sheets("台帳転記設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim 項目リスト(1 To 終行 - 1, 1 To 1)
        For 行 = 2 To 終行
            項目リスト(行 - 1, 1) = .Cells(行, 1)
            定義名 = 項目リスト(行 - 1, 1)
            Call 名前の定義追加(定義名, "=入力フォーム!R" & 行 & "C2")
            Call 台帳列関数設定(.Cells(行, 2), 行 - 1)
        Next
        .Buttons.Delete
    End With
    With Sheets("入力フォーム")
        .Range("A2").Resize(Rows.Count - 1, 2).ClearContents
        .Range("A2").Resize(Rows.Count - 1, 2).Borders.LineStyle = False
        .Range("A2").Resize(Rows.Count - 1, 2).Interior.ColorIndex = 0
        .Range("A2").Resize(UBound(項目リスト, 1), 1) = 項目リスト
        .Range("B2").Resize(UBound(項目リスト, 1), 1).Locked = False
        .Range("B2").Resize(UBound(項目リスト, 1), 1).Interior.ColorIndex = 6
        .Range("A2").Resize(UBound(項目リスト, 1), 2).Borders.LineStyle = True
        .Activate
    End With
    With Sheets("管理台帳")
        最右列 = 終行 - 1
        ReDim 台帳項目(1 To 1, 1 To 最右列)
        For 列 = 1 To 最右列
            台帳項目(1, 列) = 項目リスト(列, 1)
        Next
        .Range("A1").Resize(Rows.Count, Columns.Count).ClearContents
        .Range("A1").Resize(Rows.Count, Columns.Count).Borders.LineStyle = False
        .Range("A1").Resize(Rows.Count, Columns.Count).Interior.ColorIndex = 0
        .Range("A1").Resize(1, 最右列) = 台帳項目
        .Range("A1").Resize(1, 最右列).Borders.LineStyle = True
        .Range("A1").Resize(1, 最右列).Interior.Color = RGB(142, 169, 219)
        Range(.Columns(1), .Columns(最右列)).AutoFit
        .Cells(1, 最右列 + 2) = "選択行"
        .Cells(1, 最右列 + 2).Resize(1, 2).Borders.LineStyle = True
        .Cells(1, 最右列 + 2).Interior.ColorIndex = 6
        Call 名前の定義追加("_選択行", "=管理台帳!R1C" & 最右列 + 3)
        Call 台帳書式設定(101, 最右列)
    End With
    With Sheets("印刷様式")
        .Range("A1").Resize(Rows.Count, 2).ClearContents
        .Range("A1").Resize(Rows.Count, 2).Interior.ColorIndex = 0
        .Range("A1") = "↓着色セル＝印刷様式パーツ"
        .Range("A2").Resize(UBound(項目リスト, 1), 1) = 項目リスト
        For 行 = 2 To 終行
            Call 印刷様式参照関数設定(.Cells(行, 2), Sheets("台帳転記設定").Cells(行, 2))
        Next
        .Range("B2").Resize(UBound(項目リスト, 1), 1).Interior.ColorIndex = 6
    End With
End Sub
Sub 定義名全削除()
    Dim 定義名 As Name
    On Error Resume Next
    For Each 定義名 In ThisWorkbook.Names
        If Left(定義名.Name, 1) <> "_" Then 定義名.Delete
    Next
    On Error GoTo 0
End Sub
Sub 名前の定義追加(定義名 As String, セル範囲 As String)
    On Error GoTo エラー時処理
    ThisWorkbook.Names.Add Name:=定義名, RefersToR1C1:=セル範囲
    Exit Sub
エラー時処理:
    Dim エラー文 As String
    Select Case Err.Number
        Case 1004
            エラー文 = "不正な項目名（定義名）が含まれています（対象項目名：" & 定義名 & "）"
            エラー文 = エラー文 & vbCrLf & vbCrLf & "項目名（定義名）が以下の条件を満たしているかご確認ください"
            エラー文 = エラー文 & vbCrLf & "　・名前の先頭は英文字、かな、カナ、漢字、アンダースコア（_）のいずれかである"
            エラー文 = エラー文 & vbCrLf & "　・空白等の無効な文字が含まれていない"
            エラー文 = エラー文 & vbCrLf & "　・既存の名前と競合していない"
            MsgBox エラー文
            Call 定義名全削除
            End
    End Select
End Sub
Sub 台帳列関数設定(セル As Range, 列 As Long)
    セル.FormulaR1C1 = "=COLUMN(管理台帳!C" & 列 & ")"
End Sub
Sub 印刷様式参照関数設定(セル As Range, 列番号 As Long)
    セル.FormulaR1C1 = "=INDIRECT(ADDRESS(_選択行,COLUMN(管理台帳!C" & 列番号 & "),,,""管理台帳""))"
End Sub
