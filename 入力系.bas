Attribute VB_Name = "入力系"
Option Explicit
Sub 主キー取得(鍵1 As Variant, 鍵2 As Variant, 鍵3 As Variant)
    Dim 終行 As Long, 行 As Long
    With Sheets("台帳転記設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        For 行 = 2 To 終行
            If .Cells(行, 3) <> "" Then
                If 鍵1(1) = "" Then
                    鍵1(1) = .Cells(行, 1)
                    鍵1(2) = .Cells(行, 2)
                Else
                    If 鍵2(1) = "" Then
                        鍵2(1) = .Cells(行, 1)
                        鍵2(2) = .Cells(行, 2)
                    Else
                        If 鍵3(1) = "" Then
                            鍵3(1) = .Cells(行, 1)
                            鍵3(2) = .Cells(行, 2)
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub
Sub 入力フォームクリア()
    Dim 鍵1(1 To 2), 鍵2(1 To 2), 鍵3(1 To 2)
    Dim 終行 As Long, 行 As Long
    Call 主キー取得(鍵1, 鍵2, 鍵3)
    With Sheets("台帳転記設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim 項目リスト(2 To 終行)
        For 行 = 2 To 終行
            項目リスト(行) = .Cells(行, 1)
        Next
    End With
    With Sheets("入力フォーム")
        .Unprotect
        For 行 = 2 To 終行
            Select Case 項目リスト(行)
                Case 鍵1(1), 鍵2(1), 鍵3(1)
                Case Else: .Range(項目リスト(行)).MergeArea.ClearContents
            End Select
        Next
        .Protect
    End With
End Sub
Function 転記行検索(台帳最下行 As Long) As Long
    Dim 終行 As Long, 行 As Long
    Dim 検索鍵 As String
    Dim 鍵1(1 To 2), 鍵2(1 To 2), 鍵3(1 To 2)
    If 台帳最下行 < 2 Then Exit Function
    ReDim 行配列(2 To 台帳最下行)
    Call 主キー取得(鍵1, 鍵2, 鍵3)
    With Sheets("入力フォーム")
        If 鍵1(1) <> "" Then 検索鍵 = .Range(鍵1(1))
        If 鍵2(1) <> "" Then 検索鍵 = 検索鍵 & "-" & .Range(鍵2(1))
        If 鍵3(1) <> "" Then 検索鍵 = 検索鍵 & "-" & .Range(鍵3(1))
    End With
    With Sheets("管理台帳")
        For 行 = 2 To 台帳最下行
            If 鍵1(1) <> "" Then 行配列(行) = .Cells(行, 鍵1(2))
            If 鍵2(1) <> "" Then 行配列(行) = 行配列(行) & "-" & .Cells(行, 鍵2(2))
            If 鍵3(1) <> "" Then 行配列(行) = 行配列(行) & "-" & .Cells(行, 鍵3(2))
        Next
        For 行 = 2 To 台帳最下行
            If 検索鍵 = 行配列(行) Then
                転記行検索 = 行
                .Range("_選択行") = 行
                Exit For
            End If
        Next
    End With
End Function
Function 編集差分確認() As String
    Dim 終行 As Long, 行 As Long, 最右列 As Long, 列 As Long, 台帳最下行 As Long, 記録行 As Long
    With Sheets("台帳転記設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim 設定(1 To 終行 - 1, 1 To 2)
        For 行 = 2 To 終行
            設定(行 - 1, 1) = .Cells(行, 1)
            設定(行 - 1, 2) = .Cells(行, 2)
            If 最右列 < 設定(行 - 1, 2) Then 最右列 = 設定(行 - 1, 2)
        Next
    End With
    With Sheets("入力フォーム")
        ReDim 配列(1 To 1, 1 To 最右列)
        For 行 = 1 To UBound(設定, 1)
            配列(1, 設定(行, 2)) = .Range(設定(行, 1))
        Next
    End With
    With Sheets("管理台帳")
        For 列 = 1 To 最右列
            If 台帳最下行 < .Cells(Rows.Count, 列).End(xlUp).Row Then
                台帳最下行 = .Cells(Rows.Count, 列).End(xlUp).Row
            End If
        Next
        記録行 = 転記行検索(台帳最下行)
        If 記録行 = 0 Then Exit Function
        For 列 = 1 To 最右列
            If 配列(1, 列) <> .Cells(記録行, 列) Then
                編集差分確認 = "差分あり"
                Exit For
            End If
        Next
    End With
End Function
Sub 登録更新()
    Dim 終行 As Long, 行 As Long, 最右列 As Long, 列 As Long, 台帳最下行 As Long, 記録行 As Long
    Dim 文 As String
    With Sheets("台帳転記設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim 設定(1 To 終行 - 1, 1 To 2)
        For 行 = 2 To 終行
            設定(行 - 1, 1) = .Cells(行, 1)
            設定(行 - 1, 2) = .Cells(行, 2)
            If 最右列 < 設定(行 - 1, 2) Then 最右列 = 設定(行 - 1, 2)
        Next
    End With
    With Sheets("入力フォーム")
        ReDim 配列(1 To 1, 1 To 最右列)
        For 行 = 1 To UBound(設定, 1)
            配列(1, 設定(行, 2)) = .Range(設定(行, 1))
        Next
    End With
    With Sheets("管理台帳")
        For 列 = 1 To 最右列
            If 台帳最下行 < .Cells(Rows.Count, 列).End(xlUp).Row Then
                台帳最下行 = .Cells(Rows.Count, 列).End(xlUp).Row
            End If
        Next
        記録行 = 転記行検索(台帳最下行)
        If 記録行 = 0 Then
            記録行 = 台帳最下行 + 1
            文 = "新規登録してよろしいですか？"
        Else: 文 = "台帳を更新してよろしいですか？"
        End If
        If MsgBox(文, vbYesNo) = vbYes Then
            .Cells(記録行, 1).Resize(1, 最右列) = 配列
            MsgBox "台帳記録が完了しました"
        End If
    End With
End Sub
Sub 台帳戻し(記録行 As Long)
    Dim 終行 As Long, 行 As Long, 最右列 As Long, 列 As Long, 台帳最下行 As Long
    Application.ScreenUpdating = False
    With Sheets("台帳転記設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim 設定(1 To 終行 - 1, 1 To 2)
        For 行 = 2 To 終行
            設定(行 - 1, 1) = .Cells(行, 1)
            設定(行 - 1, 2) = .Cells(行, 2)
            If 最右列 < 設定(行 - 1, 2) Then 最右列 = 設定(行 - 1, 2)
        Next
    End With
    With Sheets("管理台帳")
        For 列 = 1 To 最右列
            If 台帳最下行 < .Cells(Rows.Count, 列).End(xlUp).Row Then
                台帳最下行 = .Cells(Rows.Count, 列).End(xlUp).Row
            End If
        Next
        If 記録行 = 0 Then 記録行 = 転記行検索(台帳最下行)
        If 記録行 = 0 Then
            Call 入力フォームクリア
            Exit Sub
        End If
        ReDim 配列(1 To 1, 1 To 最右列)
        For 列 = 1 To 最右列
            配列(1, 列) = .Cells(記録行, 列)
        Next
    End With
    With Sheets("入力フォーム")
        Call 入力フォームクリア
        .Unprotect
        Application.EnableEvents = False
        For 行 = 1 To UBound(設定, 1)
            .Range(設定(行, 1)) = 配列(1, 設定(行, 2))
        Next
        Application.EnableEvents = True
        .Protect
    End With
    Application.ScreenUpdating = True
End Sub
