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
Sub 台帳並替()
    Dim 最下行 As Long, 最右列 As Long, 列 As Long
    Dim 鍵1(1 To 2), 鍵2(1 To 2), 鍵3(1 To 2)
    With Sheets("管理台帳")
        最右列 = .Cells(1, Columns.Count).End(xlToLeft).Column
        For 列 = 1 To 最右列
            If 最下行 < .Cells(Rows.Count, 列).End(xlUp).Row Then
                最下行 = .Cells(Rows.Count, 列).End(xlUp).Row
            End If
        Next
        .Cells(1, 1).Resize(最下行, 最右列).Characters.PhoneticCharacters = ""
        Call 主キー取得(鍵1, 鍵2, 鍵3)
        With .Sort
            With .SortFields
                .Clear
                If 鍵1(1) <> "" Then
                    .Add Key:=Cells(1, 鍵1(2)), Order:=xlAscending
                End If
                If 鍵2(1) <> "" Then
                    .Add Key:=Cells(1, 鍵2(2)), Order:=xlAscending
                End If
                If 鍵3(1) <> "" Then
                    .Add Key:=Cells(1, 鍵3(2)), Order:=xlAscending
                End If
            End With
            .SetRange Range(Cells(1, 1), Cells(最下行, 最右列))
            .Header = xlYes
            .Apply
        End With
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
Sub 入力フォームオールクリア()
    Dim 終行 As Long, 行 As Long
    With Sheets("台帳転記設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim 項目リスト(2 To 終行)
        For 行 = 2 To 終行
            項目リスト(行) = .Cells(行, 1)
        Next
    End With
    With Sheets("入力フォーム")
        .Unprotect
        Application.EnableEvents = False
        For 行 = 2 To 終行
            Select Case 項目リスト(行)
                Case Else: .Range(項目リスト(行)).MergeArea.ClearContents
            End Select
        Next
        Application.EnableEvents = True
        .Protect
    End With
End Sub
Function 転記行検索() As Variant '主キーの値が未入力→「空文字列」を返す/台帳未登録→「0」を返す
    Dim 最右列 As Long, 列 As Long, 最下行 As Long, 行 As Long
    Dim 鍵1(1 To 2), 鍵2(1 To 2), 鍵3(1 To 2)
    Dim 検索鍵 As String
    With Sheets("入力フォーム")
        Call 主キー取得(鍵1, 鍵2, 鍵3)
        If 鍵1(1) <> "" Then 検索鍵 = .Range(鍵1(1))
        If 鍵2(1) <> "" Then 検索鍵 = 検索鍵 & "-" & .Range(鍵2(1))
        If 鍵3(1) <> "" Then 検索鍵 = 検索鍵 & "-" & .Range(鍵3(1))
        Select Case 検索鍵
            Case "", "-", "--"
                転記行検索 = ""
                Exit Function
        End Select
    End With
    With Sheets("管理台帳")
        最右列 = .Cells(1, 1).End(xlToRight).Column
        For 列 = 1 To 最右列
            If 最下行 < .Cells(Rows.Count, 列).End(xlUp).Row Then
                最下行 = .Cells(Rows.Count, 列).End(xlUp).Row
            End If
        Next
        If 最下行 < 2 Then
            転記行検索 = 0
            Exit Function
        End If
        ReDim 行配列(2 To 最下行)
        For 行 = 2 To 最下行
            If 鍵1(1) <> "" Then 行配列(行) = .Cells(行, 鍵1(2))
            If 鍵2(1) <> "" Then 行配列(行) = 行配列(行) & "-" & .Cells(行, 鍵2(2))
            If 鍵3(1) <> "" Then 行配列(行) = 行配列(行) & "-" & .Cells(行, 鍵3(2))
        Next
        For 行 = 2 To 最下行
            If 検索鍵 = 行配列(行) Then
                転記行検索 = 行
                .Range("_選択行") = 行
                Exit For
            End If
        Next
    End With
End Function
Function 編集差分確認() As String
    Dim 終行 As Long, 行 As Long, 最右列 As Long, 列 As Long
    Dim 記録行 As Variant
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
        記録行 = 転記行検索()
        Select Case 記録行
            Case 0
                編集差分確認 = "未登録"
                Exit Function
            Case ""
                編集差分確認 = ""
                Exit Function
        End Select
        For 列 = 1 To 最右列
            If 配列(1, 列) <> .Cells(記録行, 列) Then
                編集差分確認 = "差分あり"
                Exit For
            End If
        Next
    End With
End Function
Sub 登録更新()
    Dim 終行 As Long, 行 As Long, 最右列 As Long, 列 As Long, 最下行 As Long, 記録行 As Long
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
            If 最下行 < .Cells(Rows.Count, 列).End(xlUp).Row Then
                最下行 = .Cells(Rows.Count, 列).End(xlUp).Row
            End If
        Next
        記録行 = 転記行検索()
        If 記録行 = 0 Then
            記録行 = 最下行 + 1
            文 = "新規登録してよろしいですか？"
        Else: 文 = "台帳を更新してよろしいですか？"
        End If
        If MsgBox(文, vbYesNo) = vbYes Then
            .Cells(記録行, 1).Resize(1, 最右列) = 配列
            MsgBox "台帳記録が完了しました"
        End If
        Call 台帳並替
    End With
End Sub
Sub 台帳戻し(記録行 As Variant)
    Dim 終行 As Long, 行 As Long, 最右列 As Long, 列 As Long
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
        If 記録行 = 0 Then 記録行 = 転記行検索()
        Select Case 記録行
            Case 0, ""
                Call 入力フォームクリア
                Exit Sub
        End Select
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
