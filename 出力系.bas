Attribute VB_Name = "出力系"
Option Explicit
Sub 出力モード実行()
    Select Case Sheets("入力フォーム").Range("_出力モード名")
        Case "個別印刷PV": Call 個別印刷PV
        Case "連続印刷": Call 連続印刷
        Case "個別PDF出力": Call 個別PDF出力
        Case "連続PDF出力": Call 連続PDF出力
    End Select
End Sub
Sub 個別印刷PV()
    Application.EnableEvents = False
    Sheets("印刷様式").PrintPreview
    Application.EnableEvents = True
End Sub
Sub 連続印刷()
    Dim 開始番号 As Long, 終了番号 As Long, 行 As Long
    Dim 文 As String
    文 = "連続印刷設定" & vbCrLf & "開始行番号を入力してください。"
    開始番号 = Application.InputBox(文, Type:=1)
    Select Case 開始番号
        Case 0: Exit Sub
        Case Is < 2
            MsgBox "開始行番号には2以上の数値を入力してください"
            Exit Sub
    End Select
    文 = "終了行番号を入力してください。"
    終了番号 = Application.InputBox(文, Type:=1)
    Select Case 終了番号
        Case 0: Exit Sub
        Case Is < 開始番号
            MsgBox "終了行番号には開始行番号より大きい数値を入力してください"
            Exit Sub
    End Select
    
    Application.ScreenUpdating = False
    With Sheets("管理台帳")
        文 = "開始行：" & 開始番号 & vbCrLf & "終了行：" & 終了番号 & vbCrLf & "印刷枚数：" & 終了番号 - 開始番号 + 1 & vbCrLf & vbCrLf & "連続印刷を実行しますか？"
        If MsgBox(文, vbYesNo) = vbYes Then
            Application.EnableEvents = False
            For 行 = 開始番号 To 終了番号
                .Range("_選択行") = 行
                Sheets("印刷様式").PrintOut
            Next
            Application.EnableEvents = True
        End If
    End With
    Application.ScreenUpdating = True
End Sub
Sub PDF出力(保存先 As String)
    With Sheets("印刷様式")
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:=保存先
    End With
End Sub
Sub 個別PDF出力()
    Dim フォルダ名 As String, ファイル名 As String
    ファイル名 = Application.InputBox("PDFファイル名を入力してください", Default:=Format(Now, "yyyymmddnnss"), Type:=2)
    Select Case ファイル名
        Case False, "": Exit Sub
    End Select
    フォルダ名 = ThisWorkbook.Path & "\出力PDF"
    If Dir(フォルダ名, vbDirectory) = "" Then MkDir フォルダ名
    Call PDF出力(フォルダ名 & "\" & ファイル名 & ".pdf")
    MsgBox "ファイル名：" & ファイル名 & ".pdf" & vbCrLf & vbCrLf & "PDF出力が完了しました（本ツール同階層・「出力PDF」フォルダ内）"
End Sub
Sub 連続PDF出力()
    Dim 親フォルダ名 As String, 子フォルダ名 As String
    Dim 開始番号 As Long, 終了番号 As Long, 行 As Long
    Dim 文 As String
    文 = "連続PDF出力設定" & vbCrLf & "開始行番号を入力してください。"
    開始番号 = Application.InputBox(文, Type:=1)
    Select Case 開始番号
        Case 0: Exit Sub
        Case Is < 2
            MsgBox "開始行番号には2以上の数値を入力してください"
            Exit Sub
    End Select
    文 = "終了行番号を入力してください。"
    終了番号 = Application.InputBox(文, Type:=1)
    Select Case 終了番号
        Case 0: Exit Sub
        Case Is < 開始番号
            MsgBox "終了行番号には開始行番号より大きい数値を入力してください"
            Exit Sub
    End Select
    
    Application.ScreenUpdating = False
    With Sheets("管理台帳")
        文 = "開始行：" & 開始番号 & vbCrLf & "終了行：" & 終了番号 & vbCrLf & "出力ファイル数：" & 終了番号 - 開始番号 + 1
        文 = 文 & vbCrLf & vbCrLf & "連続PDF出力を実行しますか？" & vbCrLf & "※本ツール同階層・「出力PDF」フォルダ内に保存します"
        If MsgBox(文, vbYesNo) = vbYes Then
            親フォルダ名 = ThisWorkbook.Path & "\" & "出力PDF"
            If Dir(親フォルダ名, vbDirectory) = "" Then MkDir 親フォルダ名
            子フォルダ名 = 親フォルダ名 & "\" & Format(Now, "yyyymmddnnss")
            If Dir(子フォルダ名, vbDirectory) = "" Then MkDir 子フォルダ名
            For 行 = 開始番号 To 終了番号
                .Range("_選択行") = 行
                Call PDF出力(子フォルダ名 & "\" & Format(行, String(Len(Str(終了番号)), "0")) & ".pdf")
            Next
        End If
    End With
    Application.ScreenUpdating = True
    MsgBox "PDF出力が完了しました" & vbCrLf & vbCrLf & "保存先：本ツール同階層・「出力PDF」フォルダ内"
End Sub
