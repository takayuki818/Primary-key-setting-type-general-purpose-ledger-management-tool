Attribute VB_Name = "���͌n"
Option Explicit
Sub ��L�[�擾(��1 As Variant, ��2 As Variant, ��3 As Variant)
    Dim �I�s As Long, �s As Long
    With Sheets("�䒠�]�L�ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        For �s = 2 To �I�s
            If .Cells(�s, 3) <> "" Then
                If ��1(1) = "" Then
                    ��1(1) = .Cells(�s, 1)
                    ��1(2) = .Cells(�s, 2)
                Else
                    If ��2(1) = "" Then
                        ��2(1) = .Cells(�s, 1)
                        ��2(2) = .Cells(�s, 2)
                    Else
                        If ��3(1) = "" Then
                            ��3(1) = .Cells(�s, 1)
                            ��3(2) = .Cells(�s, 2)
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub
Sub �䒠����()
    Dim �ŉ��s As Long, �ŉE�� As Long, �� As Long
    Dim ��1(1 To 2), ��2(1 To 2), ��3(1 To 2)
    With Sheets("�Ǘ��䒠")
        �ŉE�� = .Cells(1, Columns.Count).End(xlToLeft).Column
        For �� = 1 To �ŉE��
            If �ŉ��s < .Cells(Rows.Count, ��).End(xlUp).Row Then
                �ŉ��s = .Cells(Rows.Count, ��).End(xlUp).Row
            End If
        Next
        .Cells(1, 1).Resize(�ŉ��s, �ŉE��).Characters.PhoneticCharacters = ""
        Call ��L�[�擾(��1, ��2, ��3)
        With .Sort
            With .SortFields
                .Clear
                If ��1(1) <> "" Then
                    .Add Key:=Cells(1, ��1(2)), Order:=xlAscending
                End If
                If ��2(1) <> "" Then
                    .Add Key:=Cells(1, ��2(2)), Order:=xlAscending
                End If
                If ��3(1) <> "" Then
                    .Add Key:=Cells(1, ��3(2)), Order:=xlAscending
                End If
            End With
            .SetRange Range(Cells(1, 1), Cells(�ŉ��s, �ŉE��))
            .Header = xlYes
            .Apply
        End With
    End With
End Sub
Sub ���̓t�H�[���N���A()
    Dim ��1(1 To 2), ��2(1 To 2), ��3(1 To 2)
    Dim �I�s As Long, �s As Long
    Call ��L�[�擾(��1, ��2, ��3)
    With Sheets("�䒠�]�L�ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim ���ڃ��X�g(2 To �I�s)
        For �s = 2 To �I�s
            ���ڃ��X�g(�s) = .Cells(�s, 1)
        Next
    End With
    With Sheets("���̓t�H�[��")
        .Unprotect
        For �s = 2 To �I�s
            Select Case ���ڃ��X�g(�s)
                Case ��1(1), ��2(1), ��3(1)
                Case Else: .Range(���ڃ��X�g(�s)).MergeArea.ClearContents
            End Select
        Next
        .Protect
    End With
End Sub
Sub ���̓t�H�[���I�[���N���A()
    Dim �I�s As Long, �s As Long
    With Sheets("�䒠�]�L�ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim ���ڃ��X�g(2 To �I�s)
        For �s = 2 To �I�s
            ���ڃ��X�g(�s) = .Cells(�s, 1)
        Next
    End With
    With Sheets("���̓t�H�[��")
        .Unprotect
        Application.EnableEvents = False
        For �s = 2 To �I�s
            Select Case ���ڃ��X�g(�s)
                Case Else: .Range(���ڃ��X�g(�s)).MergeArea.ClearContents
            End Select
        Next
        Application.EnableEvents = True
        .Protect
    End With
End Sub
Function �]�L�s����() As Variant '��L�[�̒l�������́��u�󕶎���v��Ԃ�/�䒠���o�^���u0�v��Ԃ�
    Dim �ŉE�� As Long, �� As Long, �ŉ��s As Long, �s As Long
    Dim ��1(1 To 2), ��2(1 To 2), ��3(1 To 2)
    Dim ������ As String
    With Sheets("���̓t�H�[��")
        Call ��L�[�擾(��1, ��2, ��3)
        If ��1(1) <> "" Then ������ = .Range(��1(1))
        If ��2(1) <> "" Then ������ = ������ & "-" & .Range(��2(1))
        If ��3(1) <> "" Then ������ = ������ & "-" & .Range(��3(1))
        Select Case ������
            Case "", "-", "--"
                �]�L�s���� = ""
                Exit Function
        End Select
    End With
    With Sheets("�Ǘ��䒠")
        �ŉE�� = .Cells(1, 1).End(xlToRight).Column
        For �� = 1 To �ŉE��
            If �ŉ��s < .Cells(Rows.Count, ��).End(xlUp).Row Then
                �ŉ��s = .Cells(Rows.Count, ��).End(xlUp).Row
            End If
        Next
        If �ŉ��s < 2 Then
            �]�L�s���� = 0
            Exit Function
        End If
        ReDim �s�z��(2 To �ŉ��s)
        For �s = 2 To �ŉ��s
            If ��1(1) <> "" Then �s�z��(�s) = .Cells(�s, ��1(2))
            If ��2(1) <> "" Then �s�z��(�s) = �s�z��(�s) & "-" & .Cells(�s, ��2(2))
            If ��3(1) <> "" Then �s�z��(�s) = �s�z��(�s) & "-" & .Cells(�s, ��3(2))
        Next
        For �s = 2 To �ŉ��s
            If ������ = �s�z��(�s) Then
                �]�L�s���� = �s
                .Range("_�I���s") = �s
                Exit For
            End If
        Next
    End With
End Function
Function �ҏW�����m�F() As String
    Dim �I�s As Long, �s As Long, �ŉE�� As Long, �� As Long
    Dim �L�^�s As Variant
    With Sheets("�䒠�]�L�ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim �ݒ�(1 To �I�s - 1, 1 To 2)
        For �s = 2 To �I�s
            �ݒ�(�s - 1, 1) = .Cells(�s, 1)
            �ݒ�(�s - 1, 2) = .Cells(�s, 2)
            If �ŉE�� < �ݒ�(�s - 1, 2) Then �ŉE�� = �ݒ�(�s - 1, 2)
        Next
    End With
    With Sheets("���̓t�H�[��")
        ReDim �z��(1 To 1, 1 To �ŉE��)
        For �s = 1 To UBound(�ݒ�, 1)
            �z��(1, �ݒ�(�s, 2)) = .Range(�ݒ�(�s, 1))
        Next
    End With
    With Sheets("�Ǘ��䒠")
        �L�^�s = �]�L�s����()
        Select Case �L�^�s
            Case 0
                �ҏW�����m�F = "���o�^"
                Exit Function
            Case ""
                �ҏW�����m�F = ""
                Exit Function
        End Select
        For �� = 1 To �ŉE��
            If �z��(1, ��) <> .Cells(�L�^�s, ��) Then
                �ҏW�����m�F = "��������"
                Exit For
            End If
        Next
    End With
End Function
Sub �o�^�X�V()
    Dim �I�s As Long, �s As Long, �ŉE�� As Long, �� As Long, �ŉ��s As Long, �L�^�s As Long
    Dim �� As String
    With Sheets("�䒠�]�L�ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim �ݒ�(1 To �I�s - 1, 1 To 2)
        For �s = 2 To �I�s
            �ݒ�(�s - 1, 1) = .Cells(�s, 1)
            �ݒ�(�s - 1, 2) = .Cells(�s, 2)
            If �ŉE�� < �ݒ�(�s - 1, 2) Then �ŉE�� = �ݒ�(�s - 1, 2)
        Next
    End With
    With Sheets("���̓t�H�[��")
        ReDim �z��(1 To 1, 1 To �ŉE��)
        For �s = 1 To UBound(�ݒ�, 1)
            �z��(1, �ݒ�(�s, 2)) = .Range(�ݒ�(�s, 1))
        Next
    End With
    With Sheets("�Ǘ��䒠")
        For �� = 1 To �ŉE��
            If �ŉ��s < .Cells(Rows.Count, ��).End(xlUp).Row Then
                �ŉ��s = .Cells(Rows.Count, ��).End(xlUp).Row
            End If
        Next
        �L�^�s = �]�L�s����()
        If �L�^�s = 0 Then
            �L�^�s = �ŉ��s + 1
            �� = "�V�K�o�^���Ă�낵���ł����H"
        Else: �� = "�䒠���X�V���Ă�낵���ł����H"
        End If
        If MsgBox(��, vbYesNo) = vbYes Then
            .Cells(�L�^�s, 1).Resize(1, �ŉE��) = �z��
            MsgBox "�䒠�L�^���������܂���"
        End If
        Call �䒠����
    End With
End Sub
Sub �䒠�߂�(�L�^�s As Variant)
    Dim �I�s As Long, �s As Long, �ŉE�� As Long, �� As Long
    Application.ScreenUpdating = False
    With Sheets("�䒠�]�L�ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim �ݒ�(1 To �I�s - 1, 1 To 2)
        For �s = 2 To �I�s
            �ݒ�(�s - 1, 1) = .Cells(�s, 1)
            �ݒ�(�s - 1, 2) = .Cells(�s, 2)
            If �ŉE�� < �ݒ�(�s - 1, 2) Then �ŉE�� = �ݒ�(�s - 1, 2)
        Next
    End With
    With Sheets("�Ǘ��䒠")
        If �L�^�s = 0 Then �L�^�s = �]�L�s����()
        Select Case �L�^�s
            Case 0, ""
                Call ���̓t�H�[���N���A
                Exit Sub
        End Select
        ReDim �z��(1 To 1, 1 To �ŉE��)
        For �� = 1 To �ŉE��
            �z��(1, ��) = .Cells(�L�^�s, ��)
        Next
    End With
    With Sheets("���̓t�H�[��")
        Call ���̓t�H�[���N���A
        .Unprotect
        Application.EnableEvents = False
        For �s = 1 To UBound(�ݒ�, 1)
            .Range(�ݒ�(�s, 1)) = �z��(1, �ݒ�(�s, 2))
        Next
        Application.EnableEvents = True
        .Protect
    End With
    Application.ScreenUpdating = True
End Sub
