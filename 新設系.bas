Attribute VB_Name = "�V�݌n"
Option Explicit
Sub �c�[���V�K�쐬()
    Dim �I�s As Long, �s As Long, �ŉE�� As Long, �� As Long
    Dim ��`�� As String
    If MsgBox("���ڐݒ�����ɐV�����c�[���x�[�X���쐬���܂�" & vbCrLf & vbCrLf & "��낵���ł����H", vbYesNo) <> vbYes Then Exit Sub
    Call �S�V�[�g�W�J
    Call ��`���S�폜
    With Sheets("�䒠�]�L�ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim ���ڃ��X�g(1 To �I�s - 1, 1 To 1)
        For �s = 2 To �I�s
            ���ڃ��X�g(�s - 1, 1) = .Cells(�s, 1)
            ��`�� = ���ڃ��X�g(�s - 1, 1)
            Call ���O�̒�`�ǉ�(��`��, "=���̓t�H�[��!R" & �s & "C2")
            Call �䒠��֐��ݒ�(.Cells(�s, 2), �s - 1)
        Next
        .Buttons.Delete
    End With
    With Sheets("���̓t�H�[��")
        .Range("A2").Resize(Rows.Count - 1, 2).ClearContents
        .Range("A2").Resize(Rows.Count - 1, 2).Borders.LineStyle = False
        .Range("A2").Resize(Rows.Count - 1, 2).Interior.ColorIndex = 0
        .Range("A2").Resize(UBound(���ڃ��X�g, 1), 1) = ���ڃ��X�g
        .Range("B2").Resize(UBound(���ڃ��X�g, 1), 1).Locked = False
        .Range("B2").Resize(UBound(���ڃ��X�g, 1), 1).Interior.ColorIndex = 6
        .Range("A2").Resize(UBound(���ڃ��X�g, 1), 2).Borders.LineStyle = True
        .Activate
    End With
    With Sheets("�Ǘ��䒠")
        �ŉE�� = �I�s - 1
        ReDim �䒠����(1 To 1, 1 To �ŉE��)
        For �� = 1 To �ŉE��
            �䒠����(1, ��) = ���ڃ��X�g(��, 1)
        Next
        .Range("A1").Resize(Rows.Count, Columns.Count).ClearContents
        .Range("A1").Resize(Rows.Count, Columns.Count).Borders.LineStyle = False
        .Range("A1").Resize(Rows.Count, Columns.Count).Interior.ColorIndex = 0
        .Range("A1").Resize(1, �ŉE��) = �䒠����
        .Range("A1").Resize(1, �ŉE��).Borders.LineStyle = True
        .Range("A1").Resize(1, �ŉE��).Interior.Color = RGB(142, 169, 219)
        Range(.Columns(1), .Columns(�ŉE��)).AutoFit
        .Cells(1, �ŉE�� + 2) = "�I���s"
        .Cells(1, �ŉE�� + 2).Resize(1, 2).Borders.LineStyle = True
        .Cells(1, �ŉE�� + 2).Interior.ColorIndex = 6
        Call ���O�̒�`�ǉ�("_�I���s", "=�Ǘ��䒠!R1C" & �ŉE�� + 3)
        Call �䒠�����ݒ�(101, �ŉE��)
    End With
    With Sheets("����l��")
        .Range("A1").Resize(Rows.Count, 2).ClearContents
        .Range("A1").Resize(Rows.Count, 2).Interior.ColorIndex = 0
        .Range("A1") = "�����F�Z��������l���p�[�c"
        .Range("A2").Resize(UBound(���ڃ��X�g, 1), 1) = ���ڃ��X�g
        For �s = 2 To �I�s
            Call ����l���Q�Ɗ֐��ݒ�(.Cells(�s, 2), Sheets("�䒠�]�L�ݒ�").Cells(�s, 2))
        Next
        .Range("B2").Resize(UBound(���ڃ��X�g, 1), 1).Interior.ColorIndex = 6
    End With
End Sub
Sub ��`���S�폜()
    Dim ��`�� As Name
    On Error Resume Next
    For Each ��`�� In ThisWorkbook.Names
        If Left(��`��.Name, 1) <> "_" Then ��`��.Delete
    Next
    On Error GoTo 0
End Sub
Sub ���O�̒�`�ǉ�(��`�� As String, �Z���͈� As String)
    On Error GoTo �G���[������
    ThisWorkbook.Names.Add Name:=��`��, RefersToR1C1:=�Z���͈�
    Exit Sub
�G���[������:
    Dim �G���[�� As String
    Select Case Err.Number
        Case 1004
            �G���[�� = "�s���ȍ��ږ��i��`���j���܂܂�Ă��܂��i�Ώۍ��ږ��F" & ��`�� & "�j"
            �G���[�� = �G���[�� & vbCrLf & vbCrLf & "���ږ��i��`���j���ȉ��̏����𖞂����Ă��邩���m�F��������"
            �G���[�� = �G���[�� & vbCrLf & "�@�E���O�̐擪�͉p�����A���ȁA�J�i�A�����A�A���_�[�X�R�A�i_�j�̂����ꂩ�ł���"
            �G���[�� = �G���[�� & vbCrLf & "�@�E�󔒓��̖����ȕ������܂܂�Ă��Ȃ�"
            �G���[�� = �G���[�� & vbCrLf & "�@�E�����̖��O�Ƌ������Ă��Ȃ�"
            MsgBox �G���[��
            Call ��`���S�폜
            End
    End Select
End Sub
Sub �䒠��֐��ݒ�(�Z�� As Range, �� As Long)
    �Z��.FormulaR1C1 = "=COLUMN(�Ǘ��䒠!C" & �� & ")"
End Sub
Sub ����l���Q�Ɗ֐��ݒ�(�Z�� As Range, ��ԍ� As Long)
    �Z��.FormulaR1C1 = "=INDIRECT(ADDRESS(_�I���s,COLUMN(�Ǘ��䒠!C" & ��ԍ� & "),,,""�Ǘ��䒠""))"
End Sub
