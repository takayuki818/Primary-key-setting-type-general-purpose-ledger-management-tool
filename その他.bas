Attribute VB_Name = "���̑�"
Option Explicit
Sub �ی�ؑ�()
    With ActiveSheet
        Select Case .ProtectContents
            Case True: .Unprotect: MsgBox "�V�[�g�ی���������܂���"
            Case False: .Protect: MsgBox "�V�[�g��ی삵�܂���"
        End Select
    End With
End Sub
Sub �S�V�[�g�W�J()
    Dim �V�[�g As Worksheet
    Application.ScreenUpdating = False
    For Each �V�[�g In Sheets
        �V�[�g.Visible = True
    Next
    Application.ScreenUpdating = True
End Sub
Sub �����Đݒ�()
    Dim �ŉ��s As Long, �ŉE�� As Long, �� As Long
    With Sheets("�Ǘ��䒠")
        �ŉE�� = .Cells(1, 1).End(xlToRight).Column
        �ŉ��s = 2
        For �� = 1 To �ŉE��
            If �ŉ��s < .Cells(Rows.Count, ��).End(xlUp).Row Then
                �ŉ��s = .Cells(Rows.Count, ��).End(xlUp).Row
            End If
        Next
        Call �䒠�����ݒ�(�ŉ��s + 100, �ŉE��)
        Application.EnableEvents = False
        .Activate
        Application.EnableEvents = True
        MsgBox "�u�䒠�v�V�[�g�̏������Đݒ肵�܂���" & vbCrLf & "���ŉ��s+100�s�܂ŏ����ݒ�s����"
    End With
End Sub
Sub �䒠�����ݒ�(�ŉ��s As Long, �ŉE�� As Long)
    Dim ���� As FormatCondition
    With Sheets("�Ǘ��䒠")
        .Range("A2").Resize(Rows.Count - 1, Columns.Count).Borders.LineStyle = False
        .Range("A2").Resize(�ŉ��s - 1, �ŉE��).Borders.LineStyle = True
        .Cells.FormatConditions.Delete
        Set ���� = .Range("A2").Resize(�ŉ��s - 1, �ŉE��).FormatConditions.Add(Type:=xlExpression, Formula1:="=ROW()=_�I���s")
        ����.Interior.Color = RGB(217, 225, 242)
    End With
End Sub
