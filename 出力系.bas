Attribute VB_Name = "�o�͌n"
Option Explicit
Sub �o�̓��[�h���s()
    Select Case Sheets("���̓t�H�[��").Range("_�o�̓��[�h��")
        Case "�ʈ��PV": Call �ʈ��PV
        Case "�A�����": Call �A�����
        Case "��PDF�o��": Call ��PDF�o��
        Case "�A��PDF�o��": Call �A��PDF�o��
    End Select
End Sub
Sub �ʈ��PV()
    Application.EnableEvents = False
    Sheets("����l��").PrintPreview
    Application.EnableEvents = True
End Sub
Sub �A�����()
    Dim �J�n�ԍ� As Long, �I���ԍ� As Long, �s As Long
    Dim �� As String
    �� = "�A������ݒ�" & vbCrLf & "�J�n�s�ԍ�����͂��Ă��������B"
    �J�n�ԍ� = Application.InputBox(��, Type:=1)
    Select Case �J�n�ԍ�
        Case 0: Exit Sub
        Case Is < 2
            MsgBox "�J�n�s�ԍ��ɂ�2�ȏ�̐��l����͂��Ă�������"
            Exit Sub
    End Select
    �� = "�I���s�ԍ�����͂��Ă��������B"
    �I���ԍ� = Application.InputBox(��, Type:=1)
    Select Case �I���ԍ�
        Case 0: Exit Sub
        Case Is < �J�n�ԍ�
            MsgBox "�I���s�ԍ��ɂ͊J�n�s�ԍ����傫�����l����͂��Ă�������"
            Exit Sub
    End Select
    
    Application.ScreenUpdating = False
    With Sheets("�Ǘ��䒠")
        �� = "�J�n�s�F" & �J�n�ԍ� & vbCrLf & "�I���s�F" & �I���ԍ� & vbCrLf & "��������F" & �I���ԍ� - �J�n�ԍ� + 1 & vbCrLf & vbCrLf & "�A����������s���܂����H"
        If MsgBox(��, vbYesNo) = vbYes Then
            Application.EnableEvents = False
            For �s = �J�n�ԍ� To �I���ԍ�
                .Range("_�I���s") = �s
                Sheets("����l��").PrintOut
            Next
            Application.EnableEvents = True
        End If
    End With
    Application.ScreenUpdating = True
End Sub
Sub PDF�o��(�ۑ��� As String)
    With Sheets("����l��")
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:=�ۑ���
    End With
End Sub
Sub ��PDF�o��()
    Dim �t�H���_�� As String, �t�@�C���� As String
    �t�@�C���� = Application.InputBox("PDF�t�@�C��������͂��Ă�������", Default:=Format(Now, "yyyymmddnnss"), Type:=2)
    Select Case �t�@�C����
        Case False, "": Exit Sub
    End Select
    �t�H���_�� = ThisWorkbook.Path & "\�o��PDF"
    If Dir(�t�H���_��, vbDirectory) = "" Then MkDir �t�H���_��
    Call PDF�o��(�t�H���_�� & "\" & �t�@�C���� & ".pdf")
    MsgBox "�t�@�C�����F" & �t�@�C���� & ".pdf" & vbCrLf & vbCrLf & "PDF�o�͂��������܂����i�{�c�[�����K�w�E�u�o��PDF�v�t�H���_���j"
End Sub
Sub �A��PDF�o��()
    Dim �e�t�H���_�� As String, �q�t�H���_�� As String
    Dim �J�n�ԍ� As Long, �I���ԍ� As Long, �s As Long
    Dim �� As String
    �� = "�A��PDF�o�͐ݒ�" & vbCrLf & "�J�n�s�ԍ�����͂��Ă��������B"
    �J�n�ԍ� = Application.InputBox(��, Type:=1)
    Select Case �J�n�ԍ�
        Case 0: Exit Sub
        Case Is < 2
            MsgBox "�J�n�s�ԍ��ɂ�2�ȏ�̐��l����͂��Ă�������"
            Exit Sub
    End Select
    �� = "�I���s�ԍ�����͂��Ă��������B"
    �I���ԍ� = Application.InputBox(��, Type:=1)
    Select Case �I���ԍ�
        Case 0: Exit Sub
        Case Is < �J�n�ԍ�
            MsgBox "�I���s�ԍ��ɂ͊J�n�s�ԍ����傫�����l����͂��Ă�������"
            Exit Sub
    End Select
    
    Application.ScreenUpdating = False
    With Sheets("�Ǘ��䒠")
        �� = "�J�n�s�F" & �J�n�ԍ� & vbCrLf & "�I���s�F" & �I���ԍ� & vbCrLf & "�o�̓t�@�C�����F" & �I���ԍ� - �J�n�ԍ� + 1
        �� = �� & vbCrLf & vbCrLf & "�A��PDF�o�͂����s���܂����H" & vbCrLf & "���{�c�[�����K�w�E�u�o��PDF�v�t�H���_���ɕۑ����܂�"
        If MsgBox(��, vbYesNo) = vbYes Then
            �e�t�H���_�� = ThisWorkbook.Path & "\" & "�o��PDF"
            If Dir(�e�t�H���_��, vbDirectory) = "" Then MkDir �e�t�H���_��
            �q�t�H���_�� = �e�t�H���_�� & "\" & Format(Now, "yyyymmddnnss")
            If Dir(�q�t�H���_��, vbDirectory) = "" Then MkDir �q�t�H���_��
            For �s = �J�n�ԍ� To �I���ԍ�
                .Range("_�I���s") = �s
                Call PDF�o��(�q�t�H���_�� & "\" & Format(�s, String(Len(Str(�I���ԍ�)), "0")) & ".pdf")
            Next
        End If
    End With
    Application.ScreenUpdating = True
    MsgBox "PDF�o�͂��������܂���" & vbCrLf & vbCrLf & "�ۑ���F�{�c�[�����K�w�E�u�o��PDF�v�t�H���_��"
End Sub
