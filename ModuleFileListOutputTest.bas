Attribute VB_Name = "Module1"
Sub �t�@�C���ꗗ�擾�e�X�g()
    Dim wbpath, wbname As String, pos As Long
    pos = InStrRev(ThisWorkbook.Path, "\")
    'wbpath = Left(ThisWorkbook.Path, pos)
    'wbname = Mid(ThisWorkbook.Path, pos + 1)
    wbpath = ThisWorkbook.Path
    wbname = ThisWorkbook.Name
    Debug.Print vbCrLf + vbCrLf
    Debug.Print "--------Start--------"
    Debug.Print "WorkBook File: "; wbpath + "\" + wbname
    Debug.Print "WorkBook Path: " + wbpath
    Debug.Print "WorkBook Name: " + wbname
    
    Dim buf As String, cnt As Long
    Dim Path As String
    Path = wbpath
    Cells(1, 1) = Path
    Cells(2, 1) = Now()
    buf = Dir(Path & "\" & "*.*")
    Do While buf <> ""
        cnt = cnt + 1
        Cells(cnt, 2) = buf
        buf = Dir()
    Loop
    '�t�@�C���o��
    'Call �z��e�L�X�g���s�o��(aryFile)
    Debug.Print "-------- End --------"
End Sub

Sub �t�@�C���ꗗ�擾�e�X�g�ċA����()
    Dim wbpath, wbname As String, pos As Long
    pos = InStrRev(ThisWorkbook.Path, "\")
    wbpath = ThisWorkbook.Path
    wbname = ThisWorkbook.Name
    Debug.Print vbCrLf + vbCrLf
    Debug.Print "--------Start--------"
    Debug.Print "WorkBook File: "; wbpath + "\" + wbname
    Debug.Print "WorkBook Path: " + wbpath
    Debug.Print "WorkBook Name: " + wbname
    
    
    Dim argDir As String
    argDir = wbpath
    
    Cells(1, 1) = argDir
    Cells(2, 1) = Now()
    
    Dim i As Long
    Dim aryDir() As String
    Dim aryFile() As String
    Dim strName As String
    
    ReDim aryDir(i)
    aryDir(i) = argDir '�����̃t�H���_��z��̐擪�ɓ����
    '���s���ʂ�������₷���悤�ɁA�e�X�g�I�ɃZ���ɏ����o���ꍇ(1/3)
    Cells(UBound(aryDir) + 1, 2) = aryDir(UBound(aryDir))
    
    '�w��t�H���_�ȉ��̑S�T�u�t�H���_���擾���A�z��aryDir�ɓ���܂��B
    i = 0
    Do
      strName = Dir(aryDir(i) & "\", vbDirectory)
      Do While strName <> ""
        If GetAttr(aryDir(i) & "\" & strName) And vbDirectory Then
          If strName <> "." And strName <> ".." Then
            ReDim Preserve aryDir(UBound(aryDir) + 1)
            aryDir(UBound(aryDir)) = aryDir(i) & "\" & strName
            '���s���ʂ�������₷���悤�ɁA�e�X�g�I�ɃZ���ɏ����o���ꍇ(2/3)
            Cells(UBound(aryDir) + 1, 2) = aryDir(UBound(aryDir))
          End If
        End If
        strName = Dir()
      Loop
      i = i + 1
      If i > UBound(aryDir) Then Exit Do
    Loop
    
    '�z��aryDir�̑S�t�H���_�ɂ��āA�t�@�C�����擾���A�z��aryFile�ɓ���܂��B
    ReDim aryFile(0)
    For i = 0 To UBound(aryDir)
      strName = Dir(aryDir(i) & "\", vbNormal + vbHidden + vbReadOnly + vbSystem)
      Do While strName <> ""
        If aryFile(0) <> "" Then
          ReDim Preserve aryFile(UBound(aryFile) + 1)
        End If
        aryFile(UBound(aryFile)) = aryDir(i) & "\" & strName
        '���s���ʂ�������₷���悤�ɁA�e�X�g�I�ɃZ���ɏ����o���ꍇ(3/3)
        Cells(UBound(aryFile) + 1, 2) = aryFile(UBound(aryFile))
        strName = Dir()
      Loop
    Next
    '�t�@�C���o��
    'Call �z��e�L�X�g���s�o��(aryFile)
End Sub


Sub �T�u�f�B���N�g�����t�@�C���ꗗ�擾�e�X�g()
    Dim wbpath, wbname As String, pos As Long
    pos = InStrRev(ThisWorkbook.Path, "\")
    'wbpath = Left(ThisWorkbook.Path, pos)
    'wbname = Mid(ThisWorkbook.Path, pos + 1)
    wbpath = ThisWorkbook.Path
    wbname = ThisWorkbook.Name
    Debug.Print vbCrLf + vbCrLf
    Debug.Print "--------Start--------"
    Debug.Print "WorkBook File: "; wbpath + "\" + wbname
    Debug.Print "WorkBook Path: " + wbpath
    Debug.Print "WorkBook Name: " + wbname
    
    
    Dim argDir As String
    argDir = wbpath
    
    Cells(1, 1) = argDir
    Cells(2, 1) = Now()
    
    Dim i As Long
    Dim aryDir() As String
    Dim aryFile() As String
    Dim strName As String
    Dim arySubDir() As String
    
    ReDim aryDir(i)
    aryDir(i) = argDir '�����̃t�H���_��z��̐擪�ɓ����
    '���s���ʂ�������₷���悤�ɁA�e�X�g�I�ɃZ���ɏ����o���ꍇ(1/3)
    Cells(UBound(aryDir) + 1, 2) = aryDir(UBound(aryDir))
    
    '�w��t�H���_�ȉ��̑S�T�u�t�H���_���擾���A�z��aryDir�ɓ���܂��B
    i = 0
    Do
      strName = Dir(aryDir(i) & "\", vbDirectory)
      Do While strName <> ""
        If GetAttr(aryDir(i) & "\" & strName) And vbDirectory Then
          If strName <> "." And strName <> ".." Then
            ReDim Preserve aryDir(UBound(aryDir) + 1)
            aryDir(UBound(aryDir)) = aryDir(i) & "\" & strName
            '���s���ʂ�������₷���悤�ɁA�e�X�g�I�ɃZ���ɏ����o���ꍇ(2/3)
            Cells(UBound(aryDir) + 1, 2) = aryDir(UBound(aryDir))
          End If
        End If
        strName = Dir()
      Loop
      i = i + 1
      If i > UBound(aryDir) Then Exit Do
    Loop
    
    '�f�B���N�g���p�X�ۑ��A���C����w��t�H���_�p�X(aryDir(0))���폜���܂��
    Call Call_Array_Remove(aryDir, 0)
    
    '�z��aryDir�̑S�t�H���_�ɂ��āA�t�@�C�����擾���A�z��aryFile�ɓ���܂��B
    ReDim aryFile(0)
    For i = 0 To UBound(aryDir)
      strName = Dir(aryDir(i) & "\", vbNormal + vbHidden + vbReadOnly + vbSystem)
      Do While strName <> ""
        If aryFile(0) <> "" Then
          ReDim Preserve aryFile(UBound(aryFile) + 1)
        End If
        aryFile(UBound(aryFile)) = aryDir(i) & "\" & strName
        '���s���ʂ�������₷���悤�ɁA�e�X�g�I�ɃZ���ɏ����o���ꍇ(3/3)
        Cells(UBound(aryFile) + 1, 3) = aryFile(UBound(aryFile))
        strName = Dir()
      Loop
    Next
    
    '�t�@�C���o��
    Call �z��e�L�X�g���s�o��(aryFile)
    
    '�w��p�X���폜���A�T�u�f�B���N�g�����ƃt�@�C�����̂ݏo��
    Call �z����̎w�蕶�����폜���ďo��(aryFile, wbpath & "\")
    
    '�t�@�C�����̂ݏo��(�g���q����)
    Call �t�@�C�����̂ݏo��(aryFile, "output_filename_only", 0)
    
    '�t�@�C�����̂ݏo��(�g���q����)
    Call �t�@�C�����̂ݏo��(aryFile, "output_filename_only_remove_extension", 1)
End Sub

'�ꎟ���z����́Z�Ԗڗv�f���폜����
Public Sub Call_Array_Remove(ByRef ary As Variant, ByVal num As Long)
    Dim i As Long
    '�폜�������Z�Ԗڂ̗v�f�ȍ~�̂�O�ɂ߂ď㏑���R�s�[����
    For i = num To UBound(ary) - 1
        ary(i) = ary(i + 1)
    Next i
    '�z����Ē�`���A�ŏI�̗v�f���l�߂�
    ReDim Preserve ary(UBound(ary) - 1)
End Sub

Sub �ꗗ�N���A()
    Range("A:AA").ClearContents
End Sub

Sub �z��e�L�X�g���s�o��(ByRef ary As Variant, Optional filename As String = "output")
    '�A�E�g�v�b�g���̃t�@�C���i���o�[�ϐ�
    Dim fnum As Integer
    fnum = FreeFile
    Dim i As Integer, j As Integer
    Open ThisWorkbook.Path & "\" & filename & ".txt" For Output As fnum
    For i = 0 To UBound(ary)
        Print #fnum, ary(i)
    Next
    Close fnum
End Sub

Sub �z����̎w�蕶�����폜���ďo��(ByRef ary As Variant, ByVal remove_str As String, Optional filename As String = "output_remove_str")
    Dim i As Long
    
    '�����폜
    For i = num To UBound(ary)
        ary(i) = Replace(ary(i), remove_str, "")
        '���s���ʂ�������₷���悤�ɁA�e�X�g�I�ɃZ���ɏ����o���ꍇ
        Cells(i + 1, 4) = ary(i)
    Next i
    
    '�����o��
    Call �z��e�L�X�g���s�o��(ary, filename)
End Sub

Sub �t�@�C�����̂ݏo��(ByRef ary As Variant, Optional filename As String = "output_filename_only", Optional remove_extension_1or0 As Integer = 0)
    Dim i As Long
    Dim filepath As String
    'Dim fpath As String
    Dim fname As String, pos As Long
    Dim extension_len As Integer
    
    '�����폜
    For i = num To UBound(ary)
        filepath = ary(i)
        pos = InStrRev(filepath, "\")
        'fpath = Left(filepath, pos)
        fname = Mid(filepath, pos + 1)
        If remove_extension_1or0 <> 0 Then
            pos = 0
            pos = InStrRev(fname, ".")
            fname = Left(fname, pos - 1)
        End If
        ary(i) = fname
        '���s���ʂ�������₷���悤�ɁA�e�X�g�I�ɃZ���ɏ����o���ꍇ
        Cells(i + 1, 5) = ary(i)
    Next i
    
    '�����o��
    Call �z��e�L�X�g���s�o��(ary, filename)
End Sub


