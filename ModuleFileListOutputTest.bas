Attribute VB_Name = "Module1"
Sub ファイル一覧取得テスト()
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
    'ファイル出力
    'Call 配列テキスト改行出力(aryFile)
    Debug.Print "-------- End --------"
End Sub

Sub ファイル一覧取得テスト再帰あり()
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
    aryDir(i) = argDir '引数のフォルダを配列の先頭に入れる
    '実行結果が分かりやすいように、テスト的にセルに書き出す場合(1/3)
    Cells(UBound(aryDir) + 1, 2) = aryDir(UBound(aryDir))
    
    '指定フォルダ以下の全サブフォルダを取得し、配列aryDirに入れます。
    i = 0
    Do
      strName = Dir(aryDir(i) & "\", vbDirectory)
      Do While strName <> ""
        If GetAttr(aryDir(i) & "\" & strName) And vbDirectory Then
          If strName <> "." And strName <> ".." Then
            ReDim Preserve aryDir(UBound(aryDir) + 1)
            aryDir(UBound(aryDir)) = aryDir(i) & "\" & strName
            '実行結果が分かりやすいように、テスト的にセルに書き出す場合(2/3)
            Cells(UBound(aryDir) + 1, 2) = aryDir(UBound(aryDir))
          End If
        End If
        strName = Dir()
      Loop
      i = i + 1
      If i > UBound(aryDir) Then Exit Do
    Loop
    
    '配列aryDirの全フォルダについて、ファイルを取得し、配列aryFileに入れます。
    ReDim aryFile(0)
    For i = 0 To UBound(aryDir)
      strName = Dir(aryDir(i) & "\", vbNormal + vbHidden + vbReadOnly + vbSystem)
      Do While strName <> ""
        If aryFile(0) <> "" Then
          ReDim Preserve aryFile(UBound(aryFile) + 1)
        End If
        aryFile(UBound(aryFile)) = aryDir(i) & "\" & strName
        '実行結果が分かりやすいように、テスト的にセルに書き出す場合(3/3)
        Cells(UBound(aryFile) + 1, 2) = aryFile(UBound(aryFile))
        strName = Dir()
      Loop
    Next
    'ファイル出力
    'Call 配列テキスト改行出力(aryFile)
End Sub


Sub サブディレクトリ内ファイル一覧取得テスト()
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
    aryDir(i) = argDir '引数のフォルダを配列の先頭に入れる
    '実行結果が分かりやすいように、テスト的にセルに書き出す場合(1/3)
    Cells(UBound(aryDir) + 1, 2) = aryDir(UBound(aryDir))
    
    '指定フォルダ以下の全サブフォルダを取得し、配列aryDirに入れます。
    i = 0
    Do
      strName = Dir(aryDir(i) & "\", vbDirectory)
      Do While strName <> ""
        If GetAttr(aryDir(i) & "\" & strName) And vbDirectory Then
          If strName <> "." And strName <> ".." Then
            ReDim Preserve aryDir(UBound(aryDir) + 1)
            aryDir(UBound(aryDir)) = aryDir(i) & "\" & strName
            '実行結果が分かりやすいように、テスト的にセルに書き出す場合(2/3)
            Cells(UBound(aryDir) + 1, 2) = aryDir(UBound(aryDir))
          End If
        End If
        strName = Dir()
      Loop
      i = i + 1
      If i > UBound(aryDir) Then Exit Do
    Loop
    
    'ディレクトリパス保存アレイから指定フォルダパス(aryDir(0))を削除します｡
    Call Call_Array_Remove(aryDir, 0)
    
    '配列aryDirの全フォルダについて、ファイルを取得し、配列aryFileに入れます。
    ReDim aryFile(0)
    For i = 0 To UBound(aryDir)
      strName = Dir(aryDir(i) & "\", vbNormal + vbHidden + vbReadOnly + vbSystem)
      Do While strName <> ""
        If aryFile(0) <> "" Then
          ReDim Preserve aryFile(UBound(aryFile) + 1)
        End If
        aryFile(UBound(aryFile)) = aryDir(i) & "\" & strName
        '実行結果が分かりやすいように、テスト的にセルに書き出す場合(3/3)
        Cells(UBound(aryFile) + 1, 3) = aryFile(UBound(aryFile))
        strName = Dir()
      Loop
    Next
    
    'ファイル出力
    Call 配列テキスト改行出力(aryFile)
    
    '指定パスを削除し、サブディレクトリ名とファイル名のみ出力
    Call 配列内の指定文字を削除して出力(aryFile, wbpath & "\")
    
    'ファイル名のみ出力(拡張子あり)
    Call ファイル名のみ出力(aryFile, "output_filename_only", 0)
    
    'ファイル名のみ出力(拡張子無し)
    Call ファイル名のみ出力(aryFile, "output_filename_only_remove_extension", 1)
End Sub

'一次元配列内の〇番目要素を削除する
Public Sub Call_Array_Remove(ByRef ary As Variant, ByVal num As Long)
    Dim i As Long
    '削除したい〇番目の要素以降のを前につめて上書きコピーする
    For i = num To UBound(ary) - 1
        ary(i) = ary(i + 1)
    Next i
    '配列を再定義し、最終の要素を詰める
    ReDim Preserve ary(UBound(ary) - 1)
End Sub

Sub 一覧クリア()
    Range("A:AA").ClearContents
End Sub

Sub 配列テキスト改行出力(ByRef ary As Variant, Optional filename As String = "output")
    'アウトプット時のファイルナンバー変数
    Dim fnum As Integer
    fnum = FreeFile
    Dim i As Integer, j As Integer
    Open ThisWorkbook.Path & "\" & filename & ".txt" For Output As fnum
    For i = 0 To UBound(ary)
        Print #fnum, ary(i)
    Next
    Close fnum
End Sub

Sub 配列内の指定文字を削除して出力(ByRef ary As Variant, ByVal remove_str As String, Optional filename As String = "output_remove_str")
    Dim i As Long
    
    '部分削除
    For i = num To UBound(ary)
        ary(i) = Replace(ary(i), remove_str, "")
        '実行結果が分かりやすいように、テスト的にセルに書き出す場合
        Cells(i + 1, 4) = ary(i)
    Next i
    
    '書き出し
    Call 配列テキスト改行出力(ary, filename)
End Sub

Sub ファイル名のみ出力(ByRef ary As Variant, Optional filename As String = "output_filename_only", Optional remove_extension_1or0 As Integer = 0)
    Dim i As Long
    Dim filepath As String
    'Dim fpath As String
    Dim fname As String, pos As Long
    Dim extension_len As Integer
    
    '部分削除
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
        '実行結果が分かりやすいように、テスト的にセルに書き出す場合
        Cells(i + 1, 5) = ary(i)
    Next i
    
    '書き出し
    Call 配列テキスト改行出力(ary, filename)
End Sub


