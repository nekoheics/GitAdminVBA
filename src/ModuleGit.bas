Attribute VB_Name = "ModuleGit"
Option Explicit

' ルートの親フォルダ
Public Const parentDir As String = "Source¥Repos¥VBA"
' Git設定ファイルの内容が埋め込まれているモジュール名
Private Const ContentsModuleName As String = "ModuleGitFilesContents"

' GitCmd引数用
Public Enum GitCommand
    Init
    Status
    Stage
    Commit
    Push
End Enum

' エラー出力
Public Sub OutputError(errPlace As String, Optional errNote As String)
    Dim msg As String
    msg = vbCrLf
    msg = "日時：" & Format$(Now(), "yyyy/mm/dd hh:nn:ss") & vbCrLf
    msg = msg & "ソース：" & Err.Source & vbCrLf
    msg = msg & "ブック名：" & ActiveWorkbook.Name & vbCrLf
    msg = msg & "場所：" & errPlace & vbCrLf
    msg = msg & "備考：" & errNote & vbCrLf
    msg = msg & "エラー番号：" & Err.Number & vbCrLf
    msg = msg & "エラー内容：" & Err.Description & vbCrLf
    Debug.Print msg
End Sub

Public Sub CreateNewRepository()
    
    ' リポジトリ名を設定する
    Dim repoName As String: repoName = SetAndThenGetReposName
    If repoName = "" Then Exit Sub
    
    ' ローカルリポジトリフォルダ作成
    Dim repoDir As String: repoDir = GetRootDir
    If repoDir = "" Then Exit Sub
    Call CreateDirIfThereNo(repoDir)

    ' ローカルリポジトリフォルダ内のサブフォルダ作成
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(repoDir & "¥.vscode") Then fso.CreateFolder repoDir & "¥.vscode"
    If Not fso.FolderExists(repoDir & "¥bin") Then fso.CreateFolder repoDir & "¥bin"
    If Not fso.FolderExists(repoDir & "¥bin¥old") Then fso.CreateFolder repoDir & "¥bin¥old"
    If Not fso.FolderExists(repoDir & "¥src") Then fso.CreateFolder repoDir & "¥src"
    Dim bookName As String: bookName = GetShortBookName(ActiveWorkbook.Name)
    Dim srcDir As String: srcDir = repoDir & "¥src"
    If Not fso.FolderExists(srcDir) Then fso.CreateFolder srcDir

    ' Git設定ファイル作成
    Call GenerateGitFiles(repoDir, srcDir)
    
    ' ブックを保存してbinフォルダにコピー
    ActiveWorkbook.Save
    Call fso.CopyFile(ActiveWorkbook.FullName, repoDir & "¥bin¥" & ActiveWorkbook.Name, True)
    Set fso = Nothing
    
    ' srcフォルダにCodeModuleをExport
    Call Decombine
    
    ' リモートリポジトリの作成
    If CreateRemoteRepos(bookName, repoName) Then
        MsgBox bookName & " 用のリポジトリの準備ができました。", vbInformation
    End If
End Sub

Public Function CreateRemoteRepos(bookName As String, repoName As String) As Boolean
    On Error GoTo Catch
    Dim rt As Boolean: rt = False
    
    ' トークンを取得
    Dim token As String: token = GetTokenFromRegistry()
    If Trim(token) = "" Then
        MsgBox "個人用アクセストークンを登録してください。", vbInformation
        CreateRemoteRepos = rt
        Exit Function
    End If

    ' HTTPオブジェクトを生成
    Dim http     As Object: Set http = CreateObject("MSXML2.XMLHTTP")

    ' GitHub APIのURL
    Dim url      As String: url = "https://api.github.com/user/repos"
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "token " & token

    ' JSONリクエストボディを作成（アクセス修飾子はprivate）
    Dim jsonBody As String: jsonBody = "{""name"":""" & repoName & """, ""private"": true}"

    ' リクエストを送信
    Call http.send(jsonBody)

    ' 結果を表示
    If http.Status = 201 Then
        Dim json    As Object: Set json = JsonConverter.ParseJson(http.responseText)
        Dim repoUrl As String: repoUrl = json("html_url")
        Call SaveSetting("Excel", bookName, "RepositoryURL", repoUrl)
        Call GitCmd(Init)
        MsgBox "リモートリポジトリが作成されました。" & vbCr & vbCr & repoUrl, vbInformation
        rt = True
    Else
        MsgBox "リモートリポジトリの作成に失敗しました。" & vbCr & vbCr & _
            "Status: " & http.Status & vbCr & http.responseText, vbExclamation
        rt = False
    End If
    GoTo Finally
Catch:
    MsgBox Err.Description, vbExclamation
    rt = False
Finally:
    Set http = Nothing
    Set json = Nothing
    CreateRemoteRepos = rt
End Function

' 引数のフォルダパスが存在しない場合に作る
Public Sub CreateDirIfThereNo(dirPath As String)

    Dim fso  As Object:  Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dirs As Variant: dirs = Split(dirPath, "¥")

    Dim i As Integer, dr As String
    For i = 0 To UBound(dirs)
        dr = dr & dirs(i) & "¥"
        ' 共有ネットワークパスを考慮
        If dr = "¥¥" Then
            i = i + 1
            dr = dr & dirs(i) & "¥"
        ElseIf Not fso.FolderExists(dr) Then
            fso.CreateFolder dr
        End If
    Next

    Set fso = Nothing

End Sub

' 引数の文字列でUTF-8テキストファイル作成（標準でBOM付きになる為、BOM除去）
Public Sub GenerateUTF8(txt As String, filePath As String)

    Dim binData As Variant, sm As Object
    
    Set sm = CreateObject("ADODB.Stream")
    With sm
        .Type = 2                ' 文字列
        .Charset = "utf-8"       ' 文字コード指定
        .Open                    ' 開く
        .WriteText txt           ' 文字列を書き込む
        .Position = 0            ' streamの先頭に移動
        .Type = 1                ' バイナリー
        .Position = 3            ' streamの先頭から3バイトをスキップ
        binData = .Read          ' バイナリー取得
        .Close: Set sm = Nothing ' 閉じて解放
    End With
    
    Set sm = CreateObject("ADODB.Stream")
    With sm
        .Type = 1                ' バイナリー
        .Open                    ' 開く
        .Write binData           ' バイナリーデータを書き込む
        .SaveToFile filePath, 2  ' 保存
        .Close: Set sm = Nothing ' 閉じて解放
    End With

End Sub

' Git設定ファイルの作成
Public Sub GenerateGitFiles(rootDir As String, srcDir As String)
    
    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub
    
    Dim txt As String, fileName As String
    With ThisWorkbook.VBProject.VBComponents(ContentsModuleName).CodeModule
        Dim i As Long, iLine As String
        For i = 1 To .CountOfLines
            iLine = .Lines(i, 1)
            ' 対象はコメント行のみ
            If UCase(Left(Trim(iLine), 3)) = "REM" Then
                fileName = Trim(Mid(Trim(iLine), 4))
                txt = ""
            ElseIf Left(Trim(iLine), 1) = "'" Then
                ' 先頭シングルクォーテーションの除去
                iLine = Mid(RTrim(iLine), 2)
                txt = txt & iLine & vbCrLf
            End If
            ' 空白行または最終行ならテキスト作成
            If (Trim(iLine) = "" Or i = .CountOfLines) And Trim(txt) <> "" Then
                If fileName = "settings.json" Then
                    Call GenerateUTF8(txt, rootDir & "¥.vscode¥" & fileName)
                Else
                    Call GenerateUTF8(txt, srcDir & "¥" & fileName)
                End If
            End If
        Next
    End With

End Sub

' プロジェクト内の全てのモジュールをsrcフォルダにExport
Public Sub ExportCodeModules(ByVal xBook As Workbook, ByVal srcDir As String)
    On Error GoTo Catch
    Dim vbPjt As VBIDE.VBProject: Set vbPjt = xBook.VBProject
    
    Dim vbCmp As VBIDE.VBComponent
    For Each vbCmp In vbPjt.VBComponents
        Select Case vbCmp.Type
            Case vbext_ct_StdModule
                vbCmp.Export srcDir & "¥" & vbCmp.Name & ".bas"
            Case vbext_ct_MSForm
                vbCmp.Export srcDir & "¥" & vbCmp.Name & ".frm"
            Case vbext_ct_ClassModule
                vbCmp.Export srcDir & "¥" & vbCmp.Name & ".cls"
            Case vbext_ct_Document
                vbCmp.Export srcDir & "¥" & vbCmp.Name & ".dcm"
        End Select
    Next
    GoTo Finally
Catch:
    OutputError "ExportCodeModules"
Finally:
    ' 何もしない
End Sub
 
' コマンドプロンプトでcmd引数を実行
Private Function RunCmd(cmd As String, Optional showInt As Integer = 0, Optional toWait As Boolean = True) As String
    On Error GoTo Catch
    Dim fso     As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim msgPath As String: msgPath = Environ$("temp") & "¥gitTmp.log"
    Dim errPath As String: errPath = Environ$("temp") & "¥gitErr.log"
    Dim wsh     As Object: Set wsh = CreateObject("WScript.Shell")
    Dim rt      As Long:   rt = wsh.Run("cmd /c " & cmd & " > " & msgPath & " 2> " & errPath, showInt, toWait)
    
    Dim msg As String
    If rt = 0 Then
        msg = CurDir & " (正常終了 - " & cmd & ")"
    Else
        msg = CurDir & " (異常終了 - " & cmd & ")"
    End If
    
    Dim msgStream As Object: Set msgStream = CreateObject("ADODB.Stream")
    msgStream.Type = 2
    msgStream.Charset = "utf-8"
    msgStream.Open
    msgStream.LoadFromFile msgPath
    Dim msgText As String
    msgText = msgStream.ReadText
    msgStream.Close: Set msgStream = Nothing
    
    Dim errStream As Object: Set errStream = CreateObject("ADODB.Stream")
    errStream.Type = 2
    errStream.Charset = "utf-8"
    errStream.Open
    errStream.LoadFromFile errPath
    Dim errText As String
    errText = errStream.ReadText
    errStream.Close: Set errStream = Nothing
    
    Select Case True
    Case Trim(msgText) = "" And Trim(errText) = ""
        ' 何もしない
    Case Trim(msgText) <> "" And Trim(errText) = ""
        msg = msg & vbCr & msgText
    Case Trim(msgText) = "" And Trim(errText) <> ""
        msg = msg & vbCr & errText
    Case Trim(msgText) <> "" And Trim(errText) <> ""
        msg = msg & vbCr & msgText & vbCr & errText
    End Select
    
    GoTo Finally
Catch:
    OutputError "RunCmd"
Finally:
    If fso.FileExists(msgPath) Then fso.DeleteFile msgPath
    If fso.FileExists(errPath) Then fso.DeleteFile errPath
    Set fso = Nothing
    Set wsh = Nothing
    RunCmd = Trim(msg)
End Function

' Gitコマンドを実行
Public Function GitCmd(cmd As GitCommand, Optional arg As String = Empty, Optional isPowerShell As Boolean = False) As Integer
    On Error GoTo Catch
    Dim rootDir As String: rootDir = GetRootDir
    If rootDir = "" Then
        MsgBox """" & ActiveWorkbook.Name & """" & vbLf & vbLf & "リポジトリ名が登録されていません。", vbInformation
        Exit Function
    End If
    
    Call ChDir(rootDir)
    
    Dim rt As String
    Select Case cmd
    Case Init
        Dim bookName As String: bookName = GetShortBookName(ActiveWorkbook.Name)
        Dim repoUrl  As String: repoUrl = GetSetting("Excel", bookName, "RepositoryURL")
        If repoUrl = "" Then
            MsgBox "リモートリポジトリを作成してください。", vbInformation
            GoTo Finally
        End If
        rt = RunCmd("git init")
        rt = rt & vbCr & RunCmd("git add .")
        rt = rt & vbCr & RunCmd("git commit -m ""リポジトリ開始""")
        rt = rt & vbCr & RunCmd("git branch -M main")
        rt = rt & vbCr & RunCmd("git remote add origin " & repoUrl)
        rt = rt & vbCr & RunCmd("git push -u origin main")
    Case Status
        rt = RunCmd("git status")
    Case Stage
        If MsgBox(ActiveWorkbook.Name & " の変更をステージします。" & vbLf & vbLf & _
                  ActiveWorkbook.Name & " の保存とエクスポートを伴います。", vbInformation + vbOKCancel) = vbOK Then
            Application.DisplayAlerts = False
            If ActiveWorkbook.path = rootDir & "¥bin" Then
                MsgBox "binフォルダ内の" & ActiveWorkbook.Name & "を開いたままステージ出来ません。" & vbLf & _
                       "ステージはキャンセルされました。, vbInformation"
                GoTo Finally
            Else
                ActiveWorkbook.Save
                Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
                Call fso.CopyFile(ActiveWorkbook.FullName, rootDir & "¥bin¥" & ActiveWorkbook.Name, True)
                Call Decombine
                rt = RunCmd("git add .")
            End If
        Else
            GoTo Finally
        End If
    Case Commit
        If arg = Empty Then
            arg = InputBox("コミットのメッセージを入力してください。")
            If arg = "" Then GoTo Finally
            If MsgBox("""" & arg & """" & vbLf & vbLf & "このメッセージでコミットします。", vbInformation + vbOKCancel) = vbCancel Then
                GoTo Finally
            End If
        End If
        rt = RunCmd("git commit -m """ & arg & """")
    Case Push
        Dim mBranch As String
        If arg = Empty Then mBranch = "main"
        rt = RunCmd("git push origin " & mBranch)
    End Select
    Debug.Print rt
    GoTo Finally
Catch:
    OutputError "GitCmd"
Finally:
    Application.DisplayAlerts = True
End Function

' ---メニュー用 ここから------------------------------------
'
Public Sub GitStage()
    Call GitCmd(Stage)
End Sub

Public Sub GitCommit()
    Call GitCmd(Commit)
End Sub

Public Sub GitPush()
    Call GitCmd(Push)
End Sub
'
' ---メニュー用 ここまで------------------------------------

' ルートディレクトリを返す
Private Function GetRootDir() As String
    Dim bookName  As String: bookName = GetShortBookName(ActiveWorkbook.Name)
    Dim repoName As String: repoName = GetSetting("Excel", bookName, "RepositoryName")
    If repoName = "" Then
        GetRootDir = ""
        Exit Function
    End If
    GetRootDir = Environ$("USERPROFILE") & "¥" & parentDir & "¥" & repoName
End Function


' srcフォルダを指定してActiveWorkbookをExport
Public Sub Decombine(Optional includeBookName As Boolean = False)
    Dim srcPath As String
    Dim rootDir As String: rootDir = GetRootDir
    If rootDir = "" Then Exit Sub
    If includeBookName Then
        srcPath = rootDir & "¥src¥" & GetShortBookName(ActiveWorkbook.Name)
    Else
        srcPath = rootDir & "¥src"
    End If
    Call CreateDirIfThereNo(srcPath)
    Call ExportCodeModules(ActiveWorkbook, srcPath)
End Sub

' リポジトリ名をレジストリに記録
Public Function SetAndThenGetReposName() As String
    Dim bookName As String: bookName = GetShortBookName(ActiveWorkbook.Name)
    Dim repoName As String: repoName = GetSetting("Excel", bookName, "RepositoryName")
    If repoName = "" Then
        repoName = InputBox("リポジトリ名を英字で入力してください。")
        If repoName = "" Then
            SetAndThenGetReposName = ""
            Exit Function
        End If
        If Not IsValidRepoName(repoName) Then
            MsgBox "リポジトリ名は無効です。" & vbCr & vbCr & _
                "リポジトリ名は英字で始まり、小文字、数字、ハイフン、アンダースコア、" & vbCr & _
                "ピリオドを含めることができ、最大256文字までです。" & vbCr & _
                "連続するハイフン、アンダースコアは使用できません。", vbInformation
            SetAndThenGetReposName = ""
            Exit Function
        End If
        Call SaveSetting("Excel", bookName, "RepositoryName", repoName)
    End If
    SetAndThenGetReposName = repoName
End Function

' 最初のアンダースコアから最後のドットまでの間の文字列を消去する
Private Function GetShortBookName(bookName As String) As String
    
    ' アンスコの位置
    Dim unsPos As Integer: unsPos = InStr(bookName, "_")
    ' ドットの位置
    Dim dotPos As Integer: dotPos = InStrRev(bookName, ".")

    If unsPos = 0 Or dotPos = 0 Then
        ' アンスコまたはドットが見つからない場合、元のファイル名を返す
        GetShortBookName = bookName
    Else
        ' アンスコの前の部分と、最後のドットの後の部分を結合
        GetShortBookName = Left(bookName, unsPos - 1) & Mid(bookName, dotPos)
    End If
End Function


' レジストリにトークンを登録
Public Sub RegisterToken()
    On Error GoTo Catch
    Dim keyStr As String
    keyStr = InputBox("GitHubの個人アクセストークンを入力してください。")
    If keyStr = "" Then Exit Sub
    Call SaveSetting("GitHub", "Token", "Classic", keyStr)
    MsgBox "GitHubの個人アクセストークンを登録しました。", vbInformation
    Exit Sub
Catch:
    MsgBox Err.Description, vbExclamation
End Sub

' レジストリからトークンを得る
Public Function GetTokenFromRegistry() As String
    GetTokenFromRegistry = GetSetting("GitHub", "Token", "Classic")
End Function

' レジストリからトークン用のキーを削除
Public Sub DeleteToken()
    Call DeleteSetting("GitHub", "Token", "Classic")
End Sub

' GitHubのリポジトリ名の有効性をチェックする
Function IsValidRepoName(repoName As String) As Boolean
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")

    ' リポジトリ名は英字で始まり、指定された文字のみ含む、最大256文字
    With regEx
        .Pattern = "^[a-zA-Z][-a-zA-Z0-9_.]*$"
        .IgnoreCase = False
        .Global = False
    End With

    ' リポジトリ名が空、長すぎる、または正規表現に一致しない場合は無効
    If Len(repoName) = 0 Or Len(repoName) > 256 Or Not regEx.Test(repoName) Then
        IsValidRepoName = False
    Else
        ' 連続するハイフン、アンダースコアをチェック
        If InStr(repoName, "--") > 0 Or InStr(repoName, "__") > 0 Then
            IsValidRepoName = False
        Else
            IsValidRepoName = True
        End If
    End If
    
    Set regEx = Nothing
End Function




' 以下、記事には不要 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub Cls()
    Application.SendKeys "^g", True
    If Application.VBE.ActiveWindow.Caption = "イミディエイト" Then
        Application.SendKeys "^a", True
        Application.SendKeys "{Del}", True
    End If
End Sub

' HKEY_CURRENT_USER¥Software¥VB and VBA Program Setting

Sub ImportDocument(path As String, xlBook As Workbook)
    Dim compos As VBComponents
    Set compos = xlBook.VBProject.VBComponents
    
    Dim impCompo As VBComponent
    Set impCompo = compos.Import(path)
    
    Dim origCompo As VBComponent
    Dim cname As String, bname As String
    cname = impCompo.Name
    bname = GetNameFromPath(path)
    
    If cname <> bname Then
        Set origCompo = compos.Item(bname)
    Else
        Dim sht As Worksheet
        Set sht = xlBook.Worksheets.Add()
        Set origCompo = compos.Item(sht.CodeName)
        
        Dim tmpname As String
        tmpname = "ImportTemp"
        While ComponentExists(compos, tmpname)
            tmpname = tmpname & "1"
        Wend
        
        impCompo.Name = tmpname
        origCompo.Name = cname
    End If
    
    Dim imod As CodeModule, omod As CodeModule
    Set imod = impCompo.CodeModule
    Set omod = origCompo.CodeModule
    omod.DeleteLines 1, omod.CountOfLines
    omod.AddFromString imod.Lines(1, imod.CountOfLines)
    
    compos.Remove impCompo
End Sub

Function ComponentExists(compos As VBComponents, Name As String) As Boolean
    Dim c As VBComponent
    On Error Resume Next
    Set c = compos.Item(Name)
    If Err.Number = 0 Then
        ComponentExists = True
    Else
        ComponentExists = False
    End If
    On Error GoTo 0
End Function

Function GetNameFromPath(path As String) As String
    GetNameFromPath = Mid(path, InStrRev(path, "¥") + 1, Len(path))
End Function

Public Sub GitInit()
    On Error GoTo Catch
    Dim urlStr As String: urlStr = InputBox("リモートリポジトリのURLを入力してください。")
    If urlStr = "" Then GoTo Finally
    Dim repoName As String: repoName = ExtractProjectName(urlStr)
    Dim fileName As String: fileName = GetShortBookName(ActiveWorkbook.Name)
    Call SaveSetting("Excel", fileName, "RepositoryName", repoName)
    Dim reposPath As String: reposPath = Environ$("USERPROFILE") & "¥" & parentDir & "¥" & repoName
    Call ChDir(reposPath)
            
    Dim cmdStr As String: cmdStr = "echo # & " & repoName & " >> README.md"
    Dim rt As Long: rt = RunCmd(cmdStr)
    rt = GitCmd(Stage)
    rt = GitCmd(Commit, "リポジトリ開始")
    cmdStr = "git branch -M main"
    rt = RunCmd(cmdStr)
    cmdStr = "git remote add origin " & urlStr
    rt = RunCmd(cmdStr)
    rt = GitCmd(Push)
    GoTo Finally
Catch:
    OutputError "GitInit"
Finally:
    
End Sub

Function ExtractProjectName(fullURL As String) As String
    Dim startCut As Integer: startCut = InStrRev(fullURL, "/") + 1
    Dim endCut   As Integer: endCut = InStr(1, fullURL, ".git")
    ExtractProjectName = Mid(fullURL, startCut, endCut - startCut)
End Function

' サブフォルダを作成する
Function CreateSubDir(parentDir As String, subDir As String) As String
    
    Dim fPath As String: fPath = parentDir & "¥" & subDir
    
    ' フォルダが存在するか確認
    If Dir(fPath, vbDirectory) = "" Then
        ' 存在しない場合、フォルダを作成
        Call MkDir(fPath)
        CreateSubDir = fPath
        Exit Function
    End If
    
    ' 連番付きのフォルダ名を生成
    Dim cnt As Integer: cnt = 1
    Dim newSubDir As String
    Do
        newSubDir = subDir & " (" & cnt & ")"
        fPath = parentDir & "¥" & newSubDir
        
        If Dir(fPath, vbDirectory) = "" Then
            MkDir fPath
            CreateSubDir = fPath
            Exit Function
        End If
        
        cnt = cnt + 1
    Loop
End Function

Sub CheckEncoding()
    Dim en As Object: Set en = CreateObject("System.Text.UTF8Encoding")
'    Dim sjis As Object: Set sjis = en.GetEncoding("shift_jis")
    Dim bin As Variant: bin = en.GetBytes_4("依頼NO.")
    Dim deco As Object: Set deco = en.GetDecoder

End Sub


' 作業用一時ファイル名
'Private Const TempFileName = "TempGitOutput"
'
'Private Function RunPowerShell(cmd As String, Optional showInt As Integer = 0, Optional toWait As Boolean = False) As String
'
'    Dim tmpPath As String: tmpPath = Environ$("temp") & "¥" & TempFileName
'    Call GenerateUTF8(" ", tmpPath)
'
'    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
'    '  -NoLogo (見出しを出さない)
'    '  -ExecutionPolicy RemoteSigned (実行権限を設定)
'    '  -Command (PowerShellのコマンドレット構文を記載）
'    Call wsh.Run("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command """ & cmd & _
'                 " | Out-File -filePath " & tmpPath & " -encoding utf8""", showInt, toWait)
'
'    Dim sm As Object: Set sm = CreateObject("ADODB.Stream")
'    sm.Type = 2
'    sm.Charset = "utf-8"
'    sm.Open
'    sm.LoadFromFile tmpPath
'    Dim tmp As String: tmp = sm.ReadText
'    sm.Close: Set sm = Nothing
'
'    RunPowerShell = tmp
'
'End Function

' リポジトリ名に禁止文字が使われていないかどうかチェック
Private Function CheckReposName(ByVal stg As String) As String
    Dim i As Integer
    For i = 1 To Len(stg)
        Select Case Asc(Mid(stg, i, 1))
        Case 0 To 127
            If InStr("/¥@‾ ", Mid(stg, i, 1)) > 0 Or _
                (i = Len(stg) And Mid(stg, i, 1) = ".") Then _
                    GoTo Invalid
        Case Else
            GoTo Invalid
        End Select
    Next
    CheckReposName = stg
    Exit Function
Invalid:
    CheckReposName = ""
End Function

