Imports System.Diagnostics

Sub Main()
    Dim process As New Process()

    ' StartInfoプロパティを設定します。
    process.StartInfo.FileName = "C:\Program Files\Git\bin\bash.exe" ' あなたのGit Bashのパスに置き換えてください。
    process.StartInfo.Arguments = "-c ""cd /path/to/your/repo; git pull; git status""" ' 実行したいGitコマンドに置き換えてください。
    process.StartInfo.UseShellExecute = False
    process.StartInfo.RedirectStandardOutput = True
    process.StartInfo.CreateNoWindow = True

    ' OutputDataReceived イベントハンドラを追加します。
    AddHandler process.OutputDataReceived, Sub(sender As Object, e As DataReceivedEventArgs)
                                               If e.Data IsNot Nothing Then
                                                   Console.WriteLine(e.Data)
                                               End If
                                           End Sub

    ' プロセスを開始します。
    process.Start()

    ' 非同期で標準出力を読み始めます。
    process.BeginOutputReadLine()

    ' プロセスが終了するのを待ちます。
    process.WaitForExit()
End Sub
