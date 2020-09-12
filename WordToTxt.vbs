'***** Wordファイルをコピーし、テキストファイルを作成するVBScript *****

Option Explicit

'ドラッグアンドドロップしたファイルの絶対パスを格納
Dim GetPathArray
Set GetPathArray = WScript.Arguments
'ファイルシステムオブジェクト
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'イテレータ
Dim pt

'作成するテキストファイルの絶対パス
Dim strTextFilePath
'Wordのオブジェクト
Dim objWordApp
Dim objWordDoc

'ファイルの数だけループする
For Each pt in GetPathArray

    'Wordの拡張子を除去し、txtをくっつける
    strTextFilePath = Left(pt, Len(pt) - 3) & "txt"

    'ワードのオブジェクトを作成
    Set objWordApp = WScript.CreateObject("Word.Application")

    'エラーが発生しなかった場合
    If Err.Number = 0 Then
        'ワードドキュメントを開く
        Set objWordDoc = objWordApp.Documents.Open(pt)

        'エラーが発生しなかった場合
        If Err.Number = 0 Then
            'テキスト形式で保存
            objWordDoc.SaveAs strTextFilePath, 2

            objWordDoc.Close
            objWordApp.Quit
        Else
            WScript.Echo "エラー：" & Err.Descripticon
        End If
    Else
        WScript.Echo "エラー：" & Err.Descripticon
    End If
Next

'オブジェクト変数をクリア
Set objFSO = Nothing
Set objWordApp = Nothing
Set objWordDoc = Nothing