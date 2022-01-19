Attribute VB_Name = "SearchRoutes"
Option Explicit

'ユーザーフォーム変数
Public formSearchParam As SearchParameters

'検索開始ボタンクリック
Public Function StartSearch()
    '検索フォームのインスタンスを作成
    Set formSearchParam = New SearchParameters
    'フォーム表示
    formSearchParam.Show
End Function

'APIを使ってルート検索
Public Function GetRouteFromGoogleMapsAPI(ByVal sOrigin As String, ByVal sDest As String, ByVal sCityName As String)
    Dim tblAPIkey As ListObject 'APIキーテーブル
    Dim sAPIkey As String   'APIキー
    Dim oHttpReq As XMLHTTP60   'APIレスポンス
    Dim dictResponse As Scripting.Dictionary 'APIレスポンス整形語のデータ
    Dim sResponse As String     'レスポンスデータ
    Dim oReg As RegExp      '正規表現
    Dim oMatch As Object        '検索結果
    Dim oEachMatch As Variant       '各検索結果
    Dim oKey As Variant         'DictionaryのKey
    
    'オブジェクトの初期化
    Set oHttpReq = New XMLHTTP60
    Set dictResponse = New Scripting.Dictionary
    Set oReg = New RegExp
    
    'APIキーを取得
    Set tblAPIkey = ThisWorkbook.Sheets("API_KEY").ListObjects("APIKEYテーブル")
    sAPIkey = tblAPIkey.ListColumns("APIキー").DataBodyRange(1).Value
    
    'APIリクエストを送る
    oHttpReq.Open "GET", "https://maps.googleapis.com/maps/api/directions/json?origin=" & sOrigin & "&destination=" & sDest & "+" & sCityName & "&key=" & sAPIkey
    oHttpReq.send
    'APIコール実行待ち
    Do While oHttpReq.readyState < 4
        DoEvents
    Loop
    
    'APIレスポンス確認
    sResponse = oHttpReq.responseText
    
    '正規表現設定
    oReg.Pattern = "[""html_instructions""].+,\n"        '正規表現パターン
    oReg.IgnoreCase = False     '大文字と小文字の区別
    oReg.Global = True     '文章全体を検索する
    
    '文章検索
    Set oMatch = oReg.Execute(sResponse)
    
    '検索結果を辞書に格納
    For Each oEachMatch In oMatch
        dictResponse.Add oEachMatch, 1
    Next
    
    '検索結果から余計な文字列を削除
    For Each oKey In dictResponse.Keys
        If InStr(1, oKey, "html_instructions", vbTextCompare) > 0 Then
            oKey = Left(oKey, InStr(1, oKey, """,", vbTextCompare) - 1)
            oKey = Right(oKey, Len(oKey) - InStr(1, oKey, ": """, vbTextCompare) - 2)
            oKey = ReplaceAllTargetStrings(oKey, "\u003cb\u003e", "")
            oKey = ReplaceAllTargetStrings(oKey, "\u003c/b\u003e", "")
            oKey = ReplaceAllTargetStrings(oKey, "/\u003cwbr/\u003e", "")
            oKey = ReplaceAllTargetStrings(oKey, "\u003cdiv style=\""font-size:0.9em\""\u003", "")
            oKey = ReplaceAllTargetStrings(oKey, "\u003c/div\u003e", "")
            'ユーザーフォームに検索結果を表示
            If formSearchParam.ResultBox.Value <> "" Then
                formSearchParam.ResultBox.Value = formSearchParam.ResultBox.Value & vbCrLf & oKey
            Else
                formSearchParam.ResultBox.Value = oKey
            End If
        Else
            dictResponse.Remove oKey
        End If
    Next
    
End Function

'文字列を1行ずつ取得
Public Function GetEachLine(ByVal sInput As String) As Scripting.Dictionary
    Dim sline As String
    Dim s As String
    Dim i As Long
    Dim lr As Long
    
    lr = 12
    For i = 1 To Len(sInput)
        s = Mid(sInput, i, 1)
        If s = vbLf Then
            Cells(lr, 2) = sline
            sline = ""
            lr = lr + 1
        ElseIf s <> vbCr Then
            sline = sline & s
        End If
    Next
End Function

'文字列を一括置換する
Public Function ReplaceAllTargetStrings(ByVal sInput As String, ByVal sFindText As String, ByVal sReplaceText As String) As String
    Do While True
        If InStr(1, sInput, sFindText) > 0 Then
            sInput = Replace(sInput, sFindText, sReplaceText)
        Else
            Exit Do
        End If
    Loop
    ReplaceAllTargetStrings = sInput
End Function
