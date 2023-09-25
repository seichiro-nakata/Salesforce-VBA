Attribute VB_Name = "Module3"
Sub SalesforceRESTAPICall()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Salesforce REST APIのエンドポイントURLを設定
    url = "https://intloop22-dev-ed.develop.my.salesforce.com/services/data/v53.0/sobjects/Account"
    
    ' HTTP GETリクエストを作成
    request.Open "GET", url, False
    
    ' テキストファイルからアクセストークンを取得
    Dim i As String
    Open Range("G5").Value & "\accessToken.txt" For Input As #1
        Line Input #1, i
    Close #1
    
    ' OAuth2.0のアクセストークンをヘッダーに追加
    request.SetRequestHeader "Authorization", "Bearer " & i
    
    ' リクエストを送信
    request.Send
    
    ' レスポンスを表示
    MsgBox request.responseText
    
End Sub

