Attribute VB_Name = "Module2"
Sub GetSalesforceAccessToken()
    Dim xmlhttp As Object
    Dim response As String
    Dim accessToken As String
    Dim clientId As String
    Dim clientSecret As String
    Dim username As String
    Dim password As String
    Dim securityToken As String

    ' 接続アプリケーションのクライアントIDとクライアントシークレット
    clientId = Range("E5").Value
    clientSecret = Range("F5").Value

    ' Salesforceユーザーのログイン情報
    username = Range("B5").Value
    password = Range("C5").Value
    securityToken = Range("D5").Value

    ' Salesforce REST APIエンドポイント（トークン取得用）
    tokenUrl = "https://login.salesforce.com/services/oauth2/token"

    ' HTTPリクエストを作成
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlhttp.Open "POST", tokenUrl, False
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

    ' リクエストボディを設定
    requestBody = "grant_type=password" & _
                  "&client_id=" & clientId & _
                  "&client_secret=" & clientSecret & _
                  "&username=" & username & _
                  "&password=" & password & securityToken

    ' リクエストを送信
    xmlhttp.send requestBody

    ' レスポンスを取得
    response = xmlhttp.responseText

    ' レスポンスからアクセストークンを抽出
    accessToken = JsonValue(response, "access_token")

    ' アクセストークンをセルに入力
    Range("G5").Value = accessToken
    
    ' ログイン成功メッセージの表示
    If Range("G5").Value = "" Then
        MsgBox "ログイン失敗"
    Else
        MsgBox "ログイン成功"
    End If
    
    
    ' クリーンアップ
    Set xmlhttp = Nothing
End Sub

Function JsonValue(JsonString As String, key As String) As String
    Dim regex As Object
    Dim matches As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = """" & key & """:" & """(.*?)"""

    Set matches = regex.Execute(JsonString)
    If matches.Count > 0 Then
        JsonValue = matches(0).SubMatches(0)
    Else
        JsonValue = ""
    End If
End Function



