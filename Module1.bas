Attribute VB_Name = "Module1"
Sub GetAccountsFromSalesforce()
    Dim url As String
    Dim xmlhttp As Object
    Dim response As String
    Dim accessToken As String

    ' Salesforce REST APIエンドポイント（オブジェクト名を変更することができます）
    url = "https://intloop22-dev-ed.develop.my.salesforce.com/services/data/v53.0/query/?q=SELECT+Id,Name+FROM+Account"

    ' Salesforceから取得したアクセストークンをセット
    accessToken = "00D5j00000BRFTJ!ARYAQN41xtNU8jatoqMRGTun1A4oq49c4joE3IUMvXhvmyM5hu9Atw5CtWZR09mcIIyIvwU81OaCwu.5CpJ_Ow5GVyLZfDQa"

    ' HTTPリクエストを作成
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "Authorization", "Bearer " & accessToken

    ' リクエストを送信
    xmlhttp.send ""

    ' レスポンスを取得
    response = xmlhttp.responseText
    
    ' JSON文字列を解析
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson(response)
    
    ' recordsプロパティからレコードのコレクションを取得
    Set Data = jsonObject("records")
    
    ' レコードのコレクションを反復処理
    Dim id As String
    For Each Record In Data
        ' 各レコードからIdプロパティ取得(項目名を変更するとデータの取り出し可能)
        id = Record("Name")
        
        ' IDを表示
        MsgBox "ID: " & id
    Next Record

    ' レスポンスの処理（ここではメッセージボックスに表示）
    MsgBox response


    ' クリーンアップ
    Set xmlhttp = Nothing
End Sub


