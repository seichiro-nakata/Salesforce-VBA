Attribute VB_Name = "Module3"
Sub SalesforceRESTAPICall()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Salesforce REST API�̃G���h�|�C���gURL��ݒ�
    url = "https://intloop22-dev-ed.develop.my.salesforce.com/services/data/v53.0/sobjects/Account"
    
    ' HTTP GET���N�G�X�g���쐬
    request.Open "GET", url, False
    
    ' �e�L�X�g�t�@�C������A�N�Z�X�g�[�N�����擾
    Dim i As String
    Open Range("G5").Value & "\accessToken.txt" For Input As #1
        Line Input #1, i
    Close #1
    
    ' OAuth2.0�̃A�N�Z�X�g�[�N�����w�b�_�[�ɒǉ�
    request.SetRequestHeader "Authorization", "Bearer " & i
    
    ' ���N�G�X�g�𑗐M
    request.Send
    
    ' ���X�|���X��\��
    MsgBox request.responseText
    
End Sub

