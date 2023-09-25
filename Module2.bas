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

    ' �ڑ��A�v���P�[�V�����̃N���C�A���gID�ƃN���C�A���g�V�[�N���b�g
    clientId = Range("E5").Value
    clientSecret = Range("F5").Value

    ' Salesforce���[�U�[�̃��O�C�����
    username = Range("B5").Value
    password = Range("C5").Value
    securityToken = Range("D5").Value

    ' Salesforce REST API�G���h�|�C���g�i�g�[�N���擾�p�j
    tokenUrl = "https://login.salesforce.com/services/oauth2/token"

    ' HTTP���N�G�X�g���쐬
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlhttp.Open "POST", tokenUrl, False
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

    ' ���N�G�X�g�{�f�B��ݒ�
    requestBody = "grant_type=password" & _
                  "&client_id=" & clientId & _
                  "&client_secret=" & clientSecret & _
                  "&username=" & username & _
                  "&password=" & password & securityToken

    ' ���N�G�X�g�𑗐M
    xmlhttp.send requestBody

    ' ���X�|���X���擾
    response = xmlhttp.responseText

    ' ���X�|���X����A�N�Z�X�g�[�N���𒊏o
    accessToken = JsonValue(response, "access_token")

    ' �A�N�Z�X�g�[�N�����e�L�X�g�t�@�C���ɓ���
    ' �e�L�X�g�t�@�C�����������݂ŊJ���i������ΐV�K�쐬�����A����ꍇ�͏㏑�������j
    Open Range("G5").Value & "\accessToken.txt" For Output As #1
    
    ' �J�����t�@�C���ɏ�������
    Print #1, accessToken
    
    ' �J�����t�@�C�������
    Close #1
    
    ' ���b�Z�[�W�̕\��
    MsgBox "�������I�����܂����B"
    
    
    ' �N���[���A�b�v
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



