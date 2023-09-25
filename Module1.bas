Attribute VB_Name = "Module1"
Sub GetAccountsFromSalesforce()
    Dim url As String
    Dim xmlhttp As Object
    Dim response As String
    Dim accessToken As String

    ' Salesforce REST API�G���h�|�C���g�i�I�u�W�F�N�g����ύX���邱�Ƃ��ł��܂��j
    url = "https://intloop22-dev-ed.develop.my.salesforce.com/services/data/v53.0/query/?q=SELECT+Id,Name+FROM+Account"

    ' Salesforce����擾�����A�N�Z�X�g�[�N�����Z�b�g
    accessToken = "00D5j00000BRFTJ!ARYAQN41xtNU8jatoqMRGTun1A4oq49c4joE3IUMvXhvmyM5hu9Atw5CtWZR09mcIIyIvwU81OaCwu.5CpJ_Ow5GVyLZfDQa"

    ' HTTP���N�G�X�g���쐬
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "Authorization", "Bearer " & accessToken

    ' ���N�G�X�g�𑗐M
    xmlhttp.send ""

    ' ���X�|���X���擾
    response = xmlhttp.responseText
    
    ' JSON����������
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson(response)
    
    ' records�v���p�e�B���烌�R�[�h�̃R���N�V�������擾
    Set Data = jsonObject("records")
    
    ' ���R�[�h�̃R���N�V�����𔽕�����
    Dim id As String
    For Each Record In Data
        ' �e���R�[�h����Id�v���p�e�B�擾(���ږ���ύX����ƃf�[�^�̎��o���\)
        id = Record("Name")
        
        ' ID��\��
        MsgBox "ID: " & id
    Next Record

    ' ���X�|���X�̏����i�����ł̓��b�Z�[�W�{�b�N�X�ɕ\���j
    MsgBox response


    ' �N���[���A�b�v
    Set xmlhttp = Nothing
End Sub


