VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomersMailCloud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ApiUser As String
Public ApiKey As String
Private Urls(3) As String
Public Url As String
Public Subject As String
Public Text As String
Private toAddress() As clsCustomersMailCloudAddress
Public fromAddress As clsCustomersMailCloudAddress

Private Sub Class_Initialize()
    Urls(0) = "https://sandbox.smtps.jp/api/v2/emails/send.json"
    Urls(1) = "https://te.smtps.jp/api/v2/emails/send.json"
    Urls(2) = "https://SUBDOMAIN.smtps.jp/api/v2/emails/send.json"
    Set fromAddress = New clsCustomersMailCloudAddress
End Sub


Public Sub Trial()
    Url = Urls(0)
End Sub

Public Sub Standard()
    Url = Urls(1)
End Sub

Public Sub Pro(subDomain As String)
    Url = Replace(Urls(2), "SUBDOMAIN", subDomain)
End Sub

Public Sub AddTo(Address As clsCustomersMailCloudAddress)
    ' Debug.Print toAddress
    'If toAddress = None Then
    '    ReDim Preserve toAddress(0)
    'End
    If Not Not toAddress Then
        ReDim Preserve toAddress(UBound(toAddress))
        Set toAddress(UBound(toAddress)) = Address
    Else
        ReDim toAddress(0) As clsCustomersMailCloudAddress
        Set toAddress(0) = Address
    End If
End Sub

Public Sub SetFrom(Address As clsCustomersMailCloudAddress)
    Set fromAddress = Address
End Sub

Public Function Send() As Boolean
    Dim params As New Dictionary
    
    params.Add "api_user", ApiUser
    params.Add "api_key", ApiKey
    params.Add "subject", Subject
    params.Add "text", Text

    
    Dim i As Integer
    Dim toParams() As Dictionary
    For i = LBound(toAddress) To UBound(toAddress)
        If Not Not toParams Then
            ReDim Preserve toParams(UBound(toParams))
        Else
            ReDim toParams(0) As Dictionary
        End If
        Set toParams(UBound(toParams)) = New Dictionary
        toParams(i).Add "name", toAddress(i).Name
        toParams(i).Add "address", toAddress(i).Address
    Next i
    params.Add "to", toParams
    
    Dim fromParams As New Dictionary
    fromParams.Add "name", fromAddress.Name
    fromParams.Add "address", fromAddress.Address
    params.Add "from", fromParams
    
    Dim httpReq As Object
    Set httpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    httpReq.Open "POST", Url, False
    httpReq.setRequestHeader "Content-Type", "application/json"
    
    httpReq.Send (JsonConverter.ConvertToJson(params))
    Send = True
End Function


