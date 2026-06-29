VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe_Signature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Dim FXMLDoc As DOMDocument50
Const DS = "xmlns:ds='http://www.w3.org/2000/09/xmldsig#'"

Function Assina(ByVal mm_xml As String, ByVal mSerial As String) As String
Dim doc As New DOMDocument50
Dim txt As String
Dim FileName As String
    On Error GoTo fim
    doc.validateOnParse = False
    doc.async = False
    doc.preserveWhiteSpace = False
    Assina = ""
    If doc.LoadXml(mm_xml) Then
        FileName = Format(Time, "hh-mm-ss")
        If SignXML(doc, mSerial) Then
            Assina = doc.xml
        End If
    Else
        MsgBox "XML não carregado, " & doc.parseError
    End If
    Exit Function
fim:
    MsgBox Err.Description
End Function

'//*****************************************************
'// RemoveSignatures:
'//  - Remover os tags Signature do XML
'//*****************************************************
Private Function RemoveSignatures(d As DOMDocument50) As Boolean
Dim oSigs As IXMLDOMNodeList
    Call d.SetProperty("SelectionNamespaces", DS)
    Set oSigs = d.selectNodes(".//ds:Signature")
    Do While (oSigs.Length <> 0)
        oSigs(0).parentNode.removeChild (oSigs(0))
        Set oSigs = d.selectNodes(".//ds:Signature")
    Loop
    Set oSigs = d.selectNodes(".//Signature")
    Do While (oSigs.Length <> 0)
        oSigs(0).parentNode.removeChild (oSigs(0))
        Set oSigs = d.selectNodes(".//Signature")
    Loop
    RemoveSignatures = (oSigs.Length = 0)
End Function

Private Function AddSign(oDoc As DOMDocument50, N As IXMLDOMNode, Cert As Certificate, Store As Store) As Boolean
Dim sProvider As String
Dim sContainer As String
    'Dim oDoc As DOMDocument
Dim pKey As IXMLDSigKey, pKeyOut As IXMLDSigKey
Dim PrivateKey As PrivateKey
Dim oDSig As New MSXML2.MXDigitalSignature50


    oDoc.preserveWhiteSpace = False
    'Set oDoc = d
    Set oDSig.Signature = N
    Set PrivateKey = Cert.PrivateKey

    '//Para criar uma KEY deve ser informado o Provider e o Container
    sProvider = Cert.PrivateKey.ProviderName
    sContainer = Cert.PrivateKey.ContainerName

    Set pKey = oDSig.createKeyFromCSP(CAPICOM_PROV_RSA_FULL, sProvider, sContainer, 0)
    Set oDSig.Signature = N
    Set oDSig.Store = Store
    Set pKeyOut = oDSig.Sign(pKey, Certificates Or PURGE)



End Function

Private Function SignXML(doc As DOMDocument50, ByVal mSerial As String) As Boolean
'  sCert: array [Boolean] of string = ('s',  '');
'  sDisp: array [Boolean] of string = ('is', 'l');
Dim i As Long        ',: Cardinal;
Dim oNode As IXMLDOMNode
Dim SetT As New Settings, Certs, StoreSrc As New Store, StoreDst As New Store, Cert As Certificate        ': OleVariant;
Dim oRps As IXMLDOMNodeList, oLote As IXMLDOMNodeList, oSigs As IXMLDOMNodeList
Dim s1 As String, s2 As String

    'Sett = CoSettings.Create
    SetT.EnablePromptForCertificateUI = True
    'StoreSrc = CoStore.Create
    Call StoreSrc.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_EXISTING_ONLY)
    'StoreDst = CoStore.Create
    Call StoreDst.Open(CAPICOM_CURRENT_USER_STORE, "TMP2", CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED)

    For i = StoreDst.Certificates.Count To 1 Step -1
        StoreDst.Remove StoreDst.Certificates(i)
    Next
    If StoreDst.Certificates.Count > 0 Then
        MsgBox "erro"
    End If

    'Set StoreDst = New Store
    Set Certs = StoreSrc.Certificates

    '//Remove certificados sem a private key.
    If Certs.Count > 0 Then
        Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY, CAPICOM_PROPID_KEY_PROV_INFO)
    End If
    '//Somente certificados com data válida.
    If Certs.Count > 0 Then
        Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID)
    End If
    'MsgBox Certs.item(1).SubjectName
    'Certs.Select


    If Certs.Count = 0 Then
        MsgBox "Sem certificados"
    Else
        Set Cert = Certs.Item(1)
        For i = Certs.Count To 1 Step -1
            If Certs.Item(i).SerialNumber = mSerial Then
                Set Cert = Certs.Item(i)
                Debug.Print Certs.Item(i).PrivateKey
                'Debug.Print Certs.Item(i).PublicKey
                Debug.Print Certs.Item(i).SubjectName
                Debug.Print Certs.Item(i).HasPrivateKey
                Exit For
            End If


        Next



    End If
   
    If Not Cert Is Nothing Then
        'begin
        '//Limpa do Store os Outros Certificados deixando só o selecionado
        '//OBS: Importante só haver 1 certificado no Store Destino
        '//     senão o XML ficará com vários X509DATA
        StoreDst.Certificates.Clear
        StoreDst.Add Cert        'Certs.item(1)
        'MsgBox Cert.SubjectName
        '// Exclui todas nodes <Signature>.
        'RemoveSignatures Doc

        'doc.preserveWhiteSpace = False
        Set FXMLDoc = doc
        'FXMLDoc.preserveWhiteSpace = False
        '// Encontra o node <LoteRps>.
        'MsgBox FXMLDoc.xml
        Set oLote = FXMLDoc.selectNodes(".//NFe")
        'MsgBox FXMLDoc.xml
        If (oLote.Length = 0) Then
            'MsgBox "Nenhuma Tag Lote foi encontrada no XML.')"
            'Exit Function
        End If
        '//Adiciona Tag de Assinatura para Lote

        'Call AddDSig(Doc, oLote.item(0))

        '// Encontra os nodes <InfRps>.
        Set oRps = doc.selectNodes(".//RPS")
        If (oRps.Length = 0) Then
            'MsgBox "Nenhuma Tag InfRps foi encontrada no XML.')"
            'Exit Function
        End If
        '//Adiciona Tag de Assinatura para CADA RPS
        Call doc.SetProperty("SelectionNamespaces", DS)
        Set oSigs = doc.selectNodes(".//ds:Signature")
        
      '  Set oSigs = doc.selectNodes(".//envEvento")
        

        If oSigs.Length > 0 Then
            For i = (oSigs.Length - 1) To 0 Step -1
                '  begin
                Set oNode = oSigs.Item(i)
                Call AddSign(doc, oNode, Cert, StoreDst)
            Next
            'End If
            '//Assina Tag de Assinatura do LoteRps
            '//oSigs = Doc.selectNodes('.//ds:Signature')
            '//if oSigs.length > 0 then
            '//  AddSign(Doc, oSigs.item[pred(oSigs.length)], Cert, Store)
        End If

        '//if (pKeyOut = nil) then
        '//  Exception.Create('Falha durante assinatura.\n')
        'except
        ' Sett = Null
        ' StoreDst = Null
        ' StoreSrc = Null
        ' Raise
        'End
        '//Fechar Store
        'Sett = Null
        'StoreDst = Null
        'StoreSrc = Null
        'Result = True
    End If
   
    SignXML = True
    Dim txt_out As String

    txt_out = doc.xml
    Do While InStr(txt_out, "  ") > 0
        txt_out = Replace(txt_out, "  ", "|")
    Loop
    txt_out = Replace(txt_out, "|", "")
    txt_out = Replace(txt_out, vbCr, "")
    txt_out = Replace(txt_out, vbLf, "")
    If doc.LoadXml(txt_out) Then
    End If
End Function








