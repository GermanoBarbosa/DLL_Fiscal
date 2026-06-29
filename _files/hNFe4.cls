VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public retNfeStatusServico  As New hNFe4_StatusServico
Public retNFe_EnviaLote2    As New hNFe4_EnviaLote2
Public retNfeInutilizacao   As New hNFe4_Inutilizacao
Public retNfeCartaCorrecao  As New hNFe4_CartaCorrecao
Public retNFe_ConsSitNFe    As New hNFe4_ConsSitNFe
Public retNFe_DistDFe       As New hNFe4_DistDFe
Public retNFe_DownloadNFe   As New hNFe4_DownloadNFe
Public retNFe_ConsReciNFe   As New hNFe4_ConsReciNFe


Dim WithEvents e_xmldom As WinHttp.WinHttpRequest
Attribute e_xmldom.VB_VarHelpID = -1
'Dim WithEvents a as hFireControl

Dim pURL_NfeStatusServico As String, pURL_NfeStatusServico_xsd As String
Dim pURL_NFeAutorizacao_nfe As String, pURL_NFeAutorizacao_xsd As String
Dim pURL_NFeAutorizacao_nfce As String
Dim pURL_NFeRecepcaoEvento4_xsd As String
Dim pURL_NFeInutilizacao As String
Dim pURL_NFeInutilizacao_xsd As String
Dim pURL_NFeRecepcaoEvento4 As String
Dim pURL_NFeConsultaProtocolo As String
Dim pURL_NFeConsultaProtocolo_xsd As String
Dim pURL_NFeDistribuicaoDFe As String
Dim pURL_NFeRetAutorizacao4 As String


Private mcert As String
Public NFe  As New hNFe4_Doc

Public m_Ambiente As Long
Public m_UF As String
Public m_UF_Cod As String

Public m_LastError As String


Sub SetCert(ByVal mCertificadoSerial As String)
  '  sCert: array [Boolean] of string = ('s',  '');
'  sDisp: array [Boolean] of string = ('is', 'l');
Dim txt As String
Dim m_pos As Long

    Dim i As Long  ',: Cardinal;
    Dim oNode As IXMLDOMNode
    Dim SetT As New Settings, Certs, StoreSrc As New Store, StoreDst As New Store, Cert As Certificate   ': OleVariant;
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
    '//Somente certificados com data v?lida.
    If Certs.Count > 0 Then
        Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID)
    End If
    'MsgBox Certs.item(1).SubjectName
    'Certs.Select
    
    
    If Certs.Count = 0 Then
        'MsgBox "Sem certificados"
    Else
        Set Cert = Certs.Item(1)
        For i = Certs.Count To 1 Step -1
            If Certs.Item(i).SerialNumber = mCertificadoSerial Then
                Set Cert = Certs.Item(i)
                
                
                'mcert = "LOCAL_MACHINE\My\" & Certs.Item(i).SubjectName
                StoreDst.Certificates.Clear
                StoreDst.Add Cert 'Certs.item(1)
                'mcert = "CURRENT_USER\\Personal\\" & Certs.Item(i).SubjectName
                
                txt = Certs.Item(i).SubjectName
                m_pos = InStr(txt, "CN=")
                If m_pos > 0 Then
                    txt = Mid$(txt, m_pos + 3)
                    m_pos = InStr(txt, ",")
                    If m_pos > 0 Then
                       txt = Left$(txt, m_pos - 1)
                       mcert = "CURRENT_USER\MY\" & txt
                    End If
                End If
            
            End If
            'Debug.Print Certs.Item(i).PrivateKey
            'Debug.Print Certs.Item(i).PublicKey.
            'Debug.Print Certs.Item(i).SubjectName
            'Debug.Print Certs.Item(i).HasPrivateKey

        Next
        
    End If

    
End Sub

Function Inicia(Optional ByVal n_ambiente As String) As Boolean
    Inicia = True

    ' ---------------------------------------------------------------
    ' Roteamento NF-e 4.00 por UF / ambiente / conting?ncia
    ' ---------------------------------------------------------------
    ' m_Ambiente = 1 (Produ??o) | 2 (Homologa??o)
    ' n_ambiente = "SVAN" | "SVC-RS" | "SVC-AN"  (conting?ncia)
    ' Se n_ambiente vazio, roteia por m_UF_Cod (cUF IBGE do emitente)
    '
    ' NFC-e (modelo 65) SEMPRE usa SVRS (independente do estado).
    ' ---------------------------------------------------------------

    If m_Ambiente = 1 Then
        ' ====================== PRODU??O ======================
        Select Case n_ambiente
            Case "SVAN", "SVC-AN"
                pURL_NfeStatusServico = "https://www.sefazvirtual.fazenda.gov.br/NFeStatusServico4/NFeStatusServico4.asmx"
                pURL_NFeAutorizacao_nfe = "https://www.sefazvirtual.fazenda.gov.br/NFeAutorizacao4/NFeAutorizacao4.asmx"
                pURL_NFeRetAutorizacao4 = "https://www.sefazvirtual.fazenda.gov.br/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"
                pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                pURL_NFeInutilizacao = "https://www.sefazvirtual.fazenda.gov.br/NFeInutilizacao4/NFeInutilizacao4.asmx"
                pURL_NFeRecepcaoEvento4 = "https://www.sefazvirtual.fazenda.gov.br/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                pURL_NFeConsultaProtocolo = "https://www.sefazvirtual.fazenda.gov.br/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx"

            Case "SVC-RS"
                pURL_NfeStatusServico = "https://nfe.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                pURL_NFeAutorizacao_nfe = "https://nfe.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                pURL_NFeRetAutorizacao4 = "https://nfe.svrs.rs.gov.br/ws/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"
                pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                pURL_NFeInutilizacao = "https://nfe.svrs.rs.gov.br/ws/NfeInutilizacao/NFeInutilizacao4.asmx"
                pURL_NFeRecepcaoEvento4 = "https://nfe.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                pURL_NFeConsultaProtocolo = "https://nfe.svrs.rs.gov.br/ws/NfeConsultaProtocolo/NFeConsultaProtocolo4.asmx"

            Case Else
                ' Roteamento por UF (estados com gateway pr?prio)
                Select Case m_UF_Cod
                    Case "13"   ' AM  Amazonas
                        pURL_NfeStatusServico = "https://nfe.sefaz.am.gov.br/services2/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfe.sefaz.am.gov.br/services2/services/NfeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfe.sefaz.am.gov.br/services2/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.sefaz.am.gov.br/services2/services/NfeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.sefaz.am.gov.br/services2/services/NfeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfe.sefaz.am.gov.br/services2/services/NfeConsultaProtocolo4"

                    Case "29"   ' BA  Bahia
                        pURL_NfeStatusServico = "https://nfe.sefaz.ba.gov.br/webservices/NFeStatusServico4/NFeStatusServico4.asmx"
                        pURL_NFeAutorizacao_nfe = "https://nfe.sefaz.ba.gov.br/webservices/NFeAutorizacao4/NFeAutorizacao4.asmx"
                        pURL_NFeRetAutorizacao4 = "https://nfe.sefaz.ba.gov.br/webservices/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.sefaz.ba.gov.br/webservices/NFeInutilizacao4/NFeInutilizacao4.asmx"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.sefaz.ba.gov.br/webservices/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                        pURL_NFeConsultaProtocolo = "https://nfe.sefaz.ba.gov.br/webservices/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx"

                    Case "23"   ' CE  Cear?
                        pURL_NfeStatusServico = "https://nfe.sefaz.ce.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfe.sefaz.ce.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfe.sefaz.ce.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.sefaz.ce.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.sefaz.ce.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfe.sefaz.ce.gov.br/nfe/services/NFeConsultaProtocolo4"

                    Case "52"   ' GO  Goi?s
                        pURL_NfeStatusServico = "https://nfe.sefaz.go.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfe.sefaz.go.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfe.sefaz.go.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.sefaz.go.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.sefaz.go.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfe.sefaz.go.gov.br/nfe/services/NFeConsultaProtocolo4"

                    Case "31"   ' MG  Minas Gerais
                        pURL_NfeStatusServico = "https://nfe.fazenda.mg.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfe.fazenda.mg.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfe.fazenda.mg.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.fazenda.mg.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.fazenda.mg.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfe.fazenda.mg.gov.br/nfe/services/NfeConsultaProtocolo4"

                    Case "50"   ' MS  Mato Grosso do Sul
                        pURL_NfeStatusServico = "https://nfe.sefaz.ms.gov.br/ws/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfe.sefaz.ms.gov.br/ws/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfe.sefaz.ms.gov.br/ws/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.sefaz.ms.gov.br/ws/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.sefaz.ms.gov.br/ws/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfe.sefaz.ms.gov.br/ws/NfeConsultaProtocolo4"

                    Case "51"   ' MT  Mato Grosso
                        pURL_NfeStatusServico = "https://nfe.sefaz.mt.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfe.sefaz.mt.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfe.sefaz.mt.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.sefaz.mt.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.sefaz.mt.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfe.sefaz.mt.gov.br/nfe/services/NfeConsultaProtocolo4"

                    Case "26"   ' PE  Pernambuco
                        pURL_NfeStatusServico = "https://nfe.sefaz.pe.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfe.sefaz.pe.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfe.sefaz.pe.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.sefaz.pe.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.sefaz.pe.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfe.sefaz.pe.gov.br/nfe/services/NfeConsultaProtocolo4"

                    Case "41"   ' PR  Paran?
                        pURL_NfeStatusServico = "https://nfe.sefaz.pr.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfe.sefaz.pr.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfe.sefaz.pr.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.sefaz.pr.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.sefaz.pr.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfe.sefaz.pr.gov.br/nfe/services/NfeConsultaProtocolo4"

                    Case "33"   ' RJ  Rio de Janeiro
                        pURL_NfeStatusServico = "https://nfe.sefaz.rj.gov.br/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfe.sefaz.rj.gov.br/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfe.sefaz.rj.gov.br/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.sefaz.rj.gov.br/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.sefaz.rj.gov.br/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfe.sefaz.rj.gov.br/services/NfeConsultaProtocolo4"

                    Case "35"   ' SP  S?o Paulo
                        pURL_NfeStatusServico = "https://nfe.fazenda.sp.gov.br/ws/nfestatusservico4.asmx"
                        pURL_NFeAutorizacao_nfe = "https://nfe.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx"
                        pURL_NFeRetAutorizacao4 = "https://nfe.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx"
                        pURL_NFeConsultaProtocolo = "https://nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx"

                    Case Else   ' Default  SVRS (AC, AL, AP, DF, ES, MA, PA, PB, PI, RN, RO, RR, RS, SC, SE, TO)
                        pURL_NfeStatusServico = "https://nfe.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        pURL_NFeAutorizacao_nfe = "https://nfe.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeRetAutorizacao4 = "https://nfe.svrs.rs.gov.br/ws/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe.svrs.rs.gov.br/ws/NfeInutilizacao/NFeInutilizacao4.asmx"
                        pURL_NFeRecepcaoEvento4 = "https://nfe.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                        pURL_NFeConsultaProtocolo = "https://nfe.svrs.rs.gov.br/ws/NfeConsultaProtocolo/NFeConsultaProtocolo4.asmx"

                End Select
        End Select

    Else
        ' ====================== HOMOLOGA��O ======================
        Select Case n_ambiente
            Case "SVAN", "SVC-AN"
                ' SVAN/SVC-AN em homologa??o  SEFAZ nacional (hom1)
                pURL_NfeStatusServico = "https://hom1.nfe.fazenda.gov.br/NFeStatusServico4/NFeStatusServico4.asmx"
                pURL_NFeAutorizacao_nfe = "https://hom1.nfe.fazenda.gov.br/NFeAutorizacao4/NFeAutorizacao4.asmx"
                pURL_NFeRetAutorizacao4 = "https://hom1.nfe.fazenda.gov.br/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"
                pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                pURL_NFeInutilizacao = "https://hom1.nfe.fazenda.gov.br/NFeInutilizacao4/NFeInutilizacao4.asmx"
                pURL_NFeRecepcaoEvento4 = "https://hom1.nfe.fazenda.gov.br/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                pURL_NFeConsultaProtocolo = "https://hom1.nfe.fazenda.gov.br/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx"

            Case "SVC-RS"
                pURL_NfeStatusServico = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                pURL_NFeAutorizacao_nfe = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                pURL_NFeRetAutorizacao4 = "https://nfe-homologacao.svrs.rs.gov.br/ws/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"
                pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                pURL_NFeInutilizacao = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeInutilizacao/NFeInutilizacao4.asmx"
                pURL_NFeRecepcaoEvento4 = "https://nfe-homologacao.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                pURL_NFeConsultaProtocolo = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeConsultaProtocolo/NFeConsultaProtocolo4.asmx"

            Case Else
                ' Roteamento por UF em homologa??o
                Select Case m_UF_Cod
                    Case "13"   ' AM
                        pURL_NfeStatusServico = "https://homnfe.sefaz.am.gov.br/services2/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://homnfe.sefaz.am.gov.br/services2/services/NfeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://homnfe.sefaz.am.gov.br/services2/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://homnfe.sefaz.am.gov.br/services2/services/NfeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://homnfe.sefaz.am.gov.br/services2/services/NfeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://homnfe.sefaz.am.gov.br/services2/services/NfeConsultaProtocolo4"

                    Case "29"   ' BA
                        pURL_NfeStatusServico = "https://hnfe.sefaz.ba.gov.br/webservices/NFeStatusServico4/NFeStatusServico4.asmx"
                        pURL_NFeAutorizacao_nfe = "https://hnfe.sefaz.ba.gov.br/webservices/NFeAutorizacao4/NFeAutorizacao4.asmx"
                        pURL_NFeRetAutorizacao4 = "https://hnfe.sefaz.ba.gov.br/webservices/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://hnfe.sefaz.ba.gov.br/webservices/NFeInutilizacao4/NFeInutilizacao4.asmx"
                        pURL_NFeRecepcaoEvento4 = "https://hnfe.sefaz.ba.gov.br/webservices/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                        pURL_NFeConsultaProtocolo = "https://hnfe.sefaz.ba.gov.br/webservices/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx"

                    Case "23"   ' CE
                        pURL_NfeStatusServico = "https://nfeh.sefaz.ce.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfeh.sefaz.ce.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfeh.sefaz.ce.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfeh.sefaz.ce.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfeh.sefaz.ce.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfeh.sefaz.ce.gov.br/nfe/services/NfeConsultaProtocolo4"

                    Case "52"   ' GO
                        pURL_NfeStatusServico = "https://homolog.sefaz.go.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://homolog.sefaz.go.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://homolog.sefaz.go.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://homolog.sefaz.go.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://homolog.sefaz.go.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://homolog.sefaz.go.gov.br/nfe/services/NFeConsultaProtocolo4"

                    Case "31"   ' MG
                        pURL_NfeStatusServico = "https://hnfe.fazenda.mg.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://hnfe.fazenda.mg.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://hnfe.fazenda.mg.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://hnfe.fazenda.mg.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://hnfe.fazenda.mg.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://hnfe.fazenda.mg.gov.br/nfe/services/NfeConsultaProtocolo4"

                    Case "50"   ' MS
                        pURL_NfeStatusServico = "https://homnfe.sefaz.ms.gov.br/ws/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://homnfe.sefaz.ms.gov.br/ws/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://homnfe.sefaz.ms.gov.br/ws/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://homnfe.sefaz.ms.gov.br/ws/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://homnfe.sefaz.ms.gov.br/ws/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://homnfe.sefaz.ms.gov.br/ws/NfeConsultaProtocolo4"

                    Case "51"   ' MT
                        pURL_NfeStatusServico = "https://homnfe.sefaz.mt.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://homnfe.sefaz.mt.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://homnfe.sefaz.mt.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://homnfe.sefaz.mt.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://homnfe.sefaz.mt.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://homnfe.sefaz.mt.gov.br/nfe/services/NfeConsultaProtocolo4"

                    Case "26"   ' PE
                        pURL_NfeStatusServico = "https://nfehomolog.sefaz.pe.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://nfehomolog.sefaz.pe.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://nfehomolog.sefaz.pe.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfehomolog.sefaz.pe.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://nfehomolog.sefaz.pe.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://nfehomolog.sefaz.pe.gov.br/nfe/services/NfeConsultaProtocolo4"

                    Case "41"   ' PR
                        pURL_NfeStatusServico = "https://homnfe.sefaz.pr.gov.br/nfe/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://homnfe.sefaz.pr.gov.br/nfe/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://homnfe.sefaz.pr.gov.br/nfe/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://homnfe.sefaz.pr.gov.br/nfe/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://homnfe.sefaz.pr.gov.br/nfe/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://homnfe.sefaz.pr.gov.br/nfe/services/NfeConsultaProtocolo4"

                    Case "33"   ' RJ
                        pURL_NfeStatusServico = "https://homologacao.nfe.sefaz.rj.gov.br/services/NfeStatusServico4"
                        pURL_NFeAutorizacao_nfe = "https://homologacao.nfe.sefaz.rj.gov.br/services/NFeAutorizacao4"
                        pURL_NFeRetAutorizacao4 = "https://homologacao.nfe.sefaz.rj.gov.br/services/NFeRetAutorizacao4"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://homologacao.nfe.sefaz.rj.gov.br/services/NFeInutilizacao4"
                        pURL_NFeRecepcaoEvento4 = "https://homologacao.nfe.sefaz.rj.gov.br/services/NFeRecepcaoEvento4"
                        pURL_NFeConsultaProtocolo = "https://homologacao.nfe.sefaz.rj.gov.br/services/NfeConsultaProtocolo4"

                    Case "35"   ' SP
                        pURL_NfeStatusServico = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfestatusservico4.asmx"
                        pURL_NFeAutorizacao_nfe = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx"
                        pURL_NFeRetAutorizacao4 = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx"
                        pURL_NFeRecepcaoEvento4 = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx"
                        pURL_NFeConsultaProtocolo = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx"

                    Case Else   ' Default  SVRS homologa??o
                        pURL_NfeStatusServico = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        pURL_NFeAutorizacao_nfe = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        pURL_NFeRetAutorizacao4 = "https://nfe-homologacao.svrs.rs.gov.br/ws/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"
                        pURL_NFeInutilizacao = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeInutilizacao/NFeInutilizacao4.asmx"
                        pURL_NFeRecepcaoEvento4 = "https://nfe-homologacao.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                        pURL_NFeConsultaProtocolo = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeConsultaProtocolo/NFeConsultaProtocolo4.asmx"

                End Select
        End Select
    End If

    ' NFeDistribuicaoDFe - servi?o nacional (sempre no ambiente nacional, independente da UF)
    If m_Ambiente = 1 Then
        pURL_NFeDistribuicaoDFe = "https://www.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"
    Else
        pURL_NFeDistribuicaoDFe = "https://hom1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"
    End If

End Function

Public Sub Initialize(ByVal m_path_Schema As String, ByVal m_path_app As String)
Dim m_Caminho As String

    m_Caminho = m_path_Schema
    m_path = m_path_app
    
    pURL_NFeAutorizacao_xsd = m_Caminho & "enviNFe_v4.00.xsd"
    pURL_NfeStatusServico_xsd = m_Caminho & "consStatServ_v4.00.xsd"
    
    pURL_NFeRecepcaoEvento4_xsd = m_Caminho & "consReciNFe_v4.00.xsd"
    pURL_NFeInutilizacao_xsd = m_Caminho & "inutNFe_v4.00.xsd"
    pURL_NFeConsultaProtocolo_xsd = m_Caminho & "consSitNFe_v4.00.xsd"
    
    If Not FileExists(pURL_NFeAutorizacao_xsd) Then
        MsgBox "Schema 4.00 N?O encontrado em: " & m_Caminho
    End If

End Sub

Function NfeStatusServico() As Boolean
          Dim xmldoc As New WinHttpRequest
          Dim txt As String
          Dim xmlnf As New MSXML2.DOMDocument50
          Dim xmldom_resp As New MSXML2.DOMDocument50
          Dim txt_xml As String

10
230       txt_xml = ""
240       txt_xml = txt_xml & "<?xml version=""1.0"" encoding=""utf-8""?>"
250       txt_xml = txt_xml & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
260       txt_xml = txt_xml & "<soap12:Body>"
270       txt_xml = txt_xml & "<nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeStatusServico4""> "
          
280       txt_xml = txt_xml & "<consStatServ versao=""4.00"" xmlns=""http://www.portalfiscal.inf.br/nfe""><tpAmb>"
290       txt_xml = txt_xml & m_Ambiente
300       txt_xml = txt_xml & "</tpAmb><cUF>"
310       txt_xml = txt_xml & m_UF_Cod
320       txt_xml = txt_xml & "</cUF><xServ>STATUS</xServ></consStatServ>"
          
          
330       txt_xml = txt_xml & "</nfeDadosMsg>"
340       txt_xml = txt_xml & "</soap12:Body>"
350       txt_xml = txt_xml & "</soap12:Envelope>"
'360       End If
         ' Debug.Print txt_xml
370       If xmlnf.LoadXml(txt_xml) Then
380           txt = ValidaSchema(xmlnf, "http://www.portalfiscal.inf.br/nfe", pURL_NfeStatusServico_xsd, False)
390           If txt = "Validado" Then
                  'Call xmldoc.Open("POST", "https://www.example.com/", False)
                  'xmldoc.Send ("TESTE")
                  'If LenB(xmldoc.ResponseText) Then
                  '  Debug.Print xmldoc.ResponseText
                  'End If
400               Call xmldoc.Open("POST", pURL_NfeStatusServico & "?wsdl", False)
410               'Call xmldoc.SetRequestHeader("content-type", "text/xml")

                    'The server cannot service the request because the media type is unsupported.
                    'Sem esse parametro
                  Call xmldoc.SetRequestHeader("content-type", "application/soap+xml")
420               'Call xmldoc.SetRequestHeader("Cache-Control", "no-cache")
                  
430               'Call xmldoc.SetRequestHeader("soapAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeStatusServico4/nfeStatusServicoNF")
                  'Call xmldoc.setRequestHeader("soapAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeAutorizacao4/nfeAutorizacaoLote")
                  
                 ' Call xmldoc.setRequestHeader("SOAPAction", """")
                'Debug.Print xmldoc.Status
440              xmldoc.SetClientCertificate (mcert)
                'Debug.Print xmldoc.Status
                  xmldoc.Option(WinHttpRequestOption_SelectCertificate) = (mcert)
                  
                  'https://www.experts-exchange.com/questions/28821379/winhttp-dll-not-working-with-TLS1-1-and-TLS1-2-on-windows-2008-r2-x64.html
                  'xmldoc.Option(WinHttpRequestOption_SecureProtocols) = 2048      'TLS 1.2
                '  xmldoc.Option(WinHttpRequestOption_SecureProtocols) = 512       'TLS 1.1
'                  xmldoc.Option(WinHttpRequestOption_SecureProtocols) = 128 Or &H20         'TLS 1.0
'                  xmldoc.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
'                  xmldoc.Option(WinHttpRequestOption_SelectCertificate) = mcert

                  'xmldoc.Option(6) = False
450               xmldoc.Send (xmlnf.xml & "")
460               If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
                      'Debug.Print "---------------------"
                      'Debug.Print pURL_NfeStatusServico
                      'Debug.Print xmldom_resp.xml
                      'Debug.Print "---------------------"

                      Dim ooo As IXMLDOMNode
                      Dim ooo2 As IXMLDOMNode
470                   If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
480                       If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
490                           If ExistSelect(ooo.childNodes, ooo, "retConsStatServ") Then
500                               If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
510                                   retNfeStatusServico.cStat = ooo2.Text
520                               End If
530                               If ExistSelect(ooo.childNodes, ooo2, "tpAmb") Then
540                                   retNfeStatusServico.tpAmb = ooo2.Text
550                               End If
560                               If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
570                                   retNfeStatusServico.xMotivo = ooo2.Text
580                               End If
590                           End If
600                       End If
610                   End If
                  Else
                    MsgBox xmldom_resp.parseError.srcText, , "parseError - 2"
                    
620               End If
630               NfeStatusServico = retNfeStatusServico.cStat = 107
640           Else
650               m_LastError = txt
660           End If
670       Else
680           MsgBox xmlnf.parseError.srcText, , "parseError"
690           SetFileBytes App.Path & "\error_1.txt", txt_xml
700           SetFileBytes App.Path & "\error_2.txt", xmlnf.Text
710       End If
720       Set xmldoc = Nothing
730       On Error GoTo 0
740       Exit Function
750       Resume Next
NfeStatusServico_Error:
760       Debug.Print Err.Number, Err.Description
770       If Err.Number = -2147012711 Then
780           MsgBox "Erro de certificado"
790       ElseIf Err.Number = -2147012852 Then
800           MsgBox "? necess?rio um certificado para concluir a autentica??o do cliente"
810       ElseIf Err.Number = -2147012851 Then
820           MsgBox "A autoridade de certifica??o n?o ? v?lida ou est? incorreta"
              'https://www.projetoacbr.com.br/forum/topic/27273-a-autoridade-de-certifica%C3%A7%C3%A3o-n%C3%A3o-%C3%A9-valida-ou-esta-incorreta/
830       Else
840           MsgBox "Error linha: " & Erl & "," & Err.Number & " (" & Err.Description & ") in procedure NfeStatusServico of M?dulo mMainTeste"
850           Debug.Print Err.Number,
860       End If
End Function

Function EnviaLote2(ByVal xml As String, xml_aut As String, ByVal mCertificadoSel As String) As Boolean
   On Error GoTo EnviaLote2_Error

10
          Dim xmlnf As New MSXML2.DOMDocument50
          Dim xmldoc As New WinHttp.WinHttpRequest    'MSXML2.XMLHTTP50
          Dim xml_ass As String
          Dim txt_xml As New hStringBuilder
          Dim txt As String
          Dim xmldom_resp As New MSXML2.DOMDocument50
          Dim txt_out As String
                   
20        txt_xml.Clear
          
30        txt_xml.Append "<?xml version=""1.0"" encoding=""utf-8""?>"
40        txt_xml.Append "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
50        txt_xml.Append "<soap12:Body>"
60        txt_xml.Append "<nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeAutorizacao4"">"
          
70        txt_xml.Append "<enviNFe xmlns=""http://www.portalfiscal.inf.br/nfe"" versao=""4.00"">"
          
80        txt_xml.Append "<idLote>1</idLote>"
90        txt_xml.Append "<indSinc>1</indSinc>"
100       txt_xml.Append Mid(xml, 1, Len(xml) - 2)
110       txt_xml.Append "</enviNFe>"
          
120       txt_xml.Append "</nfeDadosMsg>"
130       txt_xml.Append "</soap12:Body>"
140       txt_xml.Append "</soap12:Envelope>"
          
          Dim ooo_insert As IXMLDOMNode
          
          Dim m_File As String
200           m_File = m_path & "\NFe\temp\" & Format(Time, "hh-mm-ss") & ".xml"
220       Call SetFileBytes(m_File, txt_xml.toString)
          
          
          
          'leiauteNFe_v3 0.1
          'xmlnf.loadXML (xml)
          'txt = ValidaSchema(xmlnf, "http://www.portalfiscal.inf.br/nfe", App.Path & "\NFe\Schema-4.00\leiauteNFe_v4.00.xsd", False)
          'If txt = "Validado" Then
          'Else
          '    Debug.Print txt
          'End If
          
230       If xmlnf.LoadXml(txt_xml) Then

240           txt = ValidaSchema(xmlnf, "http://www.portalfiscal.inf.br/nfe", pURL_NFeAutorizacao_xsd, False)
250           If txt = "Validado" Then
260               If InStr(xml, "<mod>55</mod>") > 0 Then
270                   Call xmldoc.Open("POST", pURL_NFeAutorizacao_nfe & "?wsdl", False)
280               ElseIf InStr(xml, "<mod>65</mod>") > 0 Then
290                   Call xmldoc.Open("POST", pURL_NFeAutorizacao_nfce & "?wsdl", False)
300               End If
  
                  
310               Debug.Print pURL_NFeAutorizacao_nfce
                  
                  
320               Call xmldoc.SetRequestHeader("content-type", "application/soap+xml; charset=utf-8")
                                    

330               Call xmldoc.SetRequestHeader("Cache-Control", "no-cache")
340               Call xmldoc.SetRequestHeader("soapAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeAutorizacao4/nfeAutorizacaoLote")
350               xmldoc.SetClientCertificate (mcert)


370               xmldoc.Send (xmlnf.xml)
380               If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
                    
390                   Debug.Print xmldom_resp.Text
400                   Debug.Print xmldom_resp.xml
            
                      Dim ooo As IXMLDOMNode
                      Dim ooo2 As IXMLDOMNode
                      Dim oooTemp As IXMLDOMNode
                      Dim xmldom_resp2 As New MSXML2.DOMDocument50
410                   If Not xmldom_resp.documentElement Is Nothing Then

450                           xmldom_resp.Save m_path & "\NFe\Resp\" & NFe.m_Chave_NF & "-aut.xml"

470                       If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
480                           If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
490                               If ExistSelect(ooo.childNodes, ooo, "retEnviNFe") Then
500                                   If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
510                                       retNFe_EnviaLote2.xMotivo = ooo2.Text
520                                       Debug.Print ooo.Text
                                          
530                                   End If
540                                   If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
550                                       retNFe_EnviaLote2.cMotivo = ooo2.Text
560                                       Debug.Print ooo2.Text
570                                   End If
580                                   If ExistSelect(ooo.childNodes, ooo2, "protNFe") Then
590                                       xml_aut = ooo2.xml
600                                       If xmldom_resp2.LoadXml(ooo2.xml) Then
610                                           If ExistSelect(xmldom_resp2.childNodes, ooo2, "protNFe") Then
                                                '  MsgBox ooo2.xml
620                                               If ExistSelect(ooo2.childNodes, ooo2, "infProt") Then
630                                                   If ExistSelect(ooo2.childNodes, oooTemp, "xMotivo") Then
640                                                       retNFe_EnviaLote2.protNFe.xMotivo = oooTemp.Text
                                                   
650                                                   End If
660                                                   If ExistSelect(ooo2.childNodes, oooTemp, "cStat") Then
670                                                       retNFe_EnviaLote2.protNFe.cStat = oooTemp.Text
680                                                   End If
690                                               End If
700                                           End If
710                                       End If
                                          
                                          
                                         
                                          
720                                   End If
730                               End If
740                           End If
750                       End If
760                   End If
                Else
                    Debug.Print xmldoc.ResponseText
770               End If
780               EnviaLote2 = True
790           Else
800               Debug.Print txt
810               m_LastError = txt
820           End If
830       Else
840           MsgBox "N?o carregou o xml"
850       End If
          'Set xmldom = Nothing
860       Set xmldoc = Nothing


   On Error GoTo 0
   Exit Function

EnviaLote2_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EnviaLote2 of M?dulo de classe hNFe4 " & Erl

End Function

Function GetNota(m_xml As String, m_Chave As String, m_NF_NUM_NOTA As Long) As Boolean
    m_xml = NFe.GetNota(m_Chave, m_NF_NUM_NOTA)
    GetNota = True
End Function

Private Function ValidaSchema(ByVal docXML As DOMDocument50, ByVal strUrn As String, ByVal strXsdArquivo As String, ByVal comAssinatura As Boolean) As String
' cria um cache schema e adiciona o arquivo strXsdArquivo.
    Dim Xs As New XMLSchemaCache50
    Dim lngTam As Long, lngTam2 As Long, strCampoErrado As String
    ValidaSchema = True
    Xs.Add strUrn, strXsdArquivo

    ' cria um XML DOMDocument object.
    Dim xd As New DOMDocument50

    ' adiciona o schemaCache ao documento.
    Set xd.schemas = Xs

    ' Carrega o documento docXML.
    xd.async = False
    xd.LoadXml docXML.xml

    ' Return validation results in message to the user.
    If xd.parseError.errorCode <> 0 Then

        If Not comAssinatura Then
            'Ignorar erro de assinatura
            If InStr(1, UCase(xd.parseError.reason), "SIGNATURE") > 0 Then
                ValidaSchema = "Validado"
                Exit Function
            End If
        End If

        ValidaSchema = "Erro na valida??o: " & vbCrLf
        ValidaSchema = ValidaSchema & "- C?digo    : " & xd.parseError.errorCode & vbCrLf
        lngTam = InStr(1, xd.parseError.reason, "enumeration")
        If lngTam > 0 Then
            lngTam = InStr(1, xd.parseError.reason, "}")
            lngTam2 = InStr(lngTam + 1, xd.parseError.reason, "'")
            lngTam2 = lngTam2 - lngTam - 1
            strCampoErrado = Mid(xd.parseError.reason, lngTam + 1, lngTam2)
            ValidaSchema = ValidaSchema & "- Descri��o: O Campo " & Chr(34) & strCampoErrado & Chr(34) & " � Inv�lido." & vbCrLf & "Por Favor, Verifique os Dados de sua Nota Fiscal." & vbCrLf

            ValidaSchema = ValidaSchema & "- Linha    : " & xd.parseError.Line & vbCrLf
            ValidaSchema = ValidaSchema & "- Posi??o  : " & xd.parseError.linepos
        Else
            ValidaSchema = ValidaSchema & "- Descri??o: " & xd.parseError.reason & vbCrLf
            ValidaSchema = ValidaSchema & "- Fonte    : " & Chr(10) & Replace(xd.parseError.srcText, "><", "><" & vbCrLf) & vbCrLf
            ValidaSchema = ValidaSchema & "- Linha    : " & xd.parseError.Line
        End If
        'FunValidaSchema = FunValidaSchema & "- Fonte    : " & xd.parseError.srcText & vbCrLf
        'FunValidaSchema = FunValidaSchema & "- Linha    : " & xd.parseError.Line
        '      FunValidaSchema = False


    Else
        ValidaSchema = "Validado"
    End If

End Function

Private Function ExistSelect(childNodes As IXMLDOMNodeList, ooo As IXMLDOMNode, ByVal mName As String) As Boolean
Dim cc As Long
    For cc = 0 To childNodes.Length
        If Not childNodes(cc) Is Nothing Then
            If childNodes(cc).baseName = mName Then
                Set ooo = childNodes(cc)
                ExistSelect = True
                Exit For
            End If
        End If
    Next
End Function

Function GetCert(ByVal mCertificadoSerial As String) As Certificate
  '  sCert: array [Boolean] of string = ('s',  '');
'  sDisp: array [Boolean] of string = ('is', 'l');
    Dim i As Long  ',: Cardinal;
    Dim oNode As IXMLDOMNode
    Dim SetT As New Settings, Certs, StoreSrc As New Store, StoreDst As New Store, Cert As Certificate   ': OleVariant;
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
    '//Somente certificados com data v?lida.
    If Certs.Count > 0 Then
        Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID)
    End If
    'MsgBox Certs.item(1).SubjectName
    'Certs.Select
    
    
    If Certs.Count = 0 Then
        'MsgBox "Sem certificados"
    Else
        Set Cert = Certs.Item(1)
        For i = Certs.Count To 1 Step -1
            If Certs.Item(i).SerialNumber = mCertificadoSerial Then
                Set GetCert = Certs.Item(i)
                
                Exit For
                'mcert = "LOCAL_MACHINE\My\" & Certs.Item(i).SubjectName
                StoreDst.Certificates.Clear
                StoreDst.Add Cert 'Certs.item(1)
                mcert = "CURRENT_USER\\TMP2\\" & Certs.Item(i).SubjectName
            End If
            Debug.Print Certs.Item(i).PrivateKey
            'Debug.Print Certs.Item(i).PublicKey.
            Debug.Print Certs.Item(i).SubjectName
            Debug.Print Certs.Item(i).HasPrivateKey
             
        Next
    End If

End Function



Function NfeCancela(mCertificadoSel, ByVal n_cOrgao As String, ByVal n_chNFe As String, ByVal m_CNPJ As String, ByVal m_xJust As String, ByVal nProt As String, m_xml_send As String, m_xml_resp As String) As Boolean
    Dim xmldoc As New WinHttpRequest
    Dim txt As String
    Dim xmlnf As New MSXML2.DOMDocument50
    Dim xmldom_resp As New MSXML2.DOMDocument50
    Dim txt_xml As String



    txt_xml = ""
    txt_xml = txt_xml & "<?xml version=""1.0"" encoding=""utf-8""?>"

    txt_xml = txt_xml & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
    txt_xml = txt_xml & "  <soap12:Body>"
    txt_xml = txt_xml & "    <nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeRecepcaoEvento4"">"
    
    txt_xml = txt_xml & get_msg(mCertificadoSel, n_cOrgao, n_chNFe, m_CNPJ, m_xJust, nProt)
    'txt_xml = txt_xml & "teste"

    txt_xml = txt_xml & "</nfeDadosMsg>"
    txt_xml = txt_xml & "  </soap12:Body>"
    txt_xml = txt_xml & "</soap12:Envelope>"

 
'360       End If
   ' Debug.Print txt_xml
    If xmlnf.LoadXml(txt_xml) Then
        pURL_NFeRecepcaoEvento4_xsd = m_path & "envEvento_v1.00.xsd"
        txt = ValidaSchema(xmlnf, "http://www.portalfiscal.inf.br/nfe", pURL_NFeRecepcaoEvento4_xsd, False)
        If txt = "Validado" Then
         
            Call xmldoc.Open("POST", pURL_NFeRecepcaoEvento4 & "?wsdl", False)

              'The server cannot service the request because the media type is unsupported.
              'Sem esse parametro
            Call xmldoc.SetRequestHeader("content-type", "application/soap+xml")
            'Call xmldoc.SetRequestHeader("Cache-Control", "no-cache")
            
            'Call xmldoc.SetRequestHeader("soapAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeStatusServico4/nfeStatusServicoNF")
            'Call xmldoc.setRequestHeader("soapAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeAutorizacao4/nfeAutorizacaoLote")
            
           ' Call xmldoc.setRequestHeader("SOAPAction", """")
          'Debug.Print xmldoc.Status
           xmldoc.SetClientCertificate (mcert)
          'Debug.Print xmldoc.Status
            xmldoc.Option(WinHttpRequestOption_SelectCertificate) = (mcert)
            
            'https://www.experts-exchange.com/questions/28821379/winhttp-dll-not-working-with-TLS1-1-and-TLS1-2-on-windows-2008-r2-x64.html
            'xmldoc.Option(WinHttpRequestOption_SecureProtocols) = 2048      'TLS 1.2
          '  xmldoc.Option(WinHttpRequestOption_SecureProtocols) = 512       'TLS 1.1
'                  xmldoc.Option(WinHttpRequestOption_SecureProtocols) = 128 Or &H20         'TLS 1.0
'                  xmldoc.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
'                  xmldoc.Option(WinHttpRequestOption_SelectCertificate) = mcert

            'xmldoc.Option(6) = False
            xmldoc.Send (xmlnf.xml & "")
            m_xml_send = xmlnf.xml
            
            
            If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
                'Debug.Print "---------------------"
                'Debug.Print pURL_NfeStatusServico
                'Debug.Print xmldom_resp.xml
                'Debug.Print "---------------------"
                m_xml_resp = xmldoc.ResponseText
                
                Dim ooo As IXMLDOMNode
                Dim ooo2 As IXMLDOMNode
                If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
                    If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
                        If ExistSelect(ooo.childNodes, ooo, "retEnvEvento") Then
                            If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
                                retNfeStatusServico.cStat = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "tpAmb") Then
                                retNfeStatusServico.tpAmb = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
                                retNfeStatusServico.xMotivo = ooo2.Text
                            End If
                        End If
                    End If
                End If
            Else
              MsgBox xmldom_resp.parseError.srcText, , "parseError - 2"
              
            End If
            NfeCancela = retNfeStatusServico.cStat = 107
        Else
            m_LastError = txt
            Debug.Print txt
        End If
    Else
        MsgBox xmlnf.parseError.srcText, , "parseError"
        SetFileBytes App.Path & "\error_1.txt", txt_xml
        SetFileBytes App.Path & "\error_2.txt", xmlnf.Text
    End If
    MsgBox retNfeStatusServico.xMotivo
    Set xmldoc = Nothing
    On Error GoTo 0
    Exit Function
    Resume Next
NfeStatusServico_Error:
    Debug.Print Err.Number, Err.Description
    If Err.Number = -2147012711 Then
        MsgBox "Erro de certificado"
    ElseIf Err.Number = -2147012852 Then
        MsgBox "? necess?rio um certificado para concluir a autentica??o do cliente"
    ElseIf Err.Number = -2147012851 Then
        MsgBox "A autoridade de certifica??o n?o ? v?lida ou est? incorreta"
        'https://www.projetoacbr.com.br/forum/topic/27273-a-autoridade-de-certifica%C3%A7%C3%A3o-n%C3%A3o-%C3%A9-valida-ou-esta-incorreta/
    Else
        MsgBox "Error linha: " & Erl & "," & Err.Number & " (" & Err.Description & ") in procedure NfeStatusServico of M?dulo mMainTeste"
        Debug.Print Err.Number,
    End If
End Function

Function get_msg(mCertificadoSel, ByVal n_cOrgao As String, ByVal n_chNFe As String, ByVal m_CNPJ As String, ByVal m_xJust As String, ByVal nProt As String) As String
Dim m_txt As New hStringBuilder
    m_txt.Clear
    m_txt.Append Trim("      <envEvento xmlns=""http://www.portalfiscal.inf.br/nfe"" versao=""1.00"">")
    'm_txt.Append Trim("        <idLote>201705051652506</idLote>")
    m_txt.Append Trim("        <idLote>" & Format(Date + Time, "YYYYMMDDHHMMSS") & "</idLote>")


    Dim m_ID As String
    Dim m_ID2 As String
    m_ID = "ID" & "110111" & n_chNFe & "01"



    m_txt.Append Trim("        <evento versao=""1.00"">")
    m_txt.Append Trim("          <infEvento Id=""" & m_ID & """>")
    m_txt.Append Trim("            <cOrgao>" & n_cOrgao & "</cOrgao>")
    m_txt.Append Trim("            <tpAmb>" & m_Ambiente & "</tpAmb>")
    m_txt.Append Trim("            <CNPJ>" & m_CNPJ & "</CNPJ>")
    m_txt.Append Trim("            <chNFe>" & n_chNFe & "</chNFe>")
    m_txt.Append Trim("            <dhEvento>" & cFormatDate_NF_1(Date + Time) & "</dhEvento>")
    m_txt.Append Trim("            <tpEvento>110111</tpEvento>")
    m_txt.Append Trim("            <nSeqEvento>1</nSeqEvento>")
    m_txt.Append Trim("            <verEvento>1.00</verEvento>")
    m_txt.Append Trim("            <detEvento versao=""1.00"">")
    m_txt.Append Trim("              <descEvento>Cancelamento</descEvento>")
    m_txt.Append Trim("              <nProt>" & nProt & "</nProt>")
    m_txt.Append Trim("              <xJust>" & m_xJust & "</xJust>")
    m_txt.Append Trim("            </detEvento>")
    m_txt.Append Trim("          </infEvento>")

    m_txt.Append Trim("       <Signature xmlns=""http://www.w3.org/2000/09/xmldsig#"">")
    m_txt.Append Trim("           <SignedInfo xmlns=""http://www.w3.org/2000/09/xmldsig#"">")
    m_txt.Append Trim("               <CanonicalizationMethod Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
    m_txt.Append Trim("               <SignatureMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#rsa-sha1""/>")
    m_txt.Append Trim("               <Reference URI=""#" & m_ID & """>")
    m_txt.Append Trim("                   <Transforms>")
    m_txt.Append Trim("                       <Transform Algorithm=""http://www.w3.org/2000/09/xmldsig#enveloped-signature""/>")
    m_txt.Append Trim("                       <Transform Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
    m_txt.Append Trim("                   </Transforms>")
    m_txt.Append Trim("                   <DigestMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#sha1""/>")
    m_txt.Append Trim("                   <DigestValue>yzGYhUx1/XYYzksWB+fPR3Qc50c=</DigestValue>")
    m_txt.Append Trim("               </Reference>")
    m_txt.Append Trim("           </SignedInfo>")
    m_txt.Append Trim("           <SignatureValue></SignatureValue>")
    m_txt.Append Trim("           <KeyInfo>")
    m_txt.Append Trim("               <X509Data>")
    m_txt.Append Trim("                   <X509Certificate></X509Certificate>")
    m_txt.Append Trim("               </X509Data>")
    m_txt.Append Trim("           </KeyInfo>")
    m_txt.Append Trim("       </Signature>")
    m_txt.Append Trim("        </evento>")
    m_txt.Append "      </envEvento>"


    Dim xmlnf As New MSXML2.DOMDocument50
    Dim xml_ass As String
    xmlnf.LoadXml m_txt.toString
    Dim NFe_Signature As New hNFe_Signature

    xml_ass = NFe_Signature.Assina(xmlnf.xml, mCertificadoSel)
    xmlnf.LoadXml xml_ass

    get_msg = xml_ass
End Function

Function NfeInutilizacao(mCertificadoSel As String, ByVal n_cOrgao As String, _
    ByVal m_CNPJ As String, ByVal m_mod As String, ByVal m_serie As String, _
    ByVal m_nNFIni As Long, ByVal m_nNFFin As Long, ByVal m_xJust As String, _
    m_xml_send As String, m_xml_resp As String) As Boolean

    Dim xmldoc As New WinHttpRequest
    Dim txt As String
    Dim xmlnf As New MSXML2.DOMDocument50
    Dim xmldom_resp As New MSXML2.DOMDocument50
    Dim txt_xml As String

    txt_xml = ""
    txt_xml = txt_xml & "<?xml version=""1.0"" encoding=""utf-8""?>"
    txt_xml = txt_xml & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
    txt_xml = txt_xml & "  <soap12:Body>"
    txt_xml = txt_xml & "    <nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeInutilizacao4"">"
    txt_xml = txt_xml & gera_inutNFe_msg(mCertificadoSel, n_cOrgao, m_CNPJ, m_mod, m_serie, m_nNFIni, m_nNFFin, m_xJust)
    txt_xml = txt_xml & "</nfeDadosMsg>"
    txt_xml = txt_xml & "  </soap12:Body>"
    txt_xml = txt_xml & "</soap12:Envelope>"

    If xmlnf.LoadXml(txt_xml) Then
        txt = ValidaSchema(xmlnf, "http://www.portalfiscal.inf.br/nfe", pURL_NFeInutilizacao_xsd, False)
        If txt = "Validado" Then
            Call xmldoc.Open("POST", pURL_NFeInutilizacao & "?wsdl", False)
            Call xmldoc.SetRequestHeader("content-type", "application/soap+xml")
            xmldoc.SetClientCertificate (mcert)
            xmldoc.Option(WinHttpRequestOption_SelectCertificate) = (mcert)
            xmldoc.Send (xmlnf.xml & "")
            m_xml_send = xmlnf.xml

            If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
                m_xml_resp = xmldoc.ResponseText

                Dim ooo As IXMLDOMNode
                Dim ooo2 As IXMLDOMNode
                If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
                    If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
                        If ExistSelect(ooo.childNodes, ooo, "retInutNFe") Then
                            If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
                                retNfeInutilizacao.cStat = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
                                retNfeInutilizacao.xMotivo = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo, "infInut") Then
                                If ExistSelect(ooo.childNodes, ooo2, "tpAmb") Then
                                    retNfeInutilizacao.tpAmb = ooo2.Text
                                End If
                                If ExistSelect(ooo.childNodes, ooo2, "cUF") Then
                                    retNfeInutilizacao.cUF = ooo2.Text
                                End If
                                If ExistSelect(ooo.childNodes, ooo2, "ano") Then
                                    retNfeInutilizacao.ano = ooo2.Text
                                End If
                                If ExistSelect(ooo.childNodes, ooo2, "CNPJ") Then
                                    retNfeInutilizacao.CNPJ = ooo2.Text
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                MsgBox xmldom_resp.parseError.srcText, , "parseError - 2"
            End If
            NfeInutilizacao = retNfeInutilizacao.cStat = 102
        Else
            m_LastError = txt
            Debug.Print txt
        End If
    Else
        MsgBox xmlnf.parseError.srcText, , "parseError"
        SetFileBytes App.Path & "\error_1.txt", txt_xml
        SetFileBytes App.Path & "\error_2.txt", xmlnf.Text
    End If
    Set xmldoc = Nothing
    On Error GoTo 0
    Exit Function
End Function

Private Function gera_inutNFe_msg(mCertificadoSel As String, ByVal n_cOrgao As String, _
    ByVal m_CNPJ As String, ByVal m_mod As String, ByVal m_serie As String, _
    ByVal m_nNFIni As Long, ByVal m_nNFFin As Long, ByVal m_xJust As String) As String

    Dim m_ano As String
    Dim m_ID As String
    Dim m_txt As New hStringBuilder

    m_ano = Format(Date, "YYYY")
    m_ID = "ID" & n_cOrgao & Format(Date, "YYYYMM") & OnlyNumbers(m_CNPJ) & m_mod & Format(Val(m_serie), "000") & Format(m_nNFIni, "000000000") & Format(m_nNFFin, "000000000")

    m_txt.Clear
    m_txt.Append "      <inutNFe xmlns=""http://www.portalfiscal.inf.br/nfe"" versao=""4.00"">"
    m_txt.Append "        <infInut Id=""" & m_ID & """>"
    m_txt.Append "          <tpAmb>" & m_Ambiente & "</tpAmb>"
    m_txt.Append "          <cUF>" & n_cOrgao & "</cUF>"
    m_txt.Append "          <ano>" & m_ano & "</ano>"
    m_txt.Append "          <CNPJ>" & OnlyNumbers(m_CNPJ) & "</CNPJ>"
    m_txt.Append "          <mod>" & m_mod & "</mod>"
    m_txt.Append "          <serie>" & m_serie & "</serie>"
    m_txt.Append "          <nNFIni>" & m_nNFIni & "</nNFIni>"
    m_txt.Append "          <nNFFin>" & m_nNFFin & "</nNFFin>"
    m_txt.Append "          <xJust>" & m_xJust & "</xJust>"
    m_txt.Append "        </infInut>"
    m_txt.Append "       <Signature xmlns=""http://www.w3.org/2000/09/xmldsig#"">"
    m_txt.Append "           <SignedInfo xmlns=""http://www.w3.org/2000/09/xmldsig#"">"
    m_txt.Append "               <CanonicalizationMethod Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>"
    m_txt.Append "               <SignatureMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#rsa-sha1""/>"
    m_txt.Append "               <Reference URI=""#" & m_ID & """>"
    m_txt.Append "                   <Transforms>"
    m_txt.Append "                       <Transform Algorithm=""http://www.w3.org/2000/09/xmldsig#enveloped-signature""/>"
    m_txt.Append "                       <Transform Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>"
    m_txt.Append "                   </Transforms>"
    m_txt.Append "                   <DigestMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#sha1""/>"
    m_txt.Append "                   <DigestValue>yzGYhUx1/XYYzksWB+fPR3Qc50c=</DigestValue>"
    m_txt.Append "               </Reference>"
    m_txt.Append "           </SignedInfo>"
    m_txt.Append "           <SignatureValue></SignatureValue>"
    m_txt.Append "           <KeyInfo>"
    m_txt.Append "               <X509Data>"
    m_txt.Append "                   <X509Certificate></X509Certificate>"
    m_txt.Append "               </X509Data>"
    m_txt.Append "           </KeyInfo>"
    m_txt.Append "       </Signature>"
    m_txt.Append "      </inutNFe>"

    Dim xmlnf As New MSXML2.DOMDocument50
    Dim xml_ass As String
    xmlnf.LoadXml m_txt.toString
    Dim NFe_Signature As New hNFe_Signature

    xml_ass = NFe_Signature.Assina(xmlnf.xml, mCertificadoSel)
    xmlnf.LoadXml xml_ass

    gera_inutNFe_msg = xml_ass
End Function

Function NfeCartaCorrecao(mCertificadoSel, ByVal n_cOrgao As String, ByVal n_chNFe As String, _
    ByVal m_CNPJ As String, ByVal m_xCorrecao As String, _
    m_xml_send As String, m_xml_resp As String) As Boolean

    Dim xmldoc As New WinHttpRequest
    Dim txt As String
    Dim xmlnf As New MSXML2.DOMDocument50
    Dim xmldom_resp As New MSXML2.DOMDocument50
    Dim txt_xml As String

    txt_xml = ""
    txt_xml = txt_xml & "<?xml version=""1.0"" encoding=""utf-8""?>"
    txt_xml = txt_xml & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
    txt_xml = txt_xml & "  <soap12:Body>"
    txt_xml = txt_xml & "    <nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeRecepcaoEvento4"">"
    txt_xml = txt_xml & gera_cc_e_msg(mCertificadoSel, n_cOrgao, n_chNFe, m_CNPJ, m_xCorrecao)
    txt_xml = txt_xml & "</nfeDadosMsg>"
    txt_xml = txt_xml & "  </soap12:Body>"
    txt_xml = txt_xml & "</soap12:Envelope>"

    If xmlnf.LoadXml(txt_xml) Then
        pURL_NFeRecepcaoEvento4_xsd = m_path & "envEvento_v1.00.xsd"
        txt = ValidaSchema(xmlnf, "http://www.portalfiscal.inf.br/nfe", pURL_NFeRecepcaoEvento4_xsd, False)
        If txt = "Validado" Then
            Call xmldoc.Open("POST", pURL_NFeRecepcaoEvento4 & "?wsdl", False)
            Call xmldoc.SetRequestHeader("content-type", "application/soap+xml")
            xmldoc.SetClientCertificate (mcert)
            xmldoc.Option(WinHttpRequestOption_SelectCertificate) = (mcert)
            xmldoc.Send (xmlnf.xml & "")
            m_xml_send = xmlnf.xml

            If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
                m_xml_resp = xmldoc.ResponseText

                Dim ooo As IXMLDOMNode
                Dim ooo2 As IXMLDOMNode
                If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
                    If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
                        If ExistSelect(ooo.childNodes, ooo, "retEnvEvento") Then
                            If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
                                retNfeCartaCorrecao.cStat = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "tpAmb") Then
                                retNfeCartaCorrecao.tpAmb = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
                                retNfeCartaCorrecao.xMotivo = ooo2.Text
                            End If
                        End If
                    End If
                End If
            Else
                MsgBox xmldom_resp.parseError.srcText, , "parseError - 2"
            End If
            NfeCartaCorrecao = retNfeCartaCorrecao.cStat = 135
        Else
            m_LastError = txt
            Debug.Print txt
        End If
    Else
        MsgBox xmlnf.parseError.srcText, , "parseError"
        SetFileBytes App.Path & "\error_1.txt", txt_xml
        SetFileBytes App.Path & "\error_2.txt", xmlnf.Text
    End If
    Set xmldoc = Nothing
    On Error GoTo 0
    Exit Function
End Function

Private Function gera_cc_e_msg(mCertificadoSel, ByVal n_cOrgao As String, _
    ByVal n_chNFe As String, ByVal m_CNPJ As String, _
    ByVal m_xCorrecao As String) As String

    Dim m_txt As New hStringBuilder
    m_txt.Clear
    m_txt.Append Trim("      <envEvento xmlns=""http://www.portalfiscal.inf.br/nfe"" versao=""1.00"">")
    m_txt.Append Trim("        <idLote>" & Format(Date + Time, "YYYYMMDDHHMMSS") & "</idLote>")

    Dim m_ID As String
    m_ID = "ID" & "110110" & n_chNFe & "01"

    m_txt.Append Trim("        <evento versao=""1.00"">")
    m_txt.Append Trim("          <infEvento Id=""" & m_ID & """>")
    m_txt.Append Trim("            <cOrgao>" & n_cOrgao & "</cOrgao>")
    m_txt.Append Trim("            <tpAmb>" & m_Ambiente & "</tpAmb>")
    m_txt.Append Trim("            <CNPJ>" & OnlyNumbers(m_CNPJ) & "</CNPJ>")
    m_txt.Append Trim("            <chNFe>" & n_chNFe & "</chNFe>")
    m_txt.Append Trim("            <dhEvento>" & cFormatDate_NF_1(Date + Time) & "</dhEvento>")
    m_txt.Append Trim("            <tpEvento>110110</tpEvento>")
    m_txt.Append Trim("            <nSeqEvento>1</nSeqEvento>")
    m_txt.Append Trim("            <verEvento>1.00</verEvento>")
    m_txt.Append Trim("            <detEvento versao=""1.00"">")
    m_txt.Append Trim("              <descEvento>Carta de Correcao</descEvento>")
    m_txt.Append Trim("              <xCorrecao>" & m_xCorrecao & "</xCorrecao>")
    m_txt.Append Trim("              <xCondUso>A Carta de Correcao e disciplinada pelo paragrafo 1o-A do art. 7o do Convenio S/N, de 15 de dezembro de 1970 e pode ser utilizada para regularizacao de erro ocorrido na emissao de documento fiscal, desde que o erro nao esteja relacionado com: I - as variaveis que determinam o valor do imposto tais como: base de calculo, aliquota, diferenca de preco, quantidade, valor da operacao ou da prestacao; II - a correcao de dados cadastrais que implique mudanca do remetente ou do destinatario; III - a data de emissao ou de saida.</xCondUso>")
    m_txt.Append Trim("            </detEvento>")
    m_txt.Append Trim("          </infEvento>")

    m_txt.Append Trim("       <Signature xmlns=""http://www.w3.org/2000/09/xmldsig#"">")
    m_txt.Append Trim("           <SignedInfo xmlns=""http://www.w3.org/2000/09/xmldsig#"">")
    m_txt.Append Trim("               <CanonicalizationMethod Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
    m_txt.Append Trim("               <SignatureMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#rsa-sha1""/>")
    m_txt.Append Trim("               <Reference URI=""#" & m_ID & """>")
    m_txt.Append Trim("                   <Transforms>")
    m_txt.Append Trim("                       <Transform Algorithm=""http://www.w3.org/2000/09/xmldsig#enveloped-signature""/>")
    m_txt.Append Trim("                       <Transform Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
    m_txt.Append Trim("                   </Transforms>")
    m_txt.Append Trim("                   <DigestMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#sha1""/>")
    m_txt.Append Trim("                   <DigestValue>yzGYhUx1/XYYzksWB+fPR3Qc50c=</DigestValue>")
    m_txt.Append Trim("               </Reference>")
    m_txt.Append Trim("           </SignedInfo>")
    m_txt.Append Trim("           <SignatureValue></SignatureValue>")
    m_txt.Append Trim("           <KeyInfo>")
    m_txt.Append Trim("               <X509Data>")
    m_txt.Append Trim("                   <X509Certificate></X509Certificate>")
    m_txt.Append Trim("               </X509Data>")
    m_txt.Append Trim("           </KeyInfo>")
    m_txt.Append Trim("       </Signature>")
    m_txt.Append Trim("        </evento>")
    m_txt.Append "      </envEvento>"

    Dim xmlnf As New MSXML2.DOMDocument50
    Dim xml_ass As String
    xmlnf.LoadXml m_txt.toString
    Dim NFe_Signature As New hNFe_Signature

    xml_ass = NFe_Signature.Assina(xmlnf.xml, mCertificadoSel)
    xmlnf.LoadXml xml_ass

    gera_cc_e_msg = xml_ass
End Function

Function NfeConsultaProtocolo(ByVal chNFe As String, m_xml_send As String, m_xml_resp As String) As Boolean
    Dim xmldoc As New WinHttpRequest
    Dim txt As String
    Dim xmlnf As New MSXML2.DOMDocument50
    Dim xmldom_resp As New MSXML2.DOMDocument50
    Dim txt_xml As String

    txt_xml = ""
    txt_xml = txt_xml & "<?xml version=""1.0"" encoding=""utf-8""?>"
    txt_xml = txt_xml & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
    txt_xml = txt_xml & "  <soap12:Body>"
    txt_xml = txt_xml & "    <nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4"">"
    txt_xml = txt_xml & "      <consSitNFe versao=""4.00"" xmlns=""http://www.portalfiscal.inf.br/nfe"">"
    txt_xml = txt_xml & "        <tpAmb>" & m_Ambiente & "</tpAmb>"
    txt_xml = txt_xml & "        <xServ>CONSULTAR</xServ>"
    txt_xml = txt_xml & "        <chNFe>" & chNFe & "</chNFe>"
    txt_xml = txt_xml & "      </consSitNFe>"
    txt_xml = txt_xml & "    </nfeDadosMsg>"
    txt_xml = txt_xml & "  </soap12:Body>"
    txt_xml = txt_xml & "</soap12:Envelope>"

    If xmlnf.LoadXml(txt_xml) Then
        txt = ValidaSchema(xmlnf, "http://www.portalfiscal.inf.br/nfe", pURL_NFeConsultaProtocolo_xsd, False)
        If txt = "Validado" Then
            Call xmldoc.Open("POST", pURL_NFeConsultaProtocolo & "?wsdl", False)
            Call xmldoc.SetRequestHeader("content-type", "application/soap+xml")
            xmldoc.SetClientCertificate (mcert)
            xmldoc.Option(WinHttpRequestOption_SelectCertificate) = (mcert)
            xmldoc.Send (xmlnf.xml & "")
            m_xml_send = xmlnf.xml

            If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
                m_xml_resp = xmldoc.ResponseText

                Dim ooo As IXMLDOMNode
                Dim ooo2 As IXMLDOMNode
                If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
                    If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
                        If ExistSelect(ooo.childNodes, ooo, "retConsSitNFe") Then
                            If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
                                retNFe_ConsSitNFe.cStat = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "tpAmb") Then
                                retNFe_ConsSitNFe.tpAmb = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
                                retNFe_ConsSitNFe.xMotivo = ooo2.Text
                            End If
                        End If
                    End If
                End If
            Else
                MsgBox xmldom_resp.parseError.srcText, , "parseError - 2"
            End If
            NfeConsultaProtocolo = retNFe_ConsSitNFe.cStat = 100 Or retNFe_ConsSitNFe.cStat = 101
        Else
            m_LastError = txt
            Debug.Print txt
        End If
    Else
        MsgBox xmlnf.parseError.srcText, , "parseError"
        SetFileBytes App.Path & "\error_1.txt", txt_xml
        SetFileBytes App.Path & "\error_2.txt", xmlnf.Text
    End If
    Set xmldoc = Nothing
    On Error GoTo 0
    Exit Function
End Function

Function NfeDistDFe(ByVal CNPJ As String, ByVal NSU As String, ByVal chNFe As String, m_xml_send As String, m_xml_resp As String) As Boolean
    Dim xmldoc As New WinHttpRequest
    Dim xmlnf As New MSXML2.DOMDocument50
    Dim xmldom_resp As New MSXML2.DOMDocument50
    Dim txt_xml As String

    txt_xml = ""
    txt_xml = txt_xml & "<?xml version=""1.0"" encoding=""utf-8""?>"
    txt_xml = txt_xml & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
    txt_xml = txt_xml & "  <soap12:Body>"
    txt_xml = txt_xml & "    <nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe"">"
    txt_xml = txt_xml & "      <distDFeInt versao=""1.01"" xmlns=""http://www.portalfiscal.inf.br/nfe"">"
    txt_xml = txt_xml & "        <tpAmb>" & m_Ambiente & "</tpAmb>"
    txt_xml = txt_xml & "        <cUFAutor>" & m_UF_Cod & "</cUFAutor>"
    txt_xml = txt_xml & "        <CNPJ>" & OnlyNumbers(CNPJ) & "</CNPJ>"

    If Len(chNFe) > 0 Then
        txt_xml = txt_xml & "        <consChNFe>"
        txt_xml = txt_xml & "          <chNFe>" & chNFe & "</chNFe>"
        txt_xml = txt_xml & "        </consChNFe>"
    Else
        txt_xml = txt_xml & "        <consNSU>"
        txt_xml = txt_xml & "          <NSU>" & NSU & "</NSU>"
        txt_xml = txt_xml & "        </consNSU>"
    End If

    txt_xml = txt_xml & "      </distDFeInt>"
    txt_xml = txt_xml & "    </nfeDadosMsg>"
    txt_xml = txt_xml & "  </soap12:Body>"
    txt_xml = txt_xml & "</soap12:Envelope>"

    If xmlnf.LoadXml(txt_xml) Then
        Call xmldoc.Open("POST", pURL_NFeDistribuicaoDFe & "?wsdl", False)
        Call xmldoc.SetRequestHeader("content-type", "application/soap+xml; charset=utf-8")
        Call xmldoc.SetRequestHeader("SOAPAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe/nfeDistDFeInteresse")
        xmldoc.SetClientCertificate (mcert)
        xmldoc.Option(WinHttpRequestOption_SelectCertificate) = (mcert)
        xmldoc.Send (xmlnf.xml & "")
        m_xml_send = xmlnf.xml

        If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
            m_xml_resp = xmldoc.ResponseText

            Dim ooo As IXMLDOMNode
            Dim ooo2 As IXMLDOMNode
            If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
                If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
                    If ExistSelect(ooo.childNodes, ooo, "retDistDFeInt") Then
                        If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
                            retNFe_DistDFe.cStat = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "tpAmb") Then
                            retNFe_DistDFe.tpAmb = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
                            retNFe_DistDFe.xMotivo = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "cUFAutor") Then
                            retNFe_DistDFe.cUFAutor = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "dhRecbto") Then
                            retNFe_DistDFe.dhRecbto = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "ultNSU") Then
                            retNFe_DistDFe.ultNSU = ooo2.Text
                        End If
                    End If
                End If
            End If
        Else
            MsgBox xmldom_resp.parseError.srcText, , "parseError - NfeDistDFe"
        End If
        NfeDistDFe = retNFe_DistDFe.cStat = 138 Or retNFe_DistDFe.cStat = 137
    Else
        MsgBox xmlnf.parseError.srcText, , "parseError"
        SetFileBytes App.Path & "\error_1.txt", txt_xml
    End If
    Set xmldoc = Nothing
    On Error GoTo 0
    Exit Function
End Function

Function NfeDownloadNF(ByVal cnpj As String, ByVal cUF As String, ByVal chNFe As String, m_xml_send As String, m_xml_resp As String) As Boolean
    Dim xmldoc As New WinHttpRequest
    Dim xmlnf As New MSXML2.DOMDocument50
    Dim xmldom_resp As New MSXML2.DOMDocument50
    Dim txt_xml As String

    txt_xml = ""
    txt_xml = txt_xml & "<?xml version=""1.0"" encoding=""utf-8""?>"
    txt_xml = txt_xml & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
    txt_xml = txt_xml & "  <soap12:Body>"
    txt_xml = txt_xml & "    <nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe"">"
    txt_xml = txt_xml & "      <downloadNFe versao=""1.01"" xmlns=""http://www.portalfiscal.inf.br/nfe"">"
    txt_xml = txt_xml & "        <tpAmb>" & m_Ambiente & "</tpAmb>"
    txt_xml = txt_xml & "        <cUF>" & cUF & "</cUF>"
    txt_xml = txt_xml & "        <CNPJ>" & OnlyNumbers(cnpj) & "</CNPJ>"
    txt_xml = txt_xml & "        <chNFe>" & chNFe & "</chNFe>"
    txt_xml = txt_xml & "      </downloadNFe>"
    txt_xml = txt_xml & "    </nfeDadosMsg>"
    txt_xml = txt_xml & "  </soap12:Body>"
    txt_xml = txt_xml & "</soap12:Envelope>"

    If xmlnf.LoadXml(txt_xml) Then
        Call xmldoc.Open("POST", pURL_NFeDistribuicaoDFe & "?wsdl", False)
        Call xmldoc.SetRequestHeader("content-type", "application/soap+xml; charset=utf-8")
        Call xmldoc.SetRequestHeader("SOAPAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe/nfeDownloadNF")
        xmldoc.SetClientCertificate (mcert)
        xmldoc.Option(WinHttpRequestOption_SelectCertificate) = (mcert)
        xmldoc.Send (xmlnf.xml & "")
        m_xml_send = xmlnf.xml

        If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
            m_xml_resp = xmldoc.ResponseText

            Dim ooo As IXMLDOMNode
            Dim ooo2 As IXMLDOMNode
            If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
                If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
                    If ExistSelect(ooo.childNodes, ooo, "retDownloadNFe") Then
                        If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
                            retNFe_DownloadNFe.cStat = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "tpAmb") Then
                            retNFe_DownloadNFe.tpAmb = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
                            retNFe_DownloadNFe.xMotivo = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo, "downloadNFe") Then
                            If ExistSelect(ooo.childNodes, ooo2, "chNFe") Then
                                retNFe_DownloadNFe.chNFe = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "docZip") Then
                                retNFe_DownloadNFe.docZip = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "NSU") Then
                                retNFe_DownloadNFe.NSU = ooo2.Text
                            End If
                        End If
                    End If
                End If
            End If
        Else
            MsgBox xmldom_resp.parseError.srcText, , "parseError - NfeDownloadNF"
        End If
        NfeDownloadNF = retNFe_DownloadNFe.cStat = 138 Or retNFe_DownloadNFe.cStat = 137
    Else
        MsgBox xmlnf.parseError.srcText, , "parseError"
        SetFileBytes App.Path & "\error_1.txt", txt_xml
    End If
    Set xmldoc = Nothing
    On Error GoTo 0
    Exit Function
End Function

Function NfeConsReciNFe(ByVal nRec As String, m_xml_send As String, m_xml_resp As String) As Boolean
    Dim xmldoc As New WinHttpRequest
    Dim txt As String
    Dim xmlnf As New MSXML2.DOMDocument50
    Dim xmldom_resp As New MSXML2.DOMDocument50
    Dim txt_xml As String

    txt_xml = ""
    txt_xml = txt_xml & "<?xml version=""1.0"" encoding=""utf-8""?>"
    txt_xml = txt_xml & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
    txt_xml = txt_xml & "  <soap12:Body>"
    txt_xml = txt_xml & "    <nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeRetAutorizacao4"">"
    txt_xml = txt_xml & "      <consReciNFe versao=""4.00"" xmlns=""http://www.portalfiscal.inf.br/nfe"">"
    txt_xml = txt_xml & "        <tpAmb>" & m_Ambiente & "</tpAmb>"
    txt_xml = txt_xml & "        <nRec>" & nRec & "</nRec>"
    txt_xml = txt_xml & "      </consReciNFe>"
    txt_xml = txt_xml & "    </nfeDadosMsg>"
    txt_xml = txt_xml & "  </soap12:Body>"
    txt_xml = txt_xml & "</soap12:Envelope>"

    If xmlnf.LoadXml(txt_xml) Then
        Call xmldoc.Open("POST", pURL_NFeRetAutorizacao4 & "?wsdl", False)
        Call xmldoc.SetRequestHeader("content-type", "application/soap+xml; charset=utf-8")
        Call xmldoc.SetRequestHeader("SOAPAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeRetAutorizacao4/nfeRetAutorizacaoLote")
        xmldoc.SetClientCertificate (mcert)
        xmldoc.Option(WinHttpRequestOption_SelectCertificate) = (mcert)
        xmldoc.Send (xmlnf.xml & "")
        m_xml_send = xmlnf.xml

        If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
            m_xml_resp = xmldoc.ResponseText

            Dim ooo As IXMLDOMNode
            Dim ooo2 As IXMLDOMNode
            If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
                If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
                    If ExistSelect(ooo.childNodes, ooo, "retConsReciNFe") Then
                        If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
                            retNFe_ConsReciNFe.cStat = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "tpAmb") Then
                            retNFe_ConsReciNFe.tpAmb = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
                            retNFe_ConsReciNFe.xMotivo = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "nRec") Then
                            retNFe_ConsReciNFe.nRec = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo2, "cUF") Then
                            retNFe_ConsReciNFe.cUF = ooo2.Text
                        End If
                        If ExistSelect(ooo.childNodes, ooo, "protNFe") Then
                            If ExistSelect(ooo.childNodes, ooo2, "chNFe") Then
                                retNFe_ConsReciNFe.protNFe.chNFe = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "nProt") Then
                                retNFe_ConsReciNFe.protNFe.nProt = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "digVal") Then
                                retNFe_ConsReciNFe.protNFe.digVal = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
                                retNFe_ConsReciNFe.protNFe.cStat = ooo2.Text
                            End If
                            If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
                                retNFe_ConsReciNFe.protNFe.xMotivo = ooo2.Text
                            End If
                        End If
                    End If
                End If
            End If
        Else
            MsgBox xmldom_resp.parseError.srcText, , "parseError - NfeConsReciNFe"
        End If
        NfeConsReciNFe = retNFe_ConsReciNFe.cStat = 104
    Else
        MsgBox xmlnf.parseError.srcText, , "parseError"
        SetFileBytes App.Path & "\error_1.txt", txt_xml
    End If
    Set xmldoc = Nothing
    On Error GoTo 0
    Exit Function
End Function

