VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    

'https://desenvolvedores.migrate.com.br/downloads/

 
Public m_infAdic As New hNFe4_Doc_infAdic
Public m_infRespTec As New hNFe4_Doc_infRespTec
Public m_infIntermed As New hNFe4_Doc_infIntermed
Public m_retirada As New hNFe4_Doc_local
Public m_entrega As New hNFe4_Doc_local
Public m_vTroco As Double               ' vTroco - valor do troco (configuravel)
Public m_ide As New hNFe4_Doc_ide
Public m_emit As New hNFe4_Doc_emit
Public m_dest As New hNFe4_Doc_dest
Public m_total As New hNFe4_Doc_total
Public m_Chave_NF As String
Public m_Ambiente As Long                   'Identifica????????????o de Ambiente: 1 - Produ????????????o, 2 - Homologa????????????o.
Public m_Endereco_QRCode As String
Public m_CSC As String
Public m_modFrete As String
Public m_transp As New hNFe4_Doc_transp
Public m_veicTransp As New hNFe4_Doc_veicTransp

Dim m_Itens() As hNFe4_Doc_det, m_cItens As Long
Dim m_Reboques() As hNFe4_Doc_veicTransp, m_cReboques As Long
Dim m_Volumes() As hNFe4_Doc_vol, m_cVolumes As Long
Dim m_Pagamentos() As hNFe4_Doc_detPag, m_cPagamentos As Long
Dim m_Dup() As hNFe4_Doc_cobr_dup, m_cDup As Long

Public CobrFat As New hNFe4_Doc_cobr_fat

Dim m_autXML_CNPJ() As String, m_autXML_CPF() As String, m_cAutXML As Long

Public Property Get ItemCount() As Long
    ItemCount = m_cItens + 1
End Property

Public Property Get Item(Index As Long) As hNFe4_Doc_det
    Set Item = m_Itens(Index)
End Property

Public Property Get PagCount() As Long
    PagCount = m_cPagamentos + 1
End Property

Public Property Get pag(Index As Long) As hNFe4_Doc_detPag
    Set pag = m_Pagamentos(Index)
End Property

Public Property Get DupCount() As Long
    DupCount = m_cDup + 1
End Property

Public Property Get Dup(Index As Long) As hNFe4_Doc_cobr_dup
    Set Dup = m_Dup(Index)
End Property

Sub Clear_dup()
    m_cDup = -1
    Erase m_Dup
End Sub

Sub Add_dup(m_Item As hNFe4_Doc_cobr_dup)
    m_cDup = m_cDup + 1
    ReDim Preserve m_Dup(m_cDup)
    Set m_Dup(m_cDup) = m_Item
End Sub

Sub Clear_det()
    m_cItens = -1
    Erase m_Itens
    m_total.Clear
End Sub

Sub Clear_pag()
    m_cPagamentos = -1
    Erase m_Pagamentos
End Sub

Sub Add_pag(m_Item As hNFe4_Doc_detPag)
    m_cPagamentos = m_cPagamentos + 1
    ReDim Preserve m_Pagamentos(m_cPagamentos)
    Set m_Pagamentos(m_cPagamentos) = m_Item
End Sub

' --- autXML (autoriados download XML, max 10) ---

Public Property Get AutXMLCount() As Long
    AutXMLCount = m_cAutXML + 1
End Property

Public Property Get AutXML_CNPJ(Index As Long) As String
    AutXML_CNPJ = m_autXML_CNPJ(Index)
End Property

Public Property Get AutXML_CPF(Index As Long) As String
    AutXML_CPF = m_autXML_CPF(Index)
End Property

Sub Clear_autXML()
    m_cAutXML = -1
    Erase m_autXML_CNPJ
    Erase m_autXML_CPF
End Sub

Sub Add_autXML(m_CNPJ As String, Optional m_CPF As String)
    m_cAutXML = m_cAutXML + 1
    ReDim Preserve m_autXML_CNPJ(m_cAutXML)
    ReDim Preserve m_autXML_CPF(m_cAutXML)
    m_autXML_CNPJ(m_cAutXML) = m_CNPJ
    m_autXML_CPF(m_cAutXML) = m_CPF
End Sub

Sub Add_det(m_Item As hNFe4_Doc_det, Optional m_flag As Long)
    
    
    If m_cItens = -1 And m_Ambiente = 2 And m_flag = 0 Then
        Dim m_Item2 As New hNFe4_Doc_det
        
        Set m_Item2 = m_Item
        With m_Item
            .prod.m_cProd = "1"
            .prod.m_xProd = "NOTA FISCAL EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
            '.prod.m_uCom = 1
            '.prod.m_qCom = 1
            '.prod.m_indTot = 1
            '.prod.m_vProd = 1
            '.prod.m_vUnCom = 1
            '.prod.m_vUnTrib = 1
            '.imposto.m_vTotTrib = 1
        End With
        Add_det m_Item2, 1
        Exit Sub
    End If
    
    m_cItens = m_cItens + 1
    ReDim Preserve m_Itens(m_cItens)
    Set m_Itens(m_cItens) = m_Item
    
    ' --- ICMS ---
    If m_Item.imposto.m_ICMS_vBC <> 0 Or m_Item.imposto.m_ICMS_vICMS <> 0 Then
        m_total.m_vBC = m_total.m_vBC + m_Item.imposto.m_ICMS_vBC
        m_total.m_vICMS = m_total.m_vICMS + m_Item.imposto.m_ICMS_vICMS
    Else
        m_total.m_vBC = m_total.m_vBC + m_Item.imposto.m_ICMS_ICMS00_vBC
        m_total.m_vICMS = m_total.m_vICMS + m_Item.imposto.m_ICMS_ICMS00_vICMS
    End If
    m_total.m_vBCST = m_total.m_vBCST + m_Item.imposto.m_ICMS_vBCST
    m_total.m_vST = m_total.m_vST + m_Item.imposto.m_ICMS_vICMSST
    m_total.m_vICMSDeson = m_total.m_vICMSDeson + m_Item.imposto.m_ICMS_vICMSDeson
    m_total.m_vFCPUFDest = m_total.m_vFCPUFDest + m_Item.imposto.m_ICMSUFDest_vFCPUFDest
    m_total.m_vICMSUFDest = m_total.m_vICMSUFDest + m_Item.imposto.m_ICMSUFDest_vICMSUFDest
    m_total.m_vICMSUFRemet = m_total.m_vICMSUFRemet + m_Item.imposto.m_ICMSUFDest_vICMSUFRemet
    
    ' --- Produto / Frete ---
    m_total.m_vProd = m_total.m_vProd + m_Item.prod.m_vProd
    m_total.m_vFrete = m_total.m_vFrete + m_Item.prod.m_vFrete
    
    ' --- IPI ---
    m_total.m_vIPI = m_total.m_vIPI + m_Item.imposto.m_IPI_vIPI

    ' --- II (Imposto de Importacao) ---
    m_total.m_vII = m_total.m_vII + m_Item.imposto.m_II_vII

    ' --- PIS / COFINS ---
    m_total.m_vPIS = m_total.m_vPIS + IIf(m_Item.imposto.m_PIS_vPIS <> 0, m_Item.imposto.m_PIS_vPIS, m_Item.imposto.m_PIS_PISAliq_vPIS)
    m_total.m_vCOFINS = m_total.m_vCOFINS + IIf(m_Item.imposto.m_COFINS_vCOFINS <> 0, m_Item.imposto.m_COFINS_vCOFINS, m_Item.imposto.m_COFINS_COFINSAliq_vCOFINS)
    
    ' --- vTotTrib ---
    m_total.m_vTotTrib = m_total.m_vTotTrib + m_Item.imposto.m_vTotTrib
    
    ' --- IBS / CBS / IS (Reforma Tributaria - NT 2025.002) ---
    m_total.m_vBCIBSCBS = m_total.m_vBCIBSCBS + m_Item.imposto.m_IBSCBS_vBC
    m_total.m_vIBSUF = m_total.m_vIBSUF + m_Item.imposto.m_IBSCBS_vIBSUF
    m_total.m_vIBSMun = m_total.m_vIBSMun + m_Item.imposto.m_IBSCBS_vIBSMun
    m_total.m_vIBS = m_total.m_vIBS + m_Item.imposto.m_IBSCBS_vIBS
    m_total.m_vCBS = m_total.m_vCBS + m_Item.imposto.m_IBSCBS_vCBS
    m_total.m_vCBSST = m_total.m_vCBSST + m_Item.imposto.m_IBSCBS_vCBSST
    m_total.m_vIS = m_total.m_vIS + m_Item.imposto.m_IS_vIS
    ' gDif / gDevTrib totals
    m_total.m_vDifIBSUF = m_total.m_vDifIBSUF + m_Item.imposto.m_IBSCBS_gIBSUF_vDif
    m_total.m_vDevTribIBSUF = m_total.m_vDevTribIBSUF + m_Item.imposto.m_IBSCBS_gIBSUF_vDevTrib
    m_total.m_vDifIBSMun = m_total.m_vDifIBSMun + m_Item.imposto.m_IBSCBS_gIBSMun_vDif
    m_total.m_vDevTribIBSMun = m_total.m_vDevTribIBSMun + m_Item.imposto.m_IBSCBS_gIBSMun_vDevTrib
    m_total.m_vDifCBS = m_total.m_vDifCBS + m_Item.imposto.m_IBSCBS_gCBS_vDif
    m_total.m_vDevTribCBS = m_total.m_vDevTribCBS + m_Item.imposto.m_IBSCBS_gCBS_vDevTrib
    
    ' --- FCP ---
    m_total.m_vFCP = m_total.m_vFCP + m_Item.imposto.m_vFCP
    m_total.m_vFCPST = m_total.m_vFCPST + m_Item.imposto.m_vFCPST
    m_total.m_vFCPSTRet = m_total.m_vFCPSTRet + m_Item.imposto.m_vFCPSTRet
End Sub



Sub Clear_reboques()
    m_cReboques = -1
    Erase m_Reboques
End Sub

Sub Add_reboque(m_Reboque As hNFe4_Doc_veicTransp)
    m_cReboques = m_cReboques + 1
    ReDim Preserve m_Reboques(m_cReboques)
    Set m_Reboques(m_cReboques) = m_Reboque
End Sub

Sub Clear_volumes()
    m_cVolumes = -1
    Erase m_Volumes
End Sub

Sub Add_volume(m_Vol As hNFe4_Doc_vol)
    m_cVolumes = m_cVolumes + 1
    ReDim Preserve m_Volumes(m_cVolumes)
    Set m_Volumes(m_cVolumes) = m_Vol
End Sub

Private Function Calculo_DV11(strNumero As String) As String
    Dim intcontador As Integer, intnumero As Integer
    Dim intTotalNumero As Integer, intMultiplicador As Integer, intResto As Integer
    Dim ch As String

    'NT 2026.004: chave aceita alfanumericos  converte letras p/ ASCII
    intMultiplicador = 9

    For intcontador = Len(strNumero) To 1 Step -1
        ch = Mid(strNumero, intcontador, 1)
        If ch Like "[0-9]" Then
            intnumero = Val(ch)
        ElseIf ch Like "[A-Za-z]" Then
            intnumero = Asc(UCase(ch))
        Else
            intnumero = 0
        End If
        intTotalNumero = intTotalNumero + (intnumero * intMultiplicador)
        intMultiplicador = IIf(intMultiplicador > 2, intMultiplicador - 1, 9)
    Next

    intResto = intTotalNumero Mod 11

    Select Case intResto
    Case 0
        Calculo_DV11 = "0"
    Case 10
        Calculo_DV11 = "0"
    Case Else
        Calculo_DV11 = Trim(Str(intResto))
    End Select

End Function

Function GetNota(m_Chave_2 As String, m_NF_NUM_NOTA As Long) As String
    Dim txt_xml   As New hStringBuilder
    Dim txt_xml2   As String
    Dim m_Chave   As New hStringBuilder
    Dim m_ChaveDV As String
    Dim cc        As Long
    Dim m_HashQRCode As String

    m_NF_NUM_NOTA = m_ide.m_nNF
    m_Chave.Clear
    m_Chave.Append m_ide.m_cUF
    m_Chave.Append Format(m_ide.m_dhEmi, "YYMM")
    m_Chave.Append m_emit.m_CNPJ
    m_Chave.Append m_ide.m_mod
    m_Chave.Append Format(Val(m_ide.m_serie), "000")
    m_Chave.Append Format(Val(m_ide.m_nNF), "000000000")
    m_Chave.Append m_ide.m_tpEmis
    m_Chave.Append Format(Val(m_ide.m_cNF), "00000000")
    m_ChaveDV = Calculo_DV11(m_Chave.toString)
    m_Chave.Append m_ChaveDV
    m_Chave_2 = m_Chave
    m_Chave_NF = m_Chave
    
    txt_xml.Append Trim("   <NFe xmlns=""http://www.portalfiscal.inf.br/nfe"">")
    txt_xml.Append Trim("       <infNFe Id=""" & "NFe" & m_Chave & """ versao=""4.00"">")
    
    GetNota_BuildIde txt_xml, m_ChaveDV
    GetNota_BuildRetiradaEntrega txt_xml
    GetNota_BuildAutXML txt_xml
    GetNota_BuildEmit txt_xml
    GetNota_BuildDest txt_xml
    GetNota_BuildInfIntermed txt_xml
    
    GetNota_BuildDet txt_xml
    GetNota_BuildTotal txt_xml
    GetNota_BuildTransp txt_xml
    GetNota_BuildCobr txt_xml
    GetNota_BuildPag txt_xml
    GetNota_BuildInfAdic txt_xml
    GetNota_BuildInfRespTec txt_xml

    txt_xml.Append Trim("       </infNFe>")
    
    GetNota_BuildSignature txt_xml, m_Chave
    
    GetNota = (txt_xml)
End Function

Public Function InsertQRCode(txt_xml As String) As Boolean
Dim xml As New DOMDocument50
Dim digVal As String
Dim ooo As IXMLDOMNode
Dim txt As String
Dim txt_out As String
Dim oooSim As IXMLDOMNode
Dim Pos As Long
    If m_ide.m_mod = "55" Then
        InsertQRCode = True
    ElseIf m_ide.m_mod = "65" Then
        If xml.LoadXml(txt_xml) Then
            If ExistSelect(xml.documentElement.childNodes, oooSim, "Signature") Then
                If ExistSelect(oooSim.childNodes, ooo, "SignedInfo") Then
                    If ExistSelect(ooo.childNodes, ooo, "Reference") Then
                        If ExistSelect(ooo.childNodes, ooo, "DigestValue") Then
                            digVal = ooo.Text
                        End If
                        'retNfeStatusServico.cStat = ooo2.text
                    End If
                End If
            End If
            If Len(digVal) > 0 Then
                txt = GetQRCode(digVal)
                Pos = InStr(txt_xml, "</infNFe>")
                If Pos > 0 Then
                    txt_out = Left(txt_xml, Pos + 8)
                    txt_out = txt_out & txt
                    txt_out = txt_out & Mid(txt_xml, Pos + 9)
                    txt_xml = txt_out
                    
                    'Debug.Print Left(txt_xml, pos + 8)
                    'Debug.Print mid(txt_xml, pos + 9)
                    'Debug.Print txt
                    'GetLogger.LogDebug txt_xml
                    'Debug.Print txt_out
                    'ooon.Load txt
                    'MsgBox ooon.documentElement.childNodes(1).text
                    'xml.insertBefore ooon.selectSingleNode("infNFeSupl"), xml.childNodes.Item(1)
                    'txt = Clipboard.GetText
                    'Debug.Print txt
                    InsertQRCode = True
                End If
            End If
        End If
    End If
End Function

Private Function GetQRCode(ByVal digVal As String) As String
Dim txt_xml As String
Dim txt_xml2 As String
Dim hash As New hSHA1
Dim m_HashQRCode  As String


    txt_xml = txt_xml & "<infNFeSupl>"
    txt_xml = txt_xml & "<qrCode>"
    txt_xml = txt_xml & "<![CDATA["
    
'    txt_xml2 = txt_xml2 & "?chNFe=" & m_Chave_NF
'    txt_xml2 = txt_xml2 & "&nVersao=200"
'    txt_xml2 = txt_xml2 & "&tpAmb=" & m_Ambiente
'    If Not m_dest.m_CNPJ & m_dest.m_CPF = "" Then
'        txt_xml2 = txt_xml2 & "&cDest=" & m_dest.m_CNPJ & m_dest.m_CPF
'    End If
'    txt_xml2 = txt_xml2 & "&dhEmi=" & (Base16.Encode(cFormatDate_NF_1(m_ide.m_dhEmi)))
'    txt_xml2 = txt_xml2 & "&vNF=" & cFormatMoeda_NF_1(m_total.m_vNF)
'    txt_xml2 = txt_xml2 & "&vICMS=" & cFormatMoeda_NF_1(m_total.m_vICMS)
'    txt_xml2 = txt_xml2 & "&digVal=" & Base16.Encode(digVal)
'    txt_xml2 = txt_xml2 & "&cIdToken=1"
    
    
    txt_xml2 = ""
    txt_xml2 = txt_xml2 & m_Chave_NF
    txt_xml2 = txt_xml2 & "|2"
    txt_xml2 = txt_xml2 & "|" & m_Ambiente
    txt_xml2 = txt_xml2 & "|1"
    txt_xml2 = txt_xml2 & m_CSC
    
    m_HashQRCode = hash.SHA1AA1Hash(txt_xml2)
    
    
    'txt_xml = txt_xml & "http://webas.sefaz.pi.gov.br/nfceweb-homologacao/consultarNFCe.jsf"
    'If m_UF = "PI" Then
    '    txt_xml = txt_xml & "http://www.sefaz.pi.gov.br/nfce/qrcode"
    'ElseIf m_UF = "MA" Then
    '    txt_xml = txt_xml & "http://nfce.sefaz.ma.gov.br/portal/consultarNFCe.jsp"
    'End If
    txt_xml = txt_xml & m_Endereco_QRCode
    
    'txt_xml = txt_xml & "http://www.sefaz.pi.gov.br/nfce/consulta"
    txt_xml = txt_xml & "?p="
    txt_xml = txt_xml & m_Chave_NF          'chave_acesso
    txt_xml = txt_xml & "|2"                'versao_qrcode
    txt_xml = txt_xml & "|" & m_Ambiente    'tipo_ambiente
    txt_xml = txt_xml & "|1"           'identificador_csc
    txt_xml = txt_xml & "|" & m_HashQRCode  'codigo_hash
    
    txt_xml = txt_xml & "]]>"
    txt_xml = txt_xml & "</qrCode>"
    GetLogger.LogDebug txt_xml
    
'    If False Then
'
'        txt_xml = txt_xml & "<![CDATA[" & m_Endereco_QRCode
'        txt_xml2 = txt_xml2 & "?chNFe=" & m_Chave_NF
'        txt_xml2 = txt_xml2 & "&nVersao=200"
'        txt_xml2 = txt_xml2 & "&tpAmb=" & m_Ambiente
'        If m_dest.m_CNPJ & m_dest.m_CPF = "" Then
'            'txt_xml2 = txt_xml2 & "&cDest=00000000000"
'        Else
'            txt_xml2 = txt_xml2 & "&cDest=" & m_dest.m_CNPJ & m_dest.m_CPF
'        End If
'        txt_xml2 = txt_xml2 & "&dhEmi=" & (Base16.Encode(cFormatDate_NF_1(m_ide.m_dhEmi)))
'        txt_xml2 = txt_xml2 & "&vNF=" & cFormatMoeda_NF_1(m_total.m_vNF)
'        txt_xml2 = txt_xml2 & "&vICMS=" & cFormatMoeda_NF_1(m_total.m_vICMS)
'        txt_xml2 = txt_xml2 & "&digVal=" & Base16.Encode(digVal)
'        txt_xml2 = txt_xml2 & "&cIdToken=1"
'        m_HashQRCode = hash.SHA1AA1Hash(txt_xml2 & m_CSC)
'        txt_xml = txt_xml & txt_xml2
'        txt_xml = txt_xml & "&cHashQRCode=" & m_HashQRCode & "]]>"
'        txt_xml = txt_xml & "</qrCode>"
'    End If
    
    txt_xml = txt_xml & "<urlChave>"
    txt_xml = txt_xml & m_Endereco_QRCode
    txt_xml = txt_xml & "</urlChave>"
    txt_xml = txt_xml & "</infNFeSupl>"
    GetQRCode = txt_xml
End Function

Function cFormatMoeda_NF_1(ByVal m_valor As Double) As String
    cFormatMoeda_NF_1 = Format(m_valor, "0.00")
    cFormatMoeda_NF_1 = Replace(cFormatMoeda_NF_1, ",", ".")
End Function

Private Sub GetNota_BuildIde(txt_xml As hStringBuilder, m_ChaveDV As String)
    txt_xml.Append Trim("           <ide>")
    txt_xml.Append Trim("               <cUF>" & m_ide.m_cUF & "</cUF>")
    txt_xml.Append Trim("               <cNF>" & m_ide.m_cNF & "</cNF>")
    If Len(m_ide.m_natOp) > 0 Then txt_xml.Append Trim("               <natOp>" & m_ide.m_natOp & "</natOp>")
    If Len(m_ide.m_indPag) > 0 Then txt_xml.Append Trim("               <indPag>" & m_ide.m_indPag & "</indPag>")
    txt_xml.Append Trim("               <mod>" & m_ide.m_mod & "</mod>")
    txt_xml.Append Trim("               <serie>" & m_ide.m_serie & "</serie>")
    txt_xml.Append Trim("               <nNF>" & m_ide.m_nNF & "</nNF>")
    txt_xml.Append Trim("               <dhEmi>" & cFormatDate_NF_1(m_ide.m_dhEmi) & "</dhEmi>")
    If m_ide.m_mod = "55" Then txt_xml.Append Trim("               <dhSaiEnt>" & cFormatDate_NF_1(m_ide.m_dhSaiEnt) & "</dhSaiEnt>")
    If Len(m_ide.m_tpNF) > 0 Then txt_xml.Append Trim("               <tpNF>" & m_ide.m_tpNF & "</tpNF>")
    txt_xml.Append Trim("               <idDest>" & m_ide.m_idDest & "</idDest>")
    txt_xml.Append Trim("               <cMunFG>" & m_ide.m_cMunFG & "</cMunFG>")
    If Len(m_ide.m_cMunFGIBS) > 0 Then txt_xml.Append Trim("               <cMunFGIBS>" & m_ide.m_cMunFGIBS & "</cMunFGIBS>")
    If m_ide.m_mod = "55" Then
        txt_xml.Append Trim("               <tpImp>1</tpImp>")
    ElseIf m_ide.m_mod = "65" Then
        txt_xml.Append Trim("               <tpImp>4</tpImp>")
    End If
    txt_xml.Append Trim("               <tpEmis>" & m_ide.m_tpEmis & "</tpEmis>")
    txt_xml.Append Trim("               <cDV>" & m_ChaveDV & "</cDV>")
    txt_xml.Append Trim("               <tpAmb>" & m_Ambiente & "</tpAmb>")
    If Len(m_ide.m_finNFe) > 0 Then txt_xml.Append Trim("               <finNFe>" & m_ide.m_finNFe & "</finNFe>")
    If Len(m_ide.m_tpNFDebito) > 0 Then txt_xml.Append Trim("               <tpNFDebito>" & m_ide.m_tpNFDebito & "</tpNFDebito>")
    If Len(m_ide.m_tpNFCredito) > 0 Then txt_xml.Append Trim("               <tpNFCredito>" & m_ide.m_tpNFCredito & "</tpNFCredito>")
    If Len(m_ide.m_indFinal) > 0 Then txt_xml.Append Trim("               <indFinal>" & m_ide.m_indFinal & "</indFinal>")
    If Len(m_ide.m_indPres) > 0 Then txt_xml.Append Trim("               <indPres>" & m_ide.m_indPres & "</indPres>")
    If Len(m_ide.m_indIntermed) > 0 Then txt_xml.Append Trim("               <indIntermed>" & m_ide.m_indIntermed & "</indIntermed>")
    If Len(m_ide.m_procEmi) > 0 Then txt_xml.Append Trim("               <procEmi>" & m_ide.m_procEmi & "</procEmi>")
    txt_xml.Append Trim("               <verProc>" & App.Major & "." & App.Minor & "." & App.Revision & "</verProc>")
    If m_ide.m_tpEmis <> "1" Then
        If Len(m_ide.m_dhCont) > 0 Then txt_xml.Append Trim("               <dhCont>" & m_ide.m_dhCont & "</dhCont>")
        If Len(m_ide.m_xJust) > 0 Then txt_xml.Append Trim("               <xJust>" & m_ide.m_xJust & "</xJust>")
    End If
    If m_ide.NFrefCount > 0 Then
        Dim cc_nfref As Long
        For cc_nfref = 0 To m_ide.NFrefCount - 1
            txt_xml.Append Trim("               <NFref>")
            txt_xml.Append Trim("                   <refNFe>" & m_ide.NFref_refNFe(cc_nfref) & "</refNFe>")
            txt_xml.Append Trim("               </NFref>")
        Next
    End If
    txt_xml.Append Trim("           </ide>")
End Sub

Private Sub GetNota_BuildRetiradaEntrega(txt_xml As hStringBuilder)
    If Len(m_retirada.m_xLgr) > 0 Then
        txt_xml.Append Trim("           <retirada>")
        If Len(m_retirada.m_CNPJ) > 0 Then
            txt_xml.Append Trim("               <CNPJ>" & m_retirada.m_CNPJ & "</CNPJ>")
        ElseIf Len(m_retirada.m_CPF) > 0 Then
            txt_xml.Append Trim("               <CPF>" & m_retirada.m_CPF & "</CPF>")
        End If
        If Len(m_retirada.m_xNome) > 0 Then txt_xml.Append Trim("               <xNome>" & m_retirada.m_xNome & "</xNome>")
        txt_xml.Append Trim("               <xLgr>" & m_retirada.m_xLgr & "</xLgr>")
        txt_xml.Append Trim("               <nro>" & m_retirada.m_nro & "</nro>")
        If Len(m_retirada.m_xCpl) > 0 Then txt_xml.Append Trim("               <xCpl>" & m_retirada.m_xCpl & "</xCpl>")
        txt_xml.Append Trim("               <xBairro>" & m_retirada.m_xBairro & "</xBairro>")
        txt_xml.Append Trim("               <cMun>" & m_retirada.m_cMun & "</cMun>")
        txt_xml.Append Trim("               <xMun>" & m_retirada.m_xMun & "</xMun>")
        txt_xml.Append Trim("               <UF>" & m_retirada.m_UF & "</UF>")
        If Len(m_retirada.m_CEP) > 0 Then txt_xml.Append Trim("               <CEP>" & m_retirada.m_CEP & "</CEP>")
        If Len(m_retirada.m_cPais) > 0 Then txt_xml.Append Trim("               <cPais>" & m_retirada.m_cPais & "</cPais>")
        If Len(m_retirada.m_xPais) > 0 Then txt_xml.Append Trim("               <xPais>" & m_retirada.m_xPais & "</xPais>")
        If Len(m_retirada.m_fone) > 0 Then txt_xml.Append Trim("               <fone>" & m_retirada.m_fone & "</fone>")
        If Len(m_retirada.m_email) > 0 Then txt_xml.Append Trim("               <email>" & m_retirada.m_email & "</email>")
        If Len(m_retirada.m_IE) > 0 Then txt_xml.Append Trim("               <IE>" & m_retirada.m_IE & "</IE>")
        txt_xml.Append Trim("           </retirada>")
    End If
    If Len(m_entrega.m_xLgr) > 0 Then
        txt_xml.Append Trim("           <entrega>")
        If Len(m_entrega.m_CNPJ) > 0 Then
            txt_xml.Append Trim("               <CNPJ>" & m_entrega.m_CNPJ & "</CNPJ>")
        ElseIf Len(m_entrega.m_CPF) > 0 Then
            txt_xml.Append Trim("               <CPF>" & m_entrega.m_CPF & "</CPF>")
        End If
        If Len(m_entrega.m_xNome) > 0 Then txt_xml.Append Trim("               <xNome>" & m_entrega.m_xNome & "</xNome>")
        txt_xml.Append Trim("               <xLgr>" & m_entrega.m_xLgr & "</xLgr>")
        txt_xml.Append Trim("               <nro>" & m_entrega.m_nro & "</nro>")
        If Len(m_entrega.m_xCpl) > 0 Then txt_xml.Append Trim("               <xCpl>" & m_entrega.m_xCpl & "</xCpl>")
        txt_xml.Append Trim("               <xBairro>" & m_entrega.m_xBairro & "</xBairro>")
        txt_xml.Append Trim("               <cMun>" & m_entrega.m_cMun & "</cMun>")
        txt_xml.Append Trim("               <xMun>" & m_entrega.m_xMun & "</xMun>")
        txt_xml.Append Trim("               <UF>" & m_entrega.m_UF & "</UF>")
        If Len(m_entrega.m_CEP) > 0 Then txt_xml.Append Trim("               <CEP>" & m_entrega.m_CEP & "</CEP>")
        If Len(m_entrega.m_cPais) > 0 Then txt_xml.Append Trim("               <cPais>" & m_entrega.m_cPais & "</cPais>")
        If Len(m_entrega.m_xPais) > 0 Then txt_xml.Append Trim("               <xPais>" & m_entrega.m_xPais & "</xPais>")
        If Len(m_entrega.m_fone) > 0 Then txt_xml.Append Trim("               <fone>" & m_entrega.m_fone & "</fone>")
        If Len(m_entrega.m_email) > 0 Then txt_xml.Append Trim("               <email>" & m_entrega.m_email & "</email>")
        If Len(m_entrega.m_IE) > 0 Then txt_xml.Append Trim("               <IE>" & m_entrega.m_IE & "</IE>")
        txt_xml.Append Trim("           </entrega>")
    End If
End Sub

Private Sub GetNota_BuildAutXML(txt_xml As hStringBuilder)
    If m_cAutXML >= 0 Then
        Dim cc_autXML As Long
        For cc_autXML = 0 To m_cAutXML
            txt_xml.Append Trim("           <autXML>")
            If Len(m_autXML_CNPJ(cc_autXML)) > 0 Then txt_xml.Append Trim("               <CNPJ>" & m_autXML_CNPJ(cc_autXML) & "</CNPJ>") Else txt_xml.Append Trim("               <CPF>" & m_autXML_CPF(cc_autXML) & "</CPF>")
            txt_xml.Append Trim("           </autXML>")
        Next
    End If
End Sub

Private Sub GetNota_BuildEmit(txt_xml As hStringBuilder)
    txt_xml.Append Trim("           <emit>")
    txt_xml.Append Trim("               <CNPJ>" & m_emit.m_CNPJ & "</CNPJ>")
    txt_xml.Append Trim("               <xNome>" & m_emit.m_xNome & "</xNome>")
    txt_xml.Append Trim("               <xFant>" & m_emit.m_xFant & "</xFant>")
    txt_xml.Append Trim("               <enderEmit>")
    txt_xml.Append Trim("                   <xLgr>" & m_emit.m_enderEmit.m_xLgr & "</xLgr>")
    txt_xml.Append Trim("                   <nro>" & m_emit.m_enderEmit.m_nro & "</nro>")
    txt_xml.Append Trim("                   <xBairro>" & m_emit.m_enderEmit.m_xBairro & "</xBairro>")
    txt_xml.Append Trim("                   <cMun>" & m_emit.m_enderEmit.m_cMun & "</cMun>")
    txt_xml.Append Trim("                   <xMun>" & m_emit.m_enderEmit.m_xMun & "</xMun>")
    txt_xml.Append Trim("                   <UF>" & m_emit.m_enderEmit.m_UF & "</UF>")
    txt_xml.Append Trim("                   <CEP>" & m_emit.m_enderEmit.m_CEP & "</CEP>")
    txt_xml.Append Trim("                   <cPais>" & m_emit.m_enderEmit.m_cPais & "</cPais>")
    txt_xml.Append Trim("                   <xPais>" & m_emit.m_enderEmit.m_xPais & "</xPais>")
    txt_xml.Append Trim("                   <fone>" & m_emit.m_enderEmit.m_fone & "</fone>")
    txt_xml.Append Trim("               </enderEmit>")
    txt_xml.Append Trim("               <IE>" & m_emit.m_IE & "</IE>")
    If Len(m_emit.m_IM) Then
        txt_xml.Append Trim("               <IM>" & m_emit.m_IM & "</IM>")
        txt_xml.Append Trim("               <CNAE>" & m_emit.m_CNAE & "</CNAE>")
    End If
    txt_xml.Append Trim("               <CRT>" & m_emit.m_CRT & "</CRT>")
    txt_xml.Append Trim("           </emit>")
End Sub

Private Sub GetNota_BuildDest(txt_xml As hStringBuilder)
    If Len(m_dest.m_CNPJ) > 0 Or Len(m_dest.m_CPF) > 0 Then
        txt_xml.Append Trim("           <dest>")
        If Not m_dest.m_CNPJ = "" Then
            txt_xml.Append Trim("               <CNPJ>" & m_dest.m_CNPJ & "</CNPJ>")
        ElseIf Not m_dest.m_CPF = "" Then
            txt_xml.Append Trim("               <CPF>" & m_dest.m_CPF & "</CPF>")
        End If
        If Len(m_dest.m_xNome) > 0 Then txt_xml.Append Trim("               <xNome>" & m_dest.m_xNome & "</xNome>")
        If Len(m_dest.m_enderDest.m_xLgr) > 0 Then
            txt_xml.Append Trim("               <enderDest>")
            txt_xml.Append Trim("                   <xLgr>" & m_dest.m_enderDest.m_xLgr & "</xLgr>")
            txt_xml.Append Trim("                   <nro>" & m_dest.m_enderDest.m_nro & "</nro>")
            txt_xml.Append Trim("                   <xBairro>" & m_dest.m_enderDest.m_xBairro & "</xBairro>")
            txt_xml.Append Trim("                   <cMun>" & m_dest.m_enderDest.m_cMun & "</cMun>")
            txt_xml.Append Trim("                   <xMun>" & m_dest.m_enderDest.m_xMun & "</xMun>")
            txt_xml.Append Trim("                   <UF>" & m_dest.m_enderDest.m_UF & "</UF>")
            txt_xml.Append Trim("                   <CEP>" & m_dest.m_enderDest.m_CEP & "</CEP>")
            txt_xml.Append Trim("                   <cPais>" & m_dest.m_enderDest.m_cPais & "</cPais>")
            If Len(m_dest.m_enderDest.m_fone) Then txt_xml.Append Trim("                   <fone>" & m_dest.m_enderDest.m_fone & "</fone>")
            txt_xml.Append Trim("               </enderDest>")
            txt_xml.Append Trim("               <indIEDest>" & m_dest.m_indIEDest & "</indIEDest>")
            If Len(m_dest.m_email) Then txt_xml.Append Trim("               <email>" & m_dest.m_email & "</email>")
            If Not m_dest.m_IE = "" Then txt_xml.Append Trim("               <IE>" & m_dest.m_IE & "</IE>")
        End If
        txt_xml.Append Trim("           </dest>")
    End If
End Sub

Private Sub GetNota_BuildInfIntermed(txt_xml As hStringBuilder)
    If Len(m_infIntermed.m_CNPJ) > 0 Then
        txt_xml.Append Trim("           <infIntermed>")
        txt_xml.Append Trim("               <CNPJ>" & m_infIntermed.m_CNPJ & "</CNPJ>")
        If Len(m_infIntermed.m_idCadIntTransp) > 0 Then txt_xml.Append Trim("               <idCadIntTransp>" & m_infIntermed.m_idCadIntTransp & "</idCadIntTransp>")
        txt_xml.Append Trim("           </infIntermed>")
    End If
End Sub

Private Sub GetNota_BuildDet(txt_xml As hStringBuilder)
    Dim cc As Long
    If m_Ambiente = 12 Then
        m_cItens = m_cItens + 1
        ReDim Preserve m_Itens(m_cItens)
        For cc = m_cItens To 1 Step -1
            Set m_Itens(cc) = m_Itens(cc - 1)
        Next
        With m_Itens(0)
            .prod.m_cProd = "1"
            .prod.m_xProd = "NOTA FISCAL EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
            .prod.m_uCom = 1
            .prod.m_qCom = 1
            .prod.m_indTot = 1
            .prod.m_vProd = 1
            .prod.m_vUnCom = 1
            .prod.m_vUnTrib = 1
            .imposto.m_vTotTrib = 1
        End With
    End If
    For cc = 0 To m_cItens
        txt_xml.Append Trim("           <det nItem=""" & CStr(cc + 1) & """>")
        txt_xml.Append Trim("               <prod>")
        With m_Itens(cc).prod
            txt_xml.Append Trim("                   <cProd>" & .m_cProd & "</cProd>")
            txt_xml.Append Trim("                   <cEAN>" & IIf(Len(.m_cEAN) > 0, .m_cEAN, "SEM GTIN") & "</cEAN>")
            txt_xml.Append Trim("                   <xProd>" & (.m_xProd) & "</xProd>")
            txt_xml.Append Trim("                   <NCM>" & .m_NCM & "</NCM>")
            If Len(.m_CEST) > 0 Then txt_xml.Append Trim("                   <CEST>" & .m_CEST & "</CEST>")
            txt_xml.Append Trim("                   <CFOP>" & .m_CFOP & "</CFOP>")
            txt_xml.Append Trim("                   <uCom>" & .m_uCom & "</uCom>")
            txt_xml.Append Trim("                   <qCom>" & .m_qCom & "</qCom>")
            txt_xml.Append Trim("                   <vUnCom>" & cFormatMoeda_NF_1(.m_vUnCom) & "</vUnCom>")
            txt_xml.Append Trim("                   <vProd>" & cFormatMoeda_NF_1(.m_vProd) & "</vProd>")
            txt_xml.Append Trim("                   <cEANTrib>" & IIf(Len(.m_cEANTrib) > 0, .m_cEANTrib, "SEM GTIN") & "</cEANTrib>")
            txt_xml.Append Trim("                   <uTrib>" & .m_uTrib & "</uTrib>")
            txt_xml.Append Trim("                   <qTrib>" & .m_qTrib & "</qTrib>")
            txt_xml.Append Trim("                   <vUnTrib>" & cFormatMoeda_NF_1(.m_vUnTrib) & "</vUnTrib>")
            If .m_vFrete > 0 Then txt_xml.Append Trim("                   <vFrete>" & cFormatMoeda_NF_1(.m_vFrete) & "</vFrete>")
            txt_xml.Append Trim("                   <indTot>" & .m_indTot & "</indTot>")
            If .DiCount > 0 Then
                Dim cc_di As Long, cc_adi As Long
                For cc_di = 0 To .DiCount - 1
                    With .Di(cc_di)
                        txt_xml.Append Trim("                   <DI>")
                        txt_xml.Append Trim("                       <nDI>" & .nDI & "</nDI>")
                        If Len(.dDI) > 0 Then txt_xml.Append Trim("                       <dDI>" & .dDI & "</dDI>")
                        If Len(.xLocDesemb) > 0 Then txt_xml.Append Trim("                       <xLocDesemb>" & .xLocDesemb & "</xLocDesemb>")
                        If Len(.UFDesemb) > 0 Then txt_xml.Append Trim("                       <UFDesemb>" & .UFDesemb & "</UFDesemb>")
                        If Len(.dDesemb) > 0 Then txt_xml.Append Trim("                       <dDesemb>" & .dDesemb & "</dDesemb>")
                        If .tpViaTransp > 0 Then txt_xml.Append Trim("                       <tpViaTransp>" & .tpViaTransp & "</tpViaTransp>")
                        If .vAFRMM > 0 Then txt_xml.Append Trim("                       <vAFRMM>" & cFormatMoeda_NF_1(.vAFRMM) & "</vAFRMM>")
                        If .tpIntermedio > 0 Then txt_xml.Append Trim("                       <tpIntermedio>" & .tpIntermedio & "</tpIntermedio>")
                        If Len(.cnpj) > 0 Then txt_xml.Append Trim("                       <CNPJ>" & .cnpj & "</CNPJ>")
                        If Len(.UFTerceiro) > 0 Then txt_xml.Append Trim("                       <UFTerceiro>" & .UFTerceiro & "</UFTerceiro>")
                        If Len(.cExportador) > 0 Then txt_xml.Append Trim("                       <cExportador>" & .cExportador & "</cExportador>")
                        If .AdiCount > 0 Then
                            For cc_adi = 0 To .AdiCount - 1
                                With .Adi(cc_adi)
                                    txt_xml.Append Trim("                       <ADI>")
                                    If Len(.nAdicao) > 0 Then txt_xml.Append Trim("                           <nAdicao>" & .nAdicao & "</nAdicao>")
                                    If .nSeqAdic > 0 Then txt_xml.Append Trim("                           <nSeqAdic>" & .nSeqAdic & "</nSeqAdic>")
                                    If Len(.xFabricante) > 0 Then txt_xml.Append Trim("                           <xFabricante>" & .xFabricante & "</xFabricante>")
                                    If .vDescDI > 0 Then txt_xml.Append Trim("                           <vDescDI>" & cFormatMoeda_NF_1(.vDescDI) & "</vDescDI>")
                                    If Len(.nDraw) > 0 Then txt_xml.Append Trim("                           <nDraw>" & .nDraw & "</nDraw>")
                                    txt_xml.Append Trim("                       </ADI>")
                                End With
                            Next
                        End If
                        txt_xml.Append Trim("                   </DI>")
                    End With
                Next
            End If
            With m_Itens(cc).prod.comb
                If Len(.cProdANP) > 0 Then
                    txt_xml.Append Trim("                   <comb>")
                    txt_xml.Append Trim("                       <cProdANP>" & .cProdANP & "</cProdANP>")
                    If Len(.descANP) > 0 Then txt_xml.Append Trim("                       <descANP>" & .descANP & "</descANP>")
                    If .pGLP > 0 Then txt_xml.Append Trim("                       <pGLP>" & .pGLP & "</pGLP>")
                    If .pMist > 0 Then txt_xml.Append Trim("                       <pMist>" & .pMist & "</pMist>")
                    If Len(.CODIF) > 0 Then txt_xml.Append Trim("                       <CODIF>" & .CODIF & "</CODIF>")
                    If .qTemp > 0 Then txt_xml.Append Trim("                       <qTemp>" & .qTemp & "</qTemp>")
                    If Len(.UFComb) > 0 Then txt_xml.Append Trim("                       <UFComb>" & .UFComb & "</UFComb>")
                    With .CIDE
                        If .qBCprod > 0 Or .vAliqProd > 0 Or .vCIDE > 0 Then
                            txt_xml.Append Trim("                       <CIDE>")
                            If .qBCprod > 0 Then txt_xml.Append Trim("                           <qBCprod>" & .qBCprod & "</qBCprod>")
                            If .vAliqProd > 0 Then txt_xml.Append Trim("                           <vAliqProd>" & cFormatMoeda_NF_1(.vAliqProd) & "</vAliqProd>")
                            If .vCIDE > 0 Then txt_xml.Append Trim("                           <vCIDE>" & cFormatMoeda_NF_1(.vCIDE) & "</vCIDE>")
                            txt_xml.Append Trim("                       </CIDE>")
                        End If
                    End With
                    With .encerrante
                        If Len(.nBico) > 0 Or Len(.nTanque) > 0 Or Len(.nBomba) > 0 Or Len(.nMotor) > 0 Or Len(.nEncerrante) > 0 Then
                            txt_xml.Append Trim("                       <encerrante>")
                            If Len(.nBico) > 0 Then txt_xml.Append Trim("                           <nBico>" & .nBico & "</nBico>")
                            If Len(.nTanque) > 0 Then txt_xml.Append Trim("                           <nTanque>" & .nTanque & "</nTanque>")
                            If Len(.nBomba) > 0 Then txt_xml.Append Trim("                           <nBomba>" & .nBomba & "</nBomba>")
                            If Len(.nMotor) > 0 Then txt_xml.Append Trim("                           <nMotor>" & .nMotor & "</nMotor>")
                            If Len(.nEncerrante) > 0 Then txt_xml.Append Trim("                           <nEncerrante>" & .nEncerrante & "</nEncerrante>")
                            If Len(.dhInicio) > 0 Then txt_xml.Append Trim("                           <dhInicio>" & .dhInicio & "</dhInicio>")
                            If Len(.dhFim) > 0 Then txt_xml.Append Trim("                           <dhFim>" & .dhFim & "</dhFim>")
                            txt_xml.Append Trim("                       </encerrante>")
                        End If
                    End With
                    txt_xml.Append Trim("                   </comb>")
                End If
            End With
            If .RastroCount > 0 Then
                Dim cc_rastro As Long
                For cc_rastro = 0 To .RastroCount - 1
                    With .Rastro(cc_rastro)
                        txt_xml.Append Trim("                   <rastro>")
                        If Len(.nLote) > 0 Then txt_xml.Append Trim("                       <nLote>" & .nLote & "</nLote>")
                        If .qLote > 0 Then txt_xml.Append Trim("                       <qLote>" & .qLote & "</qLote>")
                        If Len(.dFab) > 0 Then txt_xml.Append Trim("                       <dFab>" & .dFab & "</dFab>")
                        If Len(.dVal) > 0 Then txt_xml.Append Trim("                       <dVal>" & .dVal & "</dVal>")
                        If Len(.cAgreg) > 0 Then txt_xml.Append Trim("                       <cAgreg>" & .cAgreg & "</cAgreg>")
                        txt_xml.Append Trim("                   </rastro>")
                    End With
                Next
            End If
        End With
        txt_xml.Append Trim("               </prod>")
        GetNota_BuildDetImposto txt_xml, cc
        txt_xml.Append Trim("           </det>")
    Next
End Sub

Private Sub GetNota_BuildDetImposto(txt_xml As hStringBuilder, ByVal cc As Long)
    With m_Itens(cc).imposto
        txt_xml.Append Trim("               <imposto>")
        txt_xml.Append Trim("                   <vTotTrib>" & cFormatMoeda_NF_1(.m_vTotTrib) & "</vTotTrib>")
        txt_xml.Append Trim("                   <ICMS>")
        Dim m_icms_cst As String
        m_icms_cst = .m_ICMS_ICMS00_CST
        If m_emit.m_CRT = 1 Then
            Select Case m_icms_cst
            Case "101"
                txt_xml.Append Trim("                       <ICMSSN101>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CSOSN>101</CSOSN>")
                txt_xml.Append Trim("                           <pCredSN>" & cFormatMoeda_NF_1(.m_CSOSN_pCredSN) & "</pCredSN>")
                txt_xml.Append Trim("                           <vCredICMS>" & cFormatMoeda_NF_1(.m_CSOSN_vCredICMS) & "</vCredICMS>")
                txt_xml.Append Trim("                       </ICMSSN101>")
            Case "201"
                txt_xml.Append Trim("                       <ICMSSN201>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CSOSN>201</CSOSN>")
                txt_xml.Append Trim("                           <modBCST>4</modBCST>")
                txt_xml.Append Trim("                           <pST>" & cFormatMoeda_NF_1(.m_ICMS_pST) & "</pST>")
                txt_xml.Append Trim("                           <vBCST>" & cFormatMoeda_NF_1(.m_ICMS_vBCST) & "</vBCST>")
                txt_xml.Append Trim("                           <vICMSST>" & cFormatMoeda_NF_1(.m_ICMS_vICMSST) & "</vICMSST>")
                txt_xml.Append Trim("                       </ICMSSN201>")
            Case "202"
                txt_xml.Append Trim("                       <ICMSSN202>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CSOSN>202</CSOSN>")
                txt_xml.Append Trim("                           <modBCST>4</modBCST>")
                txt_xml.Append Trim("                           <pST>" & cFormatMoeda_NF_1(.m_ICMS_pST) & "</pST>")
                txt_xml.Append Trim("                           <vBCST>" & cFormatMoeda_NF_1(.m_ICMS_vBCST) & "</vBCST>")
                txt_xml.Append Trim("                           <vICMSST>" & cFormatMoeda_NF_1(.m_ICMS_vICMSST) & "</vICMSST>")
                txt_xml.Append Trim("                       </ICMSSN202>")
            Case "203"
                txt_xml.Append Trim("                       <ICMSSN203>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CSOSN>203</CSOSN>")
                txt_xml.Append Trim("                           <modBCST>4</modBCST>")
                txt_xml.Append Trim("                           <pST>" & cFormatMoeda_NF_1(.m_ICMS_pST) & "</pST>")
                txt_xml.Append Trim("                           <vBCST>" & cFormatMoeda_NF_1(.m_ICMS_vBCST) & "</vBCST>")
                txt_xml.Append Trim("                           <vICMSST>" & cFormatMoeda_NF_1(.m_ICMS_vICMSST) & "</vICMSST>")
                txt_xml.Append Trim("                       </ICMSSN203>")
            Case "500"
                txt_xml.Append Trim("                       <ICMSSN500>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CSOSN>500</CSOSN>")
                txt_xml.Append Trim("                           <vBCSTRet>" & cFormatMoeda_NF_1(.m_ICMS_vBCSTRet) & "</vBCSTRet>")
                txt_xml.Append Trim("                           <pST>" & cFormatMoeda_NF_1(.m_ICMS_pSTRet) & "</pST>")
                txt_xml.Append Trim("                           <vICMSSTRet>" & cFormatMoeda_NF_1(.m_ICMS_vICMSSTRet) & "</vICMSSTRet>")
                txt_xml.Append Trim("                       </ICMSSN500>")
            Case "900"
                txt_xml.Append Trim("                       <ICMSSN900>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CSOSN>900</CSOSN>")
                txt_xml.Append Trim("                           <modBC>" & .m_ICMS_ICMS00_modBC & "</modBC>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_pICMS) & "</pICMS>")
                txt_xml.Append Trim("                           <vICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vICMS) & "</vICMS>")
                txt_xml.Append Trim("                           <modBCST>4</modBCST>")
                txt_xml.Append Trim("                           <pST>" & cFormatMoeda_NF_1(.m_ICMS_pST) & "</pST>")
                txt_xml.Append Trim("                           <vBCST>" & cFormatMoeda_NF_1(.m_ICMS_vBCST) & "</vBCST>")
                txt_xml.Append Trim("                           <vICMSST>" & cFormatMoeda_NF_1(.m_ICMS_vICMSST) & "</vICMSST>")
                txt_xml.Append Trim("                       </ICMSSN900>")
            Case Else
                If m_icms_cst = "103" Then
                    txt_xml.Append Trim("                       <ICMSSN103>")
                ElseIf m_icms_cst = "300" Then
                    txt_xml.Append Trim("                       <ICMSSN300>")
                ElseIf m_icms_cst = "400" Then
                    txt_xml.Append Trim("                       <ICMSSN400>")
                Else
                    txt_xml.Append Trim("                       <ICMSSN102>")
                End If
                    
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CSOSN>" & m_icms_cst & "</CSOSN>")
                txt_xml.Append Trim("                       </" & IIf(m_icms_cst = "103", "ICMSSN103", IIf(m_icms_cst = "300", "ICMSSN300", IIf(m_icms_cst = "400", "ICMSSN400", "ICMSSN102"))) & ">")
            End Select
        Else
            Select Case m_icms_cst
            Case "10"
                txt_xml.Append Trim("                       <ICMS10>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>10</CST>")
                txt_xml.Append Trim("                           <modBC>" & .m_ICMS_ICMS00_modBC & "</modBC>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_pICMS) & "</pICMS>")
                txt_xml.Append Trim("                           <vICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vICMS) & "</vICMS>")
                txt_xml.Append Trim("                           <modBCST>4</modBCST>")
                txt_xml.Append Trim("                           <pST>" & cFormatMoeda_NF_1(.m_ICMS_pST) & "</pST>")
                txt_xml.Append Trim("                           <vBCST>" & cFormatMoeda_NF_1(.m_ICMS_vBCST) & "</vBCST>")
                txt_xml.Append Trim("                           <vICMSST>" & cFormatMoeda_NF_1(.m_ICMS_vICMSST) & "</vICMSST>")
                txt_xml.Append Trim("                       </ICMS10>")
            Case "20"
                txt_xml.Append Trim("                       <ICMS20>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>20</CST>")
                txt_xml.Append Trim("                           <modBC>" & .m_ICMS_ICMS00_modBC & "</modBC>")
                txt_xml.Append Trim("                           <pRedBC>" & cFormatMoeda_NF_1(.m_ICMS_pRedBC) & "</pRedBC>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_pICMS) & "</pICMS>")
                txt_xml.Append Trim("                           <vICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vICMS) & "</vICMS>")
                txt_xml.Append Trim("                       </ICMS20>")
            Case "30"
                txt_xml.Append Trim("                       <ICMS30>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>30</CST>")
                txt_xml.Append Trim("                           <modBCST>4</modBCST>")
                txt_xml.Append Trim("                           <pST>" & cFormatMoeda_NF_1(.m_ICMS_pST) & "</pST>")
                txt_xml.Append Trim("                           <vBCST>" & cFormatMoeda_NF_1(.m_ICMS_vBCST) & "</vBCST>")
                txt_xml.Append Trim("                           <vICMSST>" & cFormatMoeda_NF_1(.m_ICMS_vICMSST) & "</vICMSST>")
                txt_xml.Append Trim("                       </ICMS30>")
            Case "40"
                txt_xml.Append Trim("                       <ICMS40>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>40</CST>")
                txt_xml.Append Trim("                           <vICMSDeson>" & cFormatMoeda_NF_1(.m_ICMS_vICMSDeson) & "</vICMSDeson>")
                If Len(.m_ICMS_motDesICMS) > 0 Then txt_xml.Append Trim("                           <motDesICMS>" & .m_ICMS_motDesICMS & "</motDesICMS>")
                txt_xml.Append Trim("                       </ICMS40>")
            Case "41"
                txt_xml.Append Trim("                       <ICMS41>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>41</CST>")
                txt_xml.Append Trim("                       </ICMS41>")
            Case "50"
                txt_xml.Append Trim("                       <ICMS50>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>50</CST>")
                txt_xml.Append Trim("                       </ICMS50>")
            Case "51"
                txt_xml.Append Trim("                       <ICMS51>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>51</CST>")
                txt_xml.Append Trim("                           <modBC>" & .m_ICMS_ICMS00_modBC & "</modBC>")
                txt_xml.Append Trim("                           <pRedBC>" & cFormatMoeda_NF_1(.m_ICMS_pRedBC) & "</pRedBC>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_pICMS) & "</pICMS>")
                txt_xml.Append Trim("                           <vICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vICMS) & "</vICMS>")
                txt_xml.Append Trim("                       </ICMS51>")
            Case "60"
                txt_xml.Append Trim("                       <ICMS60>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>60</CST>")
                txt_xml.Append Trim("                           <vBCST>" & cFormatMoeda_NF_1(.m_ICMS_vBCST) & "</vBCST>")
                txt_xml.Append Trim("                           <vICMSST>" & cFormatMoeda_NF_1(.m_ICMS_vICMSST) & "</vICMSST>")
                txt_xml.Append Trim("                       </ICMS60>")
            Case "70"
                txt_xml.Append Trim("                       <ICMS70>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>70</CST>")
                txt_xml.Append Trim("                           <modBC>" & .m_ICMS_ICMS00_modBC & "</modBC>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_pICMS) & "</pICMS>")
                txt_xml.Append Trim("                           <vICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vICMS) & "</vICMS>")
                txt_xml.Append Trim("                           <modBCST>4</modBCST>")
                txt_xml.Append Trim("                           <pST>" & cFormatMoeda_NF_1(.m_ICMS_pST) & "</pST>")
                txt_xml.Append Trim("                           <vBCST>" & cFormatMoeda_NF_1(.m_ICMS_vBCST) & "</vBCST>")
                txt_xml.Append Trim("                           <vICMSST>" & cFormatMoeda_NF_1(.m_ICMS_vICMSST) & "</vICMSST>")
                txt_xml.Append Trim("                       </ICMS70>")
            Case "90"
                txt_xml.Append Trim("                       <ICMS90>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>90</CST>")
                txt_xml.Append Trim("                           <modBC>" & .m_ICMS_ICMS00_modBC & "</modBC>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_pICMS) & "</pICMS>")
                txt_xml.Append Trim("                           <vICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vICMS) & "</vICMS>")
                txt_xml.Append Trim("                           <modBCST>4</modBCST>")
                txt_xml.Append Trim("                           <pST>" & cFormatMoeda_NF_1(.m_ICMS_pST) & "</pST>")
                txt_xml.Append Trim("                           <vBCST>" & cFormatMoeda_NF_1(.m_ICMS_vBCST) & "</vBCST>")
                txt_xml.Append Trim("                           <vICMSST>" & cFormatMoeda_NF_1(.m_ICMS_vICMSST) & "</vICMSST>")
                txt_xml.Append Trim("                       </ICMS90>")
            Case Else
                txt_xml.Append Trim("                       <ICMS00>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>" & m_icms_cst & "</CST>")
                txt_xml.Append Trim("                           <modBC>" & .m_ICMS_ICMS00_modBC & "</modBC>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_pICMS) & "</pICMS>")
                txt_xml.Append Trim("                           <vICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vICMS) & "</vICMS>")
                txt_xml.Append Trim("                       </ICMS00>")
            End Select
        End If
        txt_xml.Append Trim("                   </ICMS>")
        If m_emit.m_CRT <> 1 Then
            txt_xml.Append Trim("                   <IPI>")
            Dim m_IPI_CST As String
            m_IPI_CST = .m_IPI_IPINT_CST
            If Len(m_IPI_CST) = 0 Then m_IPI_CST = .m_IPI_CST
            If m_IPI_CST = "00" Or m_IPI_CST = "49" Or m_IPI_CST = "50" Or m_IPI_CST = "99" Then
                txt_xml.Append Trim("                       <cEnq>" & .m_IPI_cEnq & "</cEnq>")
                txt_xml.Append Trim("                       <IPITrib>")
                txt_xml.Append Trim("                           <CST>" & m_IPI_CST & "</CST>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_IPI_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pIPI>" & cFormatMoeda_NF_1(.m_IPI_pIPI) & "</pIPI>")
                txt_xml.Append Trim("                           <vIPI>" & cFormatMoeda_NF_1(.m_IPI_vIPI) & "</vIPI>")
                txt_xml.Append Trim("                       </IPITrib>")
            Else
                txt_xml.Append Trim("                       <cEnq>" & .m_IPI_cEnq & "</cEnq>")
                txt_xml.Append Trim("                       <IPINT>")
                txt_xml.Append Trim("                           <CST>" & m_IPI_CST & "</CST>")
                txt_xml.Append Trim("                       </IPINT>")
            End If
            txt_xml.Append Trim("                   </IPI>")
        End If
        If .m_II_vBC <> 0 Or .m_II_vII <> 0 Then
            txt_xml.Append Trim("                   <II>")
            txt_xml.Append Trim("                       <vBC>" & cFormatMoeda_NF_1(.m_II_vBC) & "</vBC>")
            txt_xml.Append Trim("                       <vDespAdu>" & cFormatMoeda_NF_1(.m_II_vDespAdu) & "</vDespAdu>")
            txt_xml.Append Trim("                       <vII>" & cFormatMoeda_NF_1(.m_II_vII) & "</vII>")
            txt_xml.Append Trim("                       <vIOF>" & cFormatMoeda_NF_1(.m_II_vIOF) & "</vIOF>")
            txt_xml.Append Trim("                   </II>")
        End If
        txt_xml.Append Trim("                   <PIS>")
        Dim m_PIS_CST As String
        m_PIS_CST = .m_PIS_CST
        If Len(m_PIS_CST) = 0 Then m_PIS_CST = .m_PIS_PISAliq_CST
        Select Case m_PIS_CST
        Case "01"
            txt_xml.Append Trim("                       <PISAliq>")
            txt_xml.Append Trim("                           <CST>01</CST>")
            txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_PIS_vBC) & "</vBC>")
            txt_xml.Append Trim("                           <pPIS>" & cFormatMoeda_NF_1(.m_PIS_pPIS) & "</pPIS>")
            txt_xml.Append Trim("                           <vPIS>" & cFormatMoeda_NF_1(.m_PIS_vPIS) & "</vPIS>")
            txt_xml.Append Trim("                       </PISAliq>")
        Case "02", "03"
            txt_xml.Append Trim("                       <PISQtde>")
            txt_xml.Append Trim("                           <CST>" & m_PIS_CST & "</CST>")
            txt_xml.Append Trim("                           <qBCProd>" & cFormatMoeda_NF_1(.m_PIS_qBCProd) & "</qBCProd>")
            txt_xml.Append Trim("                           <vAliqProd>" & cFormatMoeda_NF_1(.m_PIS_vAliqProd) & "</vAliqProd>")
            txt_xml.Append Trim("                           <vPIS>" & cFormatMoeda_NF_1(.m_PIS_vPIS) & "</vPIS>")
            txt_xml.Append Trim("                       </PISQtde>")
        Case "04", "05"
            txt_xml.Append Trim("                       <PISOutr>")
            txt_xml.Append Trim("                           <CST>" & m_PIS_CST & "</CST>")
            txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_PIS_vBC) & "</vBC>")
            txt_xml.Append Trim("                           <pPIS>" & cFormatMoeda_NF_1(.m_PIS_pPIS) & "</pPIS>")
            txt_xml.Append Trim("                           <vPIS>" & cFormatMoeda_NF_1(.m_PIS_vPIS) & "</vPIS>")
            txt_xml.Append Trim("                       </PISOutr>")
        Case "06", "07", "08", "09"
            txt_xml.Append Trim("                       <PISNT>")
            txt_xml.Append Trim("                           <CST>" & m_PIS_CST & "</CST>")
            txt_xml.Append Trim("                       </PISNT>")
        Case Else
            txt_xml.Append Trim("                       <PISOutr>")
            txt_xml.Append Trim("                           <CST>" & m_PIS_CST & "</CST>")
            txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_PIS_vBC) & "</vBC>")
            txt_xml.Append Trim("                           <pPIS>" & cFormatMoeda_NF_1(.m_PIS_pPIS) & "</pPIS>")
            txt_xml.Append Trim("                           <vPIS>" & cFormatMoeda_NF_1(.m_PIS_vPIS) & "</vPIS>")
            txt_xml.Append Trim("                       </PISOutr>")
        End Select
        txt_xml.Append Trim("                   </PIS>")
        If .m_PISST_vPIS <> 0 Then
            txt_xml.Append Trim("                   <PISST>")
            If .m_PISST_qBCProd <> 0 And .m_PISST_vAliqProd <> 0 Then
                txt_xml.Append Trim("                       <qBCProd>" & cFormatMoeda_NF_1(.m_PISST_qBCProd) & "</qBCProd>")
                txt_xml.Append Trim("                       <vAliqProd>" & cFormatMoeda_NF_1(.m_PISST_vAliqProd) & "</vAliqProd>")
            Else
                If .m_PISST_vBC <> 0 Then txt_xml.Append Trim("                       <vBC>" & cFormatMoeda_NF_1(.m_PISST_vBC) & "</vBC>")
                If .m_PISST_pPIS <> 0 Then txt_xml.Append Trim("                       <pPIS>" & cFormatMoeda_NF_1(.m_PISST_pPIS) & "</pPIS>")
            End If
            txt_xml.Append Trim("                       <vPIS>" & cFormatMoeda_NF_1(.m_PISST_vPIS) & "</vPIS>")
            If Len(.m_PISST_indSoma) > 0 Then txt_xml.Append Trim("                       <indSomaPISST>" & .m_PISST_indSoma & "</indSomaPISST>")
            txt_xml.Append Trim("                   </PISST>")
        End If
        txt_xml.Append Trim("                   <COFINS>")
        Dim m_COFINS_CST As String
        m_COFINS_CST = .m_COFINS_CST
        If Len(m_COFINS_CST) = 0 Then m_COFINS_CST = .m_COFINS_COFINSAliq_CST
        Select Case m_COFINS_CST
        Case "01"
            txt_xml.Append Trim("                       <COFINSAliq>")
            txt_xml.Append Trim("                           <CST>01</CST>")
            txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_COFINS_vBC) & "</vBC>")
            txt_xml.Append Trim("                           <pCOFINS>" & cFormatMoeda_NF_1(.m_COFINS_pCOFINS) & "</pCOFINS>")
            txt_xml.Append Trim("                           <vCOFINS>" & cFormatMoeda_NF_1(.m_COFINS_vCOFINS) & "</vCOFINS>")
            txt_xml.Append Trim("                       </COFINSAliq>")
        Case "02", "03"
            txt_xml.Append Trim("                       <COFINSQtde>")
            txt_xml.Append Trim("                           <CST>" & m_COFINS_CST & "</CST>")
            txt_xml.Append Trim("                           <qBCProd>" & cFormatMoeda_NF_1(.m_COFINS_qBCProd) & "</qBCProd>")
            txt_xml.Append Trim("                           <vAliqProd>" & cFormatMoeda_NF_1(.m_COFINS_vAliqProd) & "</vAliqProd>")
            txt_xml.Append Trim("                           <vCOFINS>" & cFormatMoeda_NF_1(.m_COFINS_vCOFINS) & "</vCOFINS>")
            txt_xml.Append Trim("                       </COFINSQtde>")
        Case "04", "05"
            txt_xml.Append Trim("                       <COFINSOutr>")
            txt_xml.Append Trim("                           <CST>" & m_COFINS_CST & "</CST>")
            txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_COFINS_vBC) & "</vBC>")
            txt_xml.Append Trim("                           <pCOFINS>" & cFormatMoeda_NF_1(.m_COFINS_pCOFINS) & "</pCOFINS>")
            txt_xml.Append Trim("                           <vCOFINS>" & cFormatMoeda_NF_1(.m_COFINS_vCOFINS) & "</vCOFINS>")
            txt_xml.Append Trim("                       </COFINSOutr>")
        Case "06", "07", "08", "09"
            txt_xml.Append Trim("                       <COFINSNT>")
            txt_xml.Append Trim("                           <CST>" & m_COFINS_CST & "</CST>")
            txt_xml.Append Trim("                       </COFINSNT>")
        Case Else
            txt_xml.Append Trim("                       <COFINSOutr>")
            txt_xml.Append Trim("                           <CST>" & m_COFINS_CST & "</CST>")
            txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_COFINS_vBC) & "</vBC>")
            txt_xml.Append Trim("                           <pCOFINS>" & cFormatMoeda_NF_1(.m_COFINS_pCOFINS) & "</pCOFINS>")
            txt_xml.Append Trim("                           <vCOFINS>" & cFormatMoeda_NF_1(.m_COFINS_vCOFINS) & "</vCOFINS>")
            txt_xml.Append Trim("                       </COFINSOutr>")
        End Select
        txt_xml.Append Trim("                   </COFINS>")
        If .m_COFINSST_vCOFINS <> 0 Then
            txt_xml.Append Trim("                   <COFINSST>")
            If .m_COFINSST_qBCProd <> 0 And .m_COFINSST_vAliqProd <> 0 Then
                txt_xml.Append Trim("                       <qBCProd>" & cFormatMoeda_NF_1(.m_COFINSST_qBCProd) & "</qBCProd>")
                txt_xml.Append Trim("                       <vAliqProd>" & cFormatMoeda_NF_1(.m_COFINSST_vAliqProd) & "</vAliqProd>")
            Else
                If .m_COFINSST_vBC <> 0 Then txt_xml.Append Trim("                       <vBC>" & cFormatMoeda_NF_1(.m_COFINSST_vBC) & "</vBC>")
                If .m_COFINSST_pCOFINS <> 0 Then txt_xml.Append Trim("                       <pCOFINS>" & cFormatMoeda_NF_1(.m_COFINSST_pCOFINS) & "</pCOFINS>")
            End If
            txt_xml.Append Trim("                       <vCOFINS>" & cFormatMoeda_NF_1(.m_COFINSST_vCOFINS) & "</vCOFINS>")
            If Len(.m_COFINSST_indSoma) > 0 Then txt_xml.Append Trim("                       <indSomaCOFINSST>" & .m_COFINSST_indSoma & "</indSomaCOFINSST>")
            txt_xml.Append Trim("                   </COFINSST>")
        End If
        If .m_ICMSUFDest_vBCUFDest <> 0 Or .m_ICMSUFDest_vICMSUFDest <> 0 Then
            txt_xml.Append Trim("                   <ICMSUFDest>")
            txt_xml.Append Trim("                       <vBCUFDest>" & cFormatMoeda_NF_1(.m_ICMSUFDest_vBCUFDest) & "</vBCUFDest>")
            txt_xml.Append Trim("                       <pFCPUFDest>" & cFormatMoeda_NF_1(.m_ICMSUFDest_pFCPUFDest) & "</pFCPUFDest>")
            txt_xml.Append Trim("                       <pICMSUFDest>" & cFormatMoeda_NF_1(.m_ICMSUFDest_pICMSUFDest) & "</pICMSUFDest>")
            txt_xml.Append Trim("                       <pICMSInter>" & cFormatMoeda_NF_1(.m_ICMSUFDest_pICMSInter) & "</pICMSInter>")
            txt_xml.Append Trim("                       <pICMSInterPart>" & cFormatMoeda_NF_1(.m_ICMSUFDest_pICMSInterPart) & "</pICMSInterPart>")
            txt_xml.Append Trim("                       <vFCPUFDest>" & cFormatMoeda_NF_1(.m_ICMSUFDest_vFCPUFDest) & "</vFCPUFDest>")
            txt_xml.Append Trim("                       <vICMSUFDest>" & cFormatMoeda_NF_1(.m_ICMSUFDest_vICMSUFDest) & "</vICMSUFDest>")
            txt_xml.Append Trim("                       <vICMSUFRemet>" & cFormatMoeda_NF_1(.m_ICMSUFDest_vICMSUFRemet) & "</vICMSUFRemet>")
            txt_xml.Append Trim("                   </ICMSUFDest>")
        End If
        If Len(.m_IS_CST) > 0 Then
            txt_xml.Append Trim("                   <IS>")
            txt_xml.Append Trim("                       <CSTIS>" & .m_IS_CST & "</CSTIS>")
            If Len(.m_IS_cClassTrib) > 0 Then txt_xml.Append Trim("                       <cClassTribIS>" & .m_IS_cClassTrib & "</cClassTribIS>")
            If .m_IS_vBC <> 0 Then txt_xml.Append Trim("                       <vBCIS>" & cFormatMoeda_NF_1(.m_IS_vBC) & "</vBCIS>")
            If .m_IS_pIS <> 0 Then txt_xml.Append Trim("                       <pIS>" & cFormatMoeda_NF_1(.m_IS_pIS) & "</pIS>")
            If .m_IS_pISEspec <> 0 Then txt_xml.Append Trim("                       <pISEspec>" & cFormatMoeda_NF_1(.m_IS_pISEspec) & "</pISEspec>")
            If Len(.m_IS_uTrib) > 0 Then txt_xml.Append Trim("                       <uTrib>" & .m_IS_uTrib & "</uTrib>")
            If .m_IS_qTrib <> 0 Then txt_xml.Append Trim("                       <qTrib>" & cFormatMoeda_NF_1(.m_IS_qTrib) & "</qTrib>")
            If .m_IS_vIS <> 0 Then txt_xml.Append Trim("                       <vIS>" & cFormatMoeda_NF_1(.m_IS_vIS) & "</vIS>")
            txt_xml.Append Trim("                   </IS>")
        End If
        If Len(.m_IBSCBS_CST) > 0 Then
            txt_xml.Append Trim("                   <IBSCBS>")
            txt_xml.Append Trim("                       <CST>" & .m_IBSCBS_CST & "</CST>")
            If Len(.m_IBSCBS_cClassTrib) > 0 Then txt_xml.Append Trim("                       <cClassTrib>" & .m_IBSCBS_cClassTrib & "</cClassTrib>")
            If Len(.m_IBSCBS_indDoacao) > 0 Then txt_xml.Append Trim("                       <indDoacao>" & .m_IBSCBS_indDoacao & "</indDoacao>")
            If .m_IBSCBS_vBC <> 0 Then
                txt_xml.Append Trim("                       <gIBSCBS>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_IBSCBS_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <gIBSUF>")
                txt_xml.Append Trim("                               <pIBSUF>" & cFormatMoeda_NF_1(.m_IBSCBS_pIBSUF) & "</pIBSUF>")
                If .m_IBSCBS_gIBSUF_pDif <> 0 Or .m_IBSCBS_gIBSUF_vDif <> 0 Then
                    txt_xml.Append Trim("                               <gDif>")
                    If .m_IBSCBS_gIBSUF_pDif <> 0 Then txt_xml.Append Trim("                                   <pDif>" & cFormatMoeda_NF_1(.m_IBSCBS_gIBSUF_pDif) & "</pDif>")
                    If .m_IBSCBS_gIBSUF_vDif <> 0 Then txt_xml.Append Trim("                                   <vDif>" & cFormatMoeda_NF_1(.m_IBSCBS_gIBSUF_vDif) & "</vDif>")
                    txt_xml.Append Trim("                               </gDif>")
                End If
                If .m_IBSCBS_gIBSUF_vDevTrib <> 0 Then
                    txt_xml.Append Trim("                               <gDevTrib>")
                    txt_xml.Append Trim("                                   <vDevTrib>" & cFormatMoeda_NF_1(.m_IBSCBS_gIBSUF_vDevTrib) & "</vDevTrib>")
                    txt_xml.Append Trim("                               </gDevTrib>")
                End If
                If .m_IBSCBS_gIBSUF_pRedAliq <> 0 Or .m_IBSCBS_gIBSUF_pAliqEfet <> 0 Then
                    txt_xml.Append Trim("                               <gRed>")
                    If .m_IBSCBS_gIBSUF_pRedAliq <> 0 Then txt_xml.Append Trim("                                   <pRedAliq>" & cFormatMoeda_NF_1(.m_IBSCBS_gIBSUF_pRedAliq) & "</pRedAliq>")
                    If .m_IBSCBS_gIBSUF_pAliqEfet <> 0 Then txt_xml.Append Trim("                                   <pAliqEfet>" & cFormatMoeda_NF_1(.m_IBSCBS_gIBSUF_pAliqEfet) & "</pAliqEfet>")
                    txt_xml.Append Trim("                               </gRed>")
                End If
                txt_xml.Append Trim("                               <vIBSUF>" & cFormatMoeda_NF_1(.m_IBSCBS_vIBSUF) & "</vIBSUF>")
                txt_xml.Append Trim("                           </gIBSUF>")
                txt_xml.Append Trim("                           <gIBSMun>")
                txt_xml.Append Trim("                               <pIBSMun>" & cFormatMoeda_NF_1(.m_IBSCBS_pIBSMun) & "</pIBSMun>")
                If .m_IBSCBS_gIBSMun_pDif <> 0 Or .m_IBSCBS_gIBSMun_vDif <> 0 Then
                    txt_xml.Append Trim("                               <gDif>")
                    If .m_IBSCBS_gIBSMun_pDif <> 0 Then txt_xml.Append Trim("                                   <pDif>" & cFormatMoeda_NF_1(.m_IBSCBS_gIBSMun_pDif) & "</pDif>")
                    If .m_IBSCBS_gIBSMun_vDif <> 0 Then txt_xml.Append Trim("                                   <vDif>" & cFormatMoeda_NF_1(.m_IBSCBS_gIBSMun_vDif) & "</vDif>")
                    txt_xml.Append Trim("                               </gDif>")
                End If
                If .m_IBSCBS_gIBSMun_vDevTrib <> 0 Then
                    txt_xml.Append Trim("                               <gDevTrib>")
                    txt_xml.Append Trim("                                   <vDevTrib>" & cFormatMoeda_NF_1(.m_IBSCBS_gIBSMun_vDevTrib) & "</vDevTrib>")
                    txt_xml.Append Trim("                               </gDevTrib>")
                End If
                If .m_IBSCBS_gIBSMun_pRedAliq <> 0 Or .m_IBSCBS_gIBSMun_pAliqEfet <> 0 Then
                    txt_xml.Append Trim("                               <gRed>")
                    If .m_IBSCBS_gIBSMun_pRedAliq <> 0 Then txt_xml.Append Trim("                                   <pRedAliq>" & cFormatMoeda_NF_1(.m_IBSCBS_gIBSMun_pRedAliq) & "</pRedAliq>")
                    If .m_IBSCBS_gIBSMun_pAliqEfet <> 0 Then txt_xml.Append Trim("                                   <pAliqEfet>" & cFormatMoeda_NF_1(.m_IBSCBS_gIBSMun_pAliqEfet) & "</pAliqEfet>")
                    txt_xml.Append Trim("                               </gRed>")
                End If
                txt_xml.Append Trim("                               <vIBSMun>" & cFormatMoeda_NF_1(.m_IBSCBS_vIBSMun) & "</vIBSMun>")
                txt_xml.Append Trim("                           </gIBSMun>")
                txt_xml.Append Trim("                           <vIBS>" & cFormatMoeda_NF_1(.m_IBSCBS_vIBS) & "</vIBS>")
                txt_xml.Append Trim("                           <gCBS>")
                txt_xml.Append Trim("                               <pCBS>" & cFormatMoeda_NF_1(.m_IBSCBS_pCBS) & "</pCBS>")
                If .m_IBSCBS_gCBS_pDif <> 0 Or .m_IBSCBS_gCBS_vDif <> 0 Then
                    txt_xml.Append Trim("                               <gDif>")
                    If .m_IBSCBS_gCBS_pDif <> 0 Then txt_xml.Append Trim("                                   <pDif>" & cFormatMoeda_NF_1(.m_IBSCBS_gCBS_pDif) & "</pDif>")
                    If .m_IBSCBS_gCBS_vDif <> 0 Then txt_xml.Append Trim("                                   <vDif>" & cFormatMoeda_NF_1(.m_IBSCBS_gCBS_vDif) & "</vDif>")
                    txt_xml.Append Trim("                               </gDif>")
                End If
                If .m_IBSCBS_gCBS_vDevTrib <> 0 Then
                    txt_xml.Append Trim("                               <gDevTrib>")
                    txt_xml.Append Trim("                                   <vDevTrib>" & cFormatMoeda_NF_1(.m_IBSCBS_gCBS_vDevTrib) & "</vDevTrib>")
                    txt_xml.Append Trim("                               </gDevTrib>")
                End If
                If .m_IBSCBS_gCBS_pRedAliq <> 0 Or .m_IBSCBS_gCBS_pAliqEfet <> 0 Then
                    txt_xml.Append Trim("                               <gRed>")
                    If .m_IBSCBS_gCBS_pRedAliq <> 0 Then txt_xml.Append Trim("                                   <pRedAliq>" & cFormatMoeda_NF_1(.m_IBSCBS_gCBS_pRedAliq) & "</pRedAliq>")
                    If .m_IBSCBS_gCBS_pAliqEfet <> 0 Then txt_xml.Append Trim("                                   <pAliqEfet>" & cFormatMoeda_NF_1(.m_IBSCBS_gCBS_pAliqEfet) & "</pAliqEfet>")
                    txt_xml.Append Trim("                               </gRed>")
                End If
                txt_xml.Append Trim("                               <vCBS>" & cFormatMoeda_NF_1(.m_IBSCBS_vCBS) & "</vCBS>")
                txt_xml.Append Trim("                           </gCBS>")
                If .m_IBSCBS_vCBSST <> 0 Then txt_xml.Append Trim("                           <vCBSST>" & cFormatMoeda_NF_1(.m_IBSCBS_vCBSST) & "</vCBSST>")
                txt_xml.Append Trim("                       </gIBSCBS>")
            End If
            txt_xml.Append Trim("                   </IBSCBS>")
        End If
        If .m_ISSQN_vBC > 0 Or .m_ISSQN_vISSQN > 0 Then
            txt_xml.Append Trim("                   <ISSQN>")
            If .m_ISSQN_vBC > 0 Then txt_xml.Append Trim("                       <vBC>" & cFormatMoeda_NF_1(.m_ISSQN_vBC) & "</vBC>")
            If .m_ISSQN_vAliq > 0 Then txt_xml.Append Trim("                       <vAliq>" & .m_ISSQN_vAliq & "</vAliq>")
            If .m_ISSQN_vISSQN > 0 Then txt_xml.Append Trim("                       <vISSQN>" & cFormatMoeda_NF_1(.m_ISSQN_vISSQN) & "</vISSQN>")
            If Len(.m_ISSQN_cMunFG) > 0 Then txt_xml.Append Trim("                       <cMunFG>" & .m_ISSQN_cMunFG & "</cMunFG>")
            If Len(.m_ISSQN_cListServ) > 0 Then txt_xml.Append Trim("                       <cListServ>" & .m_ISSQN_cListServ & "</cListServ>")
            If .m_ISSQN_vDeducao > 0 Then txt_xml.Append Trim("                       <vDeducao>" & cFormatMoeda_NF_1(.m_ISSQN_vDeducao) & "</vDeducao>")
            If .m_ISSQN_vOutro > 0 Then txt_xml.Append Trim("                       <vOutro>" & cFormatMoeda_NF_1(.m_ISSQN_vOutro) & "</vOutro>")
            If .m_ISSQN_vDescIncond > 0 Then txt_xml.Append Trim("                       <vDescIncond>" & cFormatMoeda_NF_1(.m_ISSQN_vDescIncond) & "</vDescIncond>")
            If .m_ISSQN_vDescCond > 0 Then txt_xml.Append Trim("                       <vDescCond>" & cFormatMoeda_NF_1(.m_ISSQN_vDescCond) & "</vDescCond>")
            If .m_ISSQN_vISSRet > 0 Then txt_xml.Append Trim("                       <vISSRet>" & cFormatMoeda_NF_1(.m_ISSQN_vISSRet) & "</vISSRet>")
            If .m_ISSQN_indISS > 0 Then txt_xml.Append Trim("                       <indISS>" & .m_ISSQN_indISS & "</indISS>")
            If Len(.m_ISSQN_cServico) > 0 Then txt_xml.Append Trim("                       <cServico>" & .m_ISSQN_cServico & "</cServico>")
            If Len(.m_ISSQN_cMun) > 0 Then txt_xml.Append Trim("                       <cMun>" & .m_ISSQN_cMun & "</cMun>")
            If Len(.m_ISSQN_cPais) > 0 Then txt_xml.Append Trim("                       <cPais>" & .m_ISSQN_cPais & "</cPais>")
            If Len(.m_ISSQN_nProcesso) > 0 Then txt_xml.Append Trim("                       <nProcesso>" & .m_ISSQN_nProcesso & "</nProcesso>")
            If .m_ISSQN_indIncentivo > 0 Then txt_xml.Append Trim("                       <indIncentivo>" & .m_ISSQN_indIncentivo & "</indIncentivo>")
            txt_xml.Append Trim("                   </ISSQN>")
        End If
        txt_xml.Append Trim("               </imposto>")
    End With
End Sub

Private Sub GetNota_BuildSignature(txt_xml As hStringBuilder, m_Chave As hStringBuilder)
    txt_xml.Append Trim("       <Signature xmlns=""http://www.w3.org/2000/09/xmldsig#"">")
    txt_xml.Append Trim("           <SignedInfo xmlns=""http://www.w3.org/2000/09/xmldsig#"">")
    txt_xml.Append Trim("               <CanonicalizationMethod Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
    txt_xml.Append Trim("               <SignatureMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#rsa-sha1""/>")
    txt_xml.Append Trim("               <Reference URI=""#" & "NFe" & m_Chave & """>")
    txt_xml.Append Trim("                   <Transforms>")
    txt_xml.Append Trim("                       <Transform Algorithm=""http://www.w3.org/2000/09/xmldsig#enveloped-signature""/>")
    txt_xml.Append Trim("                       <Transform Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
    txt_xml.Append Trim("                   </Transforms>")
    txt_xml.Append Trim("                   <DigestMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#sha1""/>")
    txt_xml.Append Trim("                   <DigestValue>yzGYhUx1/XYYzksWB+fPR3Qc50c=</DigestValue>")
    txt_xml.Append Trim("               </Reference>")
    txt_xml.Append Trim("           </SignedInfo>")
    txt_xml.Append Trim("           <SignatureValue></SignatureValue>")
    txt_xml.Append Trim("           <KeyInfo>")
    txt_xml.Append Trim("               <X509Data>")
    txt_xml.Append Trim("                   <X509Certificate></X509Certificate>")
    txt_xml.Append Trim("               </X509Data>")
    txt_xml.Append Trim("           </KeyInfo>")
    txt_xml.Append Trim("       </Signature>")
    txt_xml.Append Trim("   </NFe>")
End Sub

Private Sub GetNota_BuildInfRespTec(txt_xml As hStringBuilder)
    If Len(m_infRespTec.m_CNPJ) > 0 Then
        txt_xml.Append Trim("       <infRespTec>")
        txt_xml.Append Trim("           <CNPJ>" & m_infRespTec.m_CNPJ & "</CNPJ>")
        txt_xml.Append Trim("           <xContato>" & m_infRespTec.m_xContato & "</xContato>")
        txt_xml.Append Trim("           <email>" & m_infRespTec.m_email & "</email>")
        txt_xml.Append Trim("           <fone>" & m_infRespTec.m_fone & "</fone>")
        If Len(m_infRespTec.m_idCSRT) > 0 Then
            txt_xml.Append Trim("           <idCSRT>" & m_infRespTec.m_idCSRT & "</idCSRT>")
            txt_xml.Append Trim("           <hashCSRT>" & m_infRespTec.m_hashCSRT & "</hashCSRT>")
        End If
        txt_xml.Append Trim("       </infRespTec>")
    End If
End Sub

Private Sub GetNota_BuildInfAdic(txt_xml As hStringBuilder)
    If Len(m_infAdic.infCpl) > 0 Or Len(m_infAdic.m_infAdFisco) > 0 Or m_infAdic.obsCont_count > 0 Or m_infAdic.obsFisco_count > 0 Or m_infAdic.procRef_count > 0 Then
        txt_xml.Append Trim("<infAdic>")
        If Len(m_infAdic.m_infAdFisco) > 0 Then txt_xml.Append Trim("    <infAdFisco>" & m_infAdic.m_infAdFisco & "</infAdFisco>")
        If Len(m_infAdic.infCpl) > 0 Then txt_xml.Append Trim("    <infCpl>" & m_infAdic.infCpl & "</infCpl>")
        Dim cc_inf As Long
        For cc_inf = 0 To m_infAdic.obsCont_count - 1
            txt_xml.Append Trim("    <obsCont>")
            txt_xml.Append Trim("        <xCampo>" & m_infAdic.obsCont_xCampo(cc_inf) & "</xCampo>")
            txt_xml.Append Trim("        <xTexto>" & m_infAdic.obsCont_xTexto(cc_inf) & "</xTexto>")
            txt_xml.Append Trim("    </obsCont>")
        Next
        For cc_inf = 0 To m_infAdic.obsFisco_count - 1
            txt_xml.Append Trim("    <obsFisco>")
            txt_xml.Append Trim("        <xCampo>" & m_infAdic.obsFisco_xCampo(cc_inf) & "</xCampo>")
            txt_xml.Append Trim("        <xTexto>" & m_infAdic.obsFisco_xTexto(cc_inf) & "</xTexto>")
            txt_xml.Append Trim("    </obsFisco>")
        Next
        For cc_inf = 0 To m_infAdic.procRef_count - 1
            txt_xml.Append Trim("    <procRef>")
            txt_xml.Append Trim("        <nProc>" & m_infAdic.procRef_nProc(cc_inf) & "</nProc>")
            txt_xml.Append Trim("        <indProc>" & m_infAdic.procRef_indProc(cc_inf) & "</indProc>")
            If Len(m_infAdic.procRef_tpAto(cc_inf)) > 0 Then txt_xml.Append Trim("        <tpAto>" & m_infAdic.procRef_tpAto(cc_inf) & "</tpAto>")
            txt_xml.Append Trim("    </procRef>")
        Next
        txt_xml.Append Trim("</infAdic>")
    End If
End Sub

Private Sub GetNota_BuildPag(txt_xml As hStringBuilder)
    txt_xml.Append Trim("            <pag>")
    If m_cPagamentos = -1 Then
        m_cPagamentos = 0
        ReDim Preserve m_Pagamentos(0)
        Set m_Pagamentos(0) = New hNFe4_Doc_detPag
        m_Pagamentos(0).tPag = "01"
        m_Pagamentos(0).vPag = m_total.m_vNF
    End If
    Dim cc_pag As Long
    For cc_pag = 0 To m_cPagamentos
        With m_Pagamentos(cc_pag)
            txt_xml.Append Trim("            <detPag>")
            If Len(.indPag) > 0 Then txt_xml.Append Trim("                <indPag>" & .indPag & "</indPag>")
            txt_xml.Append Trim("                <tPag>" & .tPag & "</tPag>")
            If Len(.xPag) > 0 Then txt_xml.Append Trim("                <xPag>" & .xPag & "</xPag>")
            txt_xml.Append Trim("                <vPag>" & cFormatMoeda_NF_1(.vPag) & "</vPag>")
            If Len(.dPag) > 0 Then txt_xml.Append Trim("                <dPag>" & .dPag & "</dPag>")
            If Len(.CNPJPag) > 0 Then txt_xml.Append Trim("                <CNPJPag>" & .CNPJPag & "</CNPJPag>")
            If Len(.UFPag) > 0 Then txt_xml.Append Trim("                <UFPag>" & .UFPag & "</UFPag>")
            If Len(.card.tpIntegra) > 0 Then
                txt_xml.Append Trim("                <card>")
                txt_xml.Append Trim("                    <tpIntegra>" & .card.tpIntegra & "</tpIntegra>")
                If Len(.card.cnpj) > 0 Then txt_xml.Append Trim("                    <CNPJ>" & .card.cnpj & "</CNPJ>")
                If Len(.card.tBand) > 0 Then txt_xml.Append Trim("                    <tBand>" & .card.tBand & "</tBand>")
                If Len(.card.cAut) > 0 Then txt_xml.Append Trim("                    <cAut>" & .card.cAut & "</cAut>")
                If Len(.card.CNPJReceb) > 0 Then txt_xml.Append Trim("                    <CNPJReceb>" & .card.CNPJReceb & "</CNPJReceb>")
                If Len(.card.idTermPag) > 0 Then txt_xml.Append Trim("                    <idTermPag>" & .card.idTermPag & "</idTermPag>")
                txt_xml.Append Trim("                </card>")
            End If
            txt_xml.Append Trim("            </detPag>")
        End With
    Next
    txt_xml.Append Trim("                <vTroco>" & cFormatMoeda_NF_1(m_vTroco) & "</vTroco>")
    txt_xml.Append Trim("            </pag>")
End Sub

Private Sub GetNota_BuildCobr(txt_xml As hStringBuilder)
    If m_ide.m_mod = "55" Then
        If m_cDup >= 0 Or Len(CobrFat.nFat) > 0 Then
            txt_xml.Append Trim("           <cobr>")
            If Len(CobrFat.nFat) > 0 Then
                txt_xml.Append Trim("               <fat>")
                txt_xml.Append Trim("                   <nFat>" & CobrFat.nFat & "</nFat>")
                If CobrFat.vOrig > 0 Then txt_xml.Append Trim("                   <vOrig>" & cFormatMoeda_NF_1(CobrFat.vOrig) & "</vOrig>")
                If CobrFat.vDesc > 0 Then txt_xml.Append Trim("                   <vDesc>" & cFormatMoeda_NF_1(CobrFat.vDesc) & "</vDesc>")
                If CobrFat.vLiq > 0 Then txt_xml.Append Trim("                   <vLiq>" & cFormatMoeda_NF_1(CobrFat.vLiq) & "</vLiq>")
                txt_xml.Append Trim("               </fat>")
            End If
            Dim cc_dup As Long
            For cc_dup = 0 To m_cDup
                With m_Dup(cc_dup)
                    txt_xml.Append Trim("               <dup>")
                    If Len(.nDup) > 0 Then txt_xml.Append Trim("                   <nDup>" & .nDup & "</nDup>")
                    If Len(.dVenc) > 0 Then txt_xml.Append Trim("                   <dVenc>" & .dVenc & "</dVenc>")
                    If .vDup > 0 Then txt_xml.Append Trim("                   <vDup>" & cFormatMoeda_NF_1(.vDup) & "</vDup>")
                    txt_xml.Append Trim("               </dup>")
                End With
            Next
            txt_xml.Append Trim("           </cobr>")
        End If
    End If
End Sub

Private Sub GetNota_BuildTransp(txt_xml As hStringBuilder)
    Dim cc As Long
    txt_xml.Append Trim("           <transp>")
    txt_xml.Append Trim("               <modFrete>" & m_modFrete & "</modFrete>")
    If Len(m_transp.m_CNPJ) > 0 Or Len(m_transp.m_CPF) > 0 Then
        txt_xml.Append Trim("               <transporta>")
        If Len(m_transp.m_CNPJ) > 0 Then txt_xml.Append Trim("                   <CNPJ>" & m_transp.m_CNPJ & "</CNPJ>") Else txt_xml.Append Trim("                   <CPF>" & m_transp.m_CPF & "</CPF>")
        txt_xml.Append Trim("                   <xNome>" & m_transp.m_xNome & "</xNome>")
        If Len(m_transp.m_IE) > 0 Then txt_xml.Append Trim("                   <IE>" & m_transp.m_IE & "</IE>")
        If Len(m_transp.m_xEnder) > 0 Then txt_xml.Append Trim("                   <xEnder>" & m_transp.m_xEnder & "</xEnder>")
        txt_xml.Append Trim("                   <xMun>" & m_transp.m_xMun & "</xMun>")
        txt_xml.Append Trim("                   <UF>" & m_transp.m_UF & "</UF>")
        txt_xml.Append Trim("               </transporta>")
    End If
    If m_transp.m_ret_vServ > 0 Then
        txt_xml.Append Trim("               <retTransp>")
        txt_xml.Append Trim("                   <vServ>" & cFormatMoeda_NF_1(m_transp.m_ret_vServ) & "</vServ>")
        txt_xml.Append Trim("                   <vBCRet>" & cFormatMoeda_NF_1(m_transp.m_ret_vBCRet) & "</vBCRet>")
        txt_xml.Append Trim("                   <pICMSRet>" & cFormatMoeda_NF_1(m_transp.m_ret_pICMSRet) & "</pICMSRet>")
        txt_xml.Append Trim("                   <vICMSRet>" & cFormatMoeda_NF_1(m_transp.m_ret_vICMSRet) & "</vICMSRet>")
        txt_xml.Append Trim("                   <CFOP>" & m_transp.m_ret_CFOP & "</CFOP>")
        txt_xml.Append Trim("                   <cMunFG>" & m_transp.m_ret_cMunFG & "</cMunFG>")
        txt_xml.Append Trim("               </retTransp>")
    End If
    If Len(m_veicTransp.m_placa) > 0 Then
        txt_xml.Append Trim("               <veicTransp>")
        txt_xml.Append Trim("                   <placa>" & m_veicTransp.m_placa & "</placa>")
        txt_xml.Append Trim("                   <UF>" & m_veicTransp.m_UF & "</UF>")
        If Len(m_veicTransp.m_RNTC) > 0 Then txt_xml.Append Trim("                   <RNTC>" & m_veicTransp.m_RNTC & "</RNTC>")
        txt_xml.Append Trim("               </veicTransp>")
    End If
    For cc = 0 To m_cReboques
        With m_Reboques(cc)
            If Len(.m_placa) > 0 Then
                txt_xml.Append Trim("               <reboque>")
                txt_xml.Append Trim("                   <placa>" & .m_placa & "</placa>")
                txt_xml.Append Trim("                   <UF>" & .m_UF & "</UF>")
                If Len(.m_RNTC) > 0 Then txt_xml.Append Trim("                   <RNTC>" & .m_RNTC & "</RNTC>")
                txt_xml.Append Trim("               </reboque>")
            End If
        End With
    Next
    For cc = 0 To m_cVolumes
        With m_Volumes(cc)
            If .m_qVol > 0 Or Len(.m_esp) > 0 Then
                txt_xml.Append Trim("               <vol>")
                If .m_qVol > 0 Then txt_xml.Append Trim("                   <qVol>" & .m_qVol & "</qVol>")
                If Len(.m_esp) > 0 Then txt_xml.Append Trim("                   <esp>" & .m_esp & "</esp>")
                If Len(.m_marca) > 0 Then txt_xml.Append Trim("                   <marca>" & .m_marca & "</marca>")
                If Len(.m_nVol) > 0 Then txt_xml.Append Trim("                   <nVol>" & .m_nVol & "</nVol>")
                If .m_pesoL > 0 Then txt_xml.Append Trim("                   <pesoL>" & cFormatMoeda_NF_1(.m_pesoL) & "</pesoL>")
                If .m_pesoB > 0 Then txt_xml.Append Trim("                   <pesoB>" & cFormatMoeda_NF_1(.m_pesoB) & "</pesoB>")
                txt_xml.Append Trim("               </vol>")
            End If
        End With
    Next
    txt_xml.Append Trim("           </transp>")
End Sub

Private Sub GetNota_BuildTotal(txt_xml As hStringBuilder)
    txt_xml.Append Trim("           <total>")
    If m_ide.m_mod = "65" Then
        txt_xml.Append Trim("               <ICMSTot>")
        txt_xml.Append Trim("                   <vBC>" & cFormatMoeda_NF_1(m_total.m_vBC + 0) & "</vBC>")
        txt_xml.Append Trim("                   <vICMS>" & cFormatMoeda_NF_1(m_total.m_vICMS) & "</vICMS>")
        txt_xml.Append Trim("                   <vICMSDeson>" & cFormatMoeda_NF_1(m_total.m_vICMSDeson) & "</vICMSDeson>")
        txt_xml.Append Trim("                   <vFCPUFDest>" & cFormatMoeda_NF_1(m_total.m_vFCPUFDest) & "</vFCPUFDest>")
        txt_xml.Append Trim("                   <vICMSUFDest>" & cFormatMoeda_NF_1(m_total.m_vICMSUFDest) & "</vICMSUFDest>")
        txt_xml.Append Trim("                   <vICMSUFRemet>" & cFormatMoeda_NF_1(m_total.m_vICMSUFRemet) & "</vICMSUFRemet>")
        txt_xml.Append Trim("                   <vFCP>" & cFormatMoeda_NF_1(m_total.m_vFCP) & "</vFCP>")
        txt_xml.Append Trim("                   <vBCST>0</vBCST>")
        txt_xml.Append Trim("                   <vST>" & cFormatMoeda_NF_1(m_total.m_vST) & "</vST>")
        txt_xml.Append Trim("                   <vFCPST>" & cFormatMoeda_NF_1(m_total.m_vFCPST) & "</vFCPST>")
        txt_xml.Append Trim("                   <vFCPSTRet>" & cFormatMoeda_NF_1(m_total.m_vFCPSTRet) & "</vFCPSTRet>")
        txt_xml.Append Trim("                   <vProd>" & cFormatMoeda_NF_1(m_total.m_vProd) & "</vProd>")
        txt_xml.Append Trim("                   <vFrete>0</vFrete>")
        txt_xml.Append Trim("                   <vSeg>0</vSeg>")
        txt_xml.Append Trim("                   <vDesc>" & cFormatMoeda_NF_1(m_total.m_vDesc) & "</vDesc>")
        txt_xml.Append Trim("                   <vII>" & cFormatMoeda_NF_1(m_total.m_vII) & "</vII>")
        txt_xml.Append Trim("                   <vIPI>" & cFormatMoeda_NF_1(m_total.m_vIPI) & "</vIPI>")
        txt_xml.Append Trim("                   <vIPIDevol>0</vIPIDevol>")
        txt_xml.Append Trim("                   <vPIS>" & cFormatMoeda_NF_1(m_total.m_vPIS) & "</vPIS>")
        txt_xml.Append Trim("                   <vCOFINS>" & cFormatMoeda_NF_1(m_total.m_vCOFINS) & "</vCOFINS>")
        txt_xml.Append Trim("                   <vOutro>" & cFormatMoeda_NF_1(m_total.m_vOutro) & "</vOutro>")
        txt_xml.Append Trim("                   <vNF>" & cFormatMoeda_NF_1(m_total.m_vProd - m_total.m_vDesc - m_total.m_vICMSDeson + m_total.m_vST + m_total.m_vFrete + m_total.m_vSeg + m_total.m_vOutro + m_total.m_vII + m_total.m_vIPI) & "</vNF>")
        m_total.m_vNF = m_total.m_vProd - m_total.m_vDesc - m_total.m_vICMSDeson + m_total.m_vST + m_total.m_vFrete + m_total.m_vSeg + m_total.m_vOutro + m_total.m_vII + m_total.m_vIPI
        txt_xml.Append Trim("                   <vTotTrib>" & cFormatMoeda_NF_1(m_total.m_vTotTrib) & "</vTotTrib>")
        txt_xml.Append Trim("               </ICMSTot>")
        GetNota_BuildTotalsAfterICMS txt_xml
    ElseIf m_ide.m_mod = "55" Then
        txt_xml.Append Trim("               <ICMSTot>")
        txt_xml.Append Trim("                   <vBC>" & cFormatMoeda_NF_1(m_total.m_vBC + 0) & "</vBC>")
        txt_xml.Append Trim("                   <vICMS>" & cFormatMoeda_NF_1(m_total.m_vICMS) & "</vICMS>")
        txt_xml.Append Trim("                   <vICMSDeson>" & cFormatMoeda_NF_1(m_total.m_vICMSDeson) & "</vICMSDeson>")
        txt_xml.Append Trim("                   <vFCPUFDest>" & cFormatMoeda_NF_1(m_total.m_vFCPUFDest) & "</vFCPUFDest>")
        txt_xml.Append Trim("                   <vICMSUFDest>" & cFormatMoeda_NF_1(m_total.m_vICMSUFDest) & "</vICMSUFDest>")
        txt_xml.Append Trim("                   <vICMSUFRemet>" & cFormatMoeda_NF_1(m_total.m_vICMSUFRemet) & "</vICMSUFRemet>")
        txt_xml.Append Trim("                   <vFCP>" & cFormatMoeda_NF_1(m_total.m_vFCP) & "</vFCP>")
        txt_xml.Append Trim("                   <vBCST>0</vBCST>")
        txt_xml.Append Trim("                   <vST>" & cFormatMoeda_NF_1(m_total.m_vST) & "</vST>")
        txt_xml.Append Trim("                   <vFCPST>" & cFormatMoeda_NF_1(m_total.m_vFCPST) & "</vFCPST>")
        txt_xml.Append Trim("                   <vFCPSTRet>" & cFormatMoeda_NF_1(m_total.m_vFCPSTRet) & "</vFCPSTRet>")
        txt_xml.Append Trim("                   <vProd>" & cFormatMoeda_NF_1(m_total.m_vProd) & "</vProd>")
        txt_xml.Append Trim("                   <vFrete>" & cFormatMoeda_NF_1(m_total.m_vFrete) & "</vFrete>")
        txt_xml.Append Trim("                   <vSeg>" & cFormatMoeda_NF_1(m_total.m_vSeg) & "</vSeg>")
        txt_xml.Append Trim("                   <vDesc>" & cFormatMoeda_NF_1(m_total.m_vDesc) & "</vDesc>")
        txt_xml.Append Trim("                   <vII>" & cFormatMoeda_NF_1(m_total.m_vII) & "</vII>")
        txt_xml.Append Trim("                   <vIPI>" & cFormatMoeda_NF_1(m_total.m_vIPI) & "</vIPI>")
        txt_xml.Append Trim("                   <vIPIDevol>0</vIPIDevol>")
        txt_xml.Append Trim("                   <vPIS>" & cFormatMoeda_NF_1(m_total.m_vPIS) & "</vPIS>")
        txt_xml.Append Trim("                   <vCOFINS>" & cFormatMoeda_NF_1(m_total.m_vCOFINS) & "</vCOFINS>")
        txt_xml.Append Trim("                   <vOutro>" & cFormatMoeda_NF_1(m_total.m_vOutro) & "</vOutro>")
        m_total.m_vNF = m_total.m_vProd - m_total.m_vDesc - m_total.m_vICMSDeson + m_total.m_vST + m_total.m_vFrete + m_total.m_vSeg + m_total.m_vOutro + m_total.m_vII + m_total.m_vIPI
        txt_xml.Append Trim("                   <vNF>" & cFormatMoeda_NF_1(m_total.m_vNF) & "</vNF>")
        txt_xml.Append Trim("                   <vTotTrib>" & cFormatMoeda_NF_1(m_total.m_vTotTrib) & "</vTotTrib>")
        txt_xml.Append Trim("               </ICMSTot>")
        GetNota_BuildTotalsAfterICMS txt_xml
    End If
    txt_xml.Append Trim("           </total>")
End Sub

Private Sub GetNota_BuildTotalsAfterICMS(txt_xml As hStringBuilder)
    If m_total.m_ISSQNtot_vBC <> 0 Or m_total.m_ISSQNtot_vISS <> 0 Then
        txt_xml.Append Trim("               <ISSQNtot>")
        If m_total.m_ISSQNtot_vServ <> 0 Then txt_xml.Append Trim("                   <vServ>" & cFormatMoeda_NF_1(m_total.m_ISSQNtot_vServ) & "</vServ>")
        If m_total.m_ISSQNtot_vBC <> 0 Then txt_xml.Append Trim("                   <vBC>" & cFormatMoeda_NF_1(m_total.m_ISSQNtot_vBC) & "</vBC>")
        If m_total.m_ISSQNtot_vISS <> 0 Then txt_xml.Append Trim("                   <vISS>" & cFormatMoeda_NF_1(m_total.m_ISSQNtot_vISS) & "</vISS>")
        If m_total.m_ISSQNtot_vPIS <> 0 Then txt_xml.Append Trim("                   <vPIS>" & cFormatMoeda_NF_1(m_total.m_ISSQNtot_vPIS) & "</vPIS>")
        If m_total.m_ISSQNtot_vCOFINS <> 0 Then txt_xml.Append Trim("                   <vCOFINS>" & cFormatMoeda_NF_1(m_total.m_ISSQNtot_vCOFINS) & "</vCOFINS>")
        If Len(m_total.m_ISSQNtot_dCompet) > 0 Then txt_xml.Append Trim("                   <dCompet>" & m_total.m_ISSQNtot_dCompet & "</dCompet>")
        If m_total.m_ISSQNtot_vDeducao <> 0 Then txt_xml.Append Trim("                   <vDeducao>" & cFormatMoeda_NF_1(m_total.m_ISSQNtot_vDeducao) & "</vDeducao>")
        If m_total.m_ISSQNtot_vOutro <> 0 Then txt_xml.Append Trim("                   <vOutro>" & cFormatMoeda_NF_1(m_total.m_ISSQNtot_vOutro) & "</vOutro>")
        If m_total.m_ISSQNtot_vDescIncond <> 0 Then txt_xml.Append Trim("                   <vDescIncond>" & cFormatMoeda_NF_1(m_total.m_ISSQNtot_vDescIncond) & "</vDescIncond>")
        If m_total.m_ISSQNtot_vDescCond <> 0 Then txt_xml.Append Trim("                   <vDescCond>" & cFormatMoeda_NF_1(m_total.m_ISSQNtot_vDescCond) & "</vDescCond>")
        If m_total.m_ISSQNtot_vISSRet <> 0 Then txt_xml.Append Trim("                   <vISSRet>" & cFormatMoeda_NF_1(m_total.m_ISSQNtot_vISSRet) & "</vISSRet>")
        If Len(m_total.m_ISSQNtot_cRegTrib) > 0 Then txt_xml.Append Trim("                   <cRegTrib>" & m_total.m_ISSQNtot_cRegTrib & "</cRegTrib>")
        txt_xml.Append Trim("               </ISSQNtot>")
    End If
    If m_total.m_retTrib_vRetPIS <> 0 Or m_total.m_retTrib_vRetCOFINS <> 0 Or m_total.m_retTrib_vRetCSLL <> 0 Or m_total.m_retTrib_vRetIRRF <> 0 Or m_total.m_retTrib_vRetINSS <> 0 Or m_total.m_retTrib_vRetISS <> 0 Then
        txt_xml.Append Trim("               <retTrib>")
        If m_total.m_retTrib_vRetPIS <> 0 Then txt_xml.Append Trim("                   <vRetPIS>" & cFormatMoeda_NF_1(m_total.m_retTrib_vRetPIS) & "</vRetPIS>")
        If m_total.m_retTrib_vRetCOFINS <> 0 Then txt_xml.Append Trim("                   <vRetCOFINS>" & cFormatMoeda_NF_1(m_total.m_retTrib_vRetCOFINS) & "</vRetCOFINS>")
        If m_total.m_retTrib_vRetCSLL <> 0 Then txt_xml.Append Trim("                   <vRetCSLL>" & cFormatMoeda_NF_1(m_total.m_retTrib_vRetCSLL) & "</vRetCSLL>")
        If m_total.m_retTrib_vRetIRRF <> 0 Then txt_xml.Append Trim("                   <vRetIRRF>" & cFormatMoeda_NF_1(m_total.m_retTrib_vRetIRRF) & "</vRetIRRF>")
        If m_total.m_retTrib_vRetINSS <> 0 Then txt_xml.Append Trim("                   <vRetINSS>" & cFormatMoeda_NF_1(m_total.m_retTrib_vRetINSS) & "</vRetINSS>")
        If m_total.m_retTrib_vRetISS <> 0 Then txt_xml.Append Trim("                   <vRetISS>" & cFormatMoeda_NF_1(m_total.m_retTrib_vRetISS) & "</vRetISS>")
        txt_xml.Append Trim("               </retTrib>")
    End If
    If m_total.m_vIS <> 0 Then
        txt_xml.Append Trim("               <ISTot>")
        txt_xml.Append Trim("                   <vIS>" & cFormatMoeda_NF_1(m_total.m_vIS) & "</vIS>")
        txt_xml.Append Trim("               </ISTot>")
    End If
    If m_total.m_vIBS <> 0 Or m_total.m_vCBS <> 0 Then
        txt_xml.Append Trim("               <IBSCBSTot>")
        txt_xml.Append Trim("                   <vBCIBSCBS>" & cFormatMoeda_NF_1(m_total.m_vBCIBSCBS) & "</vBCIBSCBS>")
        txt_xml.Append Trim("                   <gIBS>")
        txt_xml.Append Trim("                       <gIBSUF>")
        txt_xml.Append Trim("                           <vDif>" & cFormatMoeda_NF_1(m_total.m_vDifIBSUF) & "</vDif>")
        txt_xml.Append Trim("                           <vDevTrib>" & cFormatMoeda_NF_1(m_total.m_vDevTribIBSUF) & "</vDevTrib>")
        txt_xml.Append Trim("                           <vIBSUF>" & cFormatMoeda_NF_1(m_total.m_vIBSUF) & "</vIBSUF>")
        txt_xml.Append Trim("                       </gIBSUF>")
        txt_xml.Append Trim("                       <gIBSMun>")
        txt_xml.Append Trim("                           <vDif>" & cFormatMoeda_NF_1(m_total.m_vDifIBSMun) & "</vDif>")
        txt_xml.Append Trim("                           <vDevTrib>" & cFormatMoeda_NF_1(m_total.m_vDevTribIBSMun) & "</vDevTrib>")
        txt_xml.Append Trim("                           <vIBSMun>" & cFormatMoeda_NF_1(m_total.m_vIBSMun) & "</vIBSMun>")
        txt_xml.Append Trim("                       </gIBSMun>")
        txt_xml.Append Trim("                       <vIBS>" & cFormatMoeda_NF_1(m_total.m_vIBS) & "</vIBS>")
        txt_xml.Append Trim("                   </gIBS>")
        txt_xml.Append Trim("                   <gCBS>")
        txt_xml.Append Trim("                       <vDif>" & cFormatMoeda_NF_1(m_total.m_vDifCBS) & "</vDif>")
        txt_xml.Append Trim("                       <vDevTrib>" & cFormatMoeda_NF_1(m_total.m_vDevTribCBS) & "</vDevTrib>")
        txt_xml.Append Trim("                       <vCBS>" & cFormatMoeda_NF_1(m_total.m_vCBS) & "</vCBS>")
        txt_xml.Append Trim("                   </gCBS>")
        If m_total.m_vIBSEstCred <> 0 Or m_total.m_vCBSEstCred <> 0 Then
            txt_xml.Append Trim("                   <gEstornoCred>")
            txt_xml.Append Trim("                       <vIBSEstCred>" & cFormatMoeda_NF_1(m_total.m_vIBSEstCred) & "</vIBSEstCred>")
            txt_xml.Append Trim("                       <vCBSEstCred>" & cFormatMoeda_NF_1(m_total.m_vCBSEstCred) & "</vCBSEstCred>")
            txt_xml.Append Trim("                   </gEstornoCred>")
        End If
        txt_xml.Append Trim("               </IBSCBSTot>")
    End If
End Sub

Private Sub Class_Initialize()
    m_cItens = -1
    m_cReboques = -1
    m_cVolumes = -1
    m_cPagamentos = -1
    m_cDup = -1
    m_modFrete = "9"
End Sub

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


