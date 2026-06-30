VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_ide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public m_cNF As String
Public m_cUF As String
Public m_mod As String
Public m_serie  As String
Public m_nNF As String
Public m_dhEmi As Date

Public m_cMunFG  As String
Public m_cMunFGIBS As String            ' informar municipio ocorrencia fato gerador IBS/CBS (indPres=5 e sem endereco dest/entrega)
Public m_natOp As String
Public m_tpNF As String
Public m_indPag As String
Public m_dhSaiEnt
Public m_idDest As Long '>2</idDest>
'Public m_cMunFG>4208203</cMunFG>
'Public m_tpImp>1</tpImp>
Public m_tpEmis As String
'Public m_cDV>3</cDV>
'Public m_tpAmb>2</tpAmb>
Public m_finNFe As String
Public m_tpNFDebito As String           ' tpNFDebito - tipo nota debito (obrigatorio se finNFe=4 e CRT=3, NT 2025.002)
Public m_tpNFCredito As String          ' tpNFCredito - tipo nota credito (obrigatorio se finNFe=5 e CRT=3, NT 2025.002)
Public m_indFinal As String
Public m_indPres As String
Public m_indIntermed As String          ' indIntermed - indicador de intermediador/marketplace (0=sem, 1=com)
Public m_procEmi As String
Public m_dhCont As String               ' dhCont - data/hora contingencia (obrigatorio se tpEmis <> 1)
Public m_xJust As String                ' xJust - justificativa contingencia (obrigatorio se tpEmis <> 1)
Public m_verProc As String              ' versao do aplicativo emissor

' --- NFref (NF referenciadas) ---
Dim m_NFref_refNFe() As String, m_cNFref As Long

Public Property Get NFrefCount() As Long
    NFrefCount = m_cNFref + 1
End Property

Public Property Get NFref_refNFe(Index As Long) As String
    NFref_refNFe = m_NFref_refNFe(Index)
End Property

Sub Clear_NFref()
    m_cNFref = -1
    Erase m_NFref_refNFe
End Sub

Sub Add_NFref(m_refNFe As String)
    m_cNFref = m_cNFref + 1
    ReDim Preserve m_NFref_refNFe(m_cNFref)
    m_NFref_refNFe(m_cNFref) = m_refNFe
End Sub

