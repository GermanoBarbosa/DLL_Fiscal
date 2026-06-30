VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_total"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public m_vBC As Double '>686.10</vBC>")
Public m_vICMS As Double '>48.03</vICMS>")
Public m_vICMSDeson As Double '>0</vICMSDeson>")
Public m_vFCPUFDest As Double '>0</vFCPUFDest>")
Public m_vICMSUFDest As Double '>27.44</vICMSUFDest>")
Public m_vICMSUFRemet As Double '>41.16</vICMSUFRemet>")
Public m_vBCST As Double '>0</vBCST>")
Public m_vST As Double '>0</vST>")
Public m_vProd As Double '>520.00</vProd>")
Public m_vFrete As Double '>166.10</vFrete>")
Public m_vSeg As Double '>0</vSeg>")
Public m_vDesc As Double '>0</vDesc>")
Public m_vII As Double '>'0</vII>")
Public m_vIPI As Double '>0</vIPI>")
Public m_vPIS As Double '>'11.32</vPIS>")
Public m_vCOFINS As Double '>52.14</vCOFINS>")
Public m_vOutro As Double '>0</vOutro>")
Public m_vNF As Double '>686.10</vNF>")
Public m_vTotTrib As Double '>317.39</vTotTrib>")
' === IBS/CBS totais (NT 2025.002 - IBSCBSTot) ===
Public m_vBCIBSCBS As Double
Public m_vIBSUF As Double
Public m_vIBSMun As Double
Public m_vIBS As Double
Public m_vCBS As Double
Public m_vCBSST As Double
' gDif / gDevTrib totals (NT 2025.002)
Public m_vDifIBSUF As Double
Public m_vDevTribIBSUF As Double
Public m_vDifIBSMun As Double
Public m_vDevTribIBSMun As Double
Public m_vDifCBS As Double
Public m_vDevTribCBS As Double
' gEstornoCred totals (NT 2025.002)
Public m_vIBSEstCred As Double
Public m_vCBSEstCred As Double
Public m_vIS As Double                          ' ISTot
' === ISSQNtot ===
Public m_ISSQNtot_vServ As Double
Public m_ISSQNtot_vBC As Double
Public m_ISSQNtot_vISS As Double
Public m_ISSQNtot_vPIS As Double
Public m_ISSQNtot_vCOFINS As Double
Public m_ISSQNtot_dCompet As String
Public m_ISSQNtot_vDeducao As Double
Public m_ISSQNtot_vOutro As Double
Public m_ISSQNtot_vDescIncond As Double
Public m_ISSQNtot_vDescCond As Double
Public m_ISSQNtot_vISSRet As Double
Public m_ISSQNtot_cRegTrib As String
' === retTrib (Retencao de Tributos Federais) ===
Public m_retTrib_vRetPIS As Double
Public m_retTrib_vRetCOFINS As Double
Public m_retTrib_vRetCSLL As Double
Public m_retTrib_vRetIRRF As Double
Public m_retTrib_vRetINSS As Double
Public m_retTrib_vRetISS As Double
' === FCP ===
Public m_vFCP As Double
Public m_vFCPST As Double
Public m_vFCPSTRet As Double

Sub Clear()
m_vBC = 0
m_vICMS = 0
m_vICMSDeson = 0
m_vFCPUFDest = 0
m_vICMSUFDest = 0
m_vICMSUFRemet = 0
m_vBCST = 0
m_vST = 0
m_vProd = 0
m_vFrete = 0
m_vSeg = 0
m_vDesc = 0
m_vII = 0
m_vIPI = 0
m_vPIS = 0
m_vCOFINS = 0
m_vOutro = 0
m_vNF = 0
m_vTotTrib = 0
m_vBCIBSCBS = 0
m_vIBSUF = 0
m_vIBSMun = 0
m_vIBS = 0
m_vCBS = 0
m_vCBSST = 0
m_vDifIBSUF = 0
m_vDevTribIBSUF = 0
m_vDifIBSMun = 0
m_vDevTribIBSMun = 0
m_vDifCBS = 0
m_vDevTribCBS = 0
m_vIBSEstCred = 0
m_vCBSEstCred = 0
m_vIS = 0
m_ISSQNtot_vServ = 0
m_ISSQNtot_vBC = 0
m_ISSQNtot_vISS = 0
m_ISSQNtot_vPIS = 0
m_ISSQNtot_vCOFINS = 0
m_ISSQNtot_dCompet = ""
m_ISSQNtot_vDeducao = 0
m_ISSQNtot_vOutro = 0
m_ISSQNtot_vDescIncond = 0
m_ISSQNtot_vDescCond = 0
m_ISSQNtot_vISSRet = 0
m_ISSQNtot_cRegTrib = ""
m_retTrib_vRetPIS = 0
m_retTrib_vRetCOFINS = 0
m_retTrib_vRetCSLL = 0
m_retTrib_vRetIRRF = 0
m_retTrib_vRetINSS = 0
m_retTrib_vRetISS = 0
m_vFCP = 0
m_vFCPST = 0
m_vFCPSTRet = 0
End Sub

