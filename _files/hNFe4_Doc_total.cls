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
Public m_vIS As Double                          ' ISTot
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
m_vIS = 0
m_vFCP = 0
m_vFCPST = 0
m_vFCPSTRet = 0
End Sub

