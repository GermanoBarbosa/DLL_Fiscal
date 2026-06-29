VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_det_prod_DI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public nDI As String
Public dDI As String
Public xLocDesemb As String
Public UFDesemb As String
Public dDesemb As String
Public tpViaTransp As Long
Public vAFRMM As Double
Public tpIntermedio As Long
Public CNPJ As String
Public UFTerceiro As String
Public cExportador As String

Dim m_ADI() As hNFe4_Doc_det_prod_DI_ADI, m_cADI As Long

Public Property Get AdiCount() As Long
    AdiCount = m_cADI + 1
End Property

Public Property Get Adi(Index As Long) As hNFe4_Doc_det_prod_DI_ADI
    Set Adi = m_ADI(Index)
End Property

Sub Clear_adi()
    m_cADI = -1
    Erase m_ADI
End Sub

Sub Add_adi(m_Item As hNFe4_Doc_det_prod_DI_ADI)
    m_cADI = m_cADI + 1
    ReDim Preserve m_ADI(m_cADI)
    Set m_ADI(m_cADI) = m_Item
End Sub

Private Sub Class_Initialize()
    m_cADI = -1
End Sub
