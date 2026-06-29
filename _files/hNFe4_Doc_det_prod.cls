VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_det_prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public m_cProd      As String
Public m_cEAN       As String
Public m_xProd      As String
Public m_NCM        As String
Public m_CFOP       As String
Public m_uCom       As String
Public m_qCom       As Double
Public m_vUnCom     As Double
Public m_vProd      As Double
Public m_cEANTrib   As String
Public m_uTrib      As String
Public m_qTrib      As Double
Public m_vUnTrib    As Double
Public m_vFrete     As Double
Public m_indTot     As Long
Public m_CEST       As String

Dim m_DI() As hNFe4_Doc_det_prod_DI
Public m_cDI As Long
Private m_comb_int As New hNFe4_Doc_det_prod_comb

Public Property Get DiCount() As Long
    DiCount = m_cDI + 1
End Property

Public Property Get Di(Index As Long) As hNFe4_Doc_det_prod_DI
    Set Di = m_DI(Index)
End Property


Public Property Get comb() As hNFe4_Doc_det_prod_comb
    Set comb = m_comb_int
End Property

Sub Clear_di()
    m_cDI = -1
    Erase m_DI
End Sub

Sub Add_di(m_Item As hNFe4_Doc_det_prod_DI)
    m_cDI = m_cDI + 1
    ReDim Preserve m_DI(m_cDI)
    Set m_DI(m_cDI) = m_Item
End Sub

Private Sub Class_Initialize()
    m_cDI = -1
End Sub


