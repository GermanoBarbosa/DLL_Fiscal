VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_infAdic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public infCpl As String
Public m_infAdFisco As String

Dim m_obsCont_xCampo() As String, m_obsCont_xTexto() As String, m_cObsCont As Long
Dim m_obsFisco_xCampo() As String, m_obsFisco_xTexto() As String, m_cObsFisco As Long
Dim m_procRef_nProc() As String, m_procRef_indProc() As String, m_procRef_tpAto() As String, m_cProcRef As Long

Public Sub Add_obsCont(ByVal xCampo As String, ByVal xTexto As String)
    m_cObsCont = m_cObsCont + 1
    ReDim Preserve m_obsCont_xCampo(m_cObsCont)
    ReDim Preserve m_obsCont_xTexto(m_cObsCont)
    m_obsCont_xCampo(m_cObsCont) = xCampo
    m_obsCont_xTexto(m_cObsCont) = xTexto
End Sub

Public Sub Add_obsFisco(ByVal xCampo As String, ByVal xTexto As String)
    m_cObsFisco = m_cObsFisco + 1
    ReDim Preserve m_obsFisco_xCampo(m_cObsFisco)
    ReDim Preserve m_obsFisco_xTexto(m_cObsFisco)
    m_obsFisco_xCampo(m_cObsFisco) = xCampo
    m_obsFisco_xTexto(m_cObsFisco) = xTexto
End Sub

Public Sub Add_procRef(ByVal nProc As String, ByVal indProc As String, Optional ByVal tpAto As String = "")
    m_cProcRef = m_cProcRef + 1
    ReDim Preserve m_procRef_nProc(m_cProcRef)
    ReDim Preserve m_procRef_indProc(m_cProcRef)
    ReDim Preserve m_procRef_tpAto(m_cProcRef)
    m_procRef_nProc(m_cProcRef) = nProc
    m_procRef_indProc(m_cProcRef) = indProc
    m_procRef_tpAto(m_cProcRef) = tpAto
End Sub

Public Property Get obsCont_count() As Long
    obsCont_count = m_cObsCont + 1
End Property

Public Property Get obsCont_xCampo(ByVal idx As Long) As String
    obsCont_xCampo = m_obsCont_xCampo(idx)
End Property

Public Property Get obsCont_xTexto(ByVal idx As Long) As String
    obsCont_xTexto = m_obsCont_xTexto(idx)
End Property

Public Property Get obsFisco_count() As Long
    obsFisco_count = m_cObsFisco + 1
End Property

Public Property Get obsFisco_xCampo(ByVal idx As Long) As String
    obsFisco_xCampo = m_obsFisco_xCampo(idx)
End Property

Public Property Get obsFisco_xTexto(ByVal idx As Long) As String
    obsFisco_xTexto = m_obsFisco_xTexto(idx)
End Property

Public Property Get procRef_count() As Long
    procRef_count = m_cProcRef + 1
End Property

Public Property Get procRef_nProc(ByVal idx As Long) As String
    procRef_nProc = m_procRef_nProc(idx)
End Property

Public Property Get procRef_indProc(ByVal idx As Long) As String
    procRef_indProc = m_procRef_indProc(idx)
End Property

Public Property Get procRef_tpAto(ByVal idx As Long) As String
    procRef_tpAto = m_procRef_tpAto(idx)
End Property

Public Sub Clear()
    m_cObsCont = -1
    m_cObsFisco = -1
    m_cProcRef = -1
    Erase m_obsCont_xCampo
    Erase m_obsCont_xTexto
    Erase m_obsFisco_xCampo
    Erase m_obsFisco_xTexto
    Erase m_procRef_nProc
    Erase m_procRef_indProc
    Erase m_procRef_tpAto
End Sub

Private Sub Class_Initialize()
    m_cObsCont = -1
    m_cObsFisco = -1
    m_cProcRef = -1
End Sub
