VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum LogLevel
    LOG_DEBUG = 0
    LOG_INFO = 1
    LOG_WARN = 2
    LOG_ERROR = 3
    LOG_FATAL = 4
    LOG_NONE = 5
End Enum

Private m_Level As LogLevel
Private m_OutputDebug As Boolean
Private m_OutputFile As Boolean
Private m_FilePath As String
Private m_FileNum As Integer
Private m_OutputMsgBox As Boolean
Private m_MsgBoxMinLevel As LogLevel
Private m_Listeners As Collection
Private m_LevelNames() As String

Private Sub Class_Initialize()
    m_Level = LOG_DEBUG
    m_OutputDebug = True
    m_OutputFile = False
    m_OutputMsgBox = True
    m_MsgBoxMinLevel = LOG_ERROR
    Set m_Listeners = New Collection
    m_LevelNames = Split("DEBUG,INFO,WARN,ERROR,FATAL", ",")
End Sub

Private Sub Class_Terminate()
    CloseFile
End Sub

Public Property Get LogLevel() As LogLevel
    LogLevel = m_Level
End Property

Public Property Let LogLevel(ByVal vNewValue As LogLevel)
    m_Level = vNewValue
End Property

Public Property Get OutputToDebug() As Boolean
    OutputToDebug = m_OutputDebug
End Property

Public Property Let OutputToDebug(ByVal vNewValue As Boolean)
    m_OutputDebug = vNewValue
End Property

Public Property Get OutputToFile() As Boolean
    OutputToFile = m_OutputFile
End Property

Public Property Get OutputToMsgBox() As Boolean
    OutputToMsgBox = m_OutputMsgBox
End Property

Public Property Let OutputToMsgBox(ByVal vNewValue As Boolean)
    m_OutputMsgBox = vNewValue
End Property

Public Property Get MsgBoxMinLevel() As LogLevel
    MsgBoxMinLevel = m_MsgBoxMinLevel
End Property

Public Property Let MsgBoxMinLevel(ByVal vNewValue As LogLevel)
    m_MsgBoxMinLevel = vNewValue
End Property

Public Property Get FilePath() As String
    FilePath = m_FilePath
End Property

Public Sub SetFile(ByVal Path As String)
    CloseFile
    m_FilePath = Path
    m_OutputFile = True
    OpenFile
End Sub

Public Sub CloseLogFile()
    CloseFile
    m_OutputFile = False
End Sub

Public Sub AddListener(ByVal Listener As ILogListener)
    m_Listeners.Add Listener
End Sub

Public Sub RemoveListener(ByVal Listener As ILogListener)
Dim i As Long
    For i = m_Listeners.Count To 1 Step -1
        If m_Listeners(i) Is Listener Then
            m_Listeners.Remove i
            Exit For
        End If
    Next
End Sub

Public Sub ClearListeners()
    Set m_Listeners = New Collection
End Sub

Public Sub LogDebug(ByVal Msg As String)
    LogMessage LOG_DEBUG, Msg
End Sub

Public Sub LogInfo(ByVal Msg As String)
    LogMessage LOG_INFO, Msg
End Sub

Public Sub LogWarn(ByVal Msg As String)
    LogMessage LOG_WARN, Msg
End Sub

Public Sub LogError(ByVal Msg As String)
    LogMessage LOG_ERROR, Msg
End Sub

Public Sub LogFatal(ByVal Msg As String)
    LogMessage LOG_FATAL, Msg
End Sub

Private Sub LogMessage(ByVal Level As LogLevel, ByVal Msg As String)
    If Level < m_Level Then Exit Sub

    Dim Prefix As String
    Prefix = "[" & m_LevelNames(Level) & "] "
    Dim FullMsg As String
    FullMsg = Prefix & Msg

    If m_OutputDebug Then
        Debug.Print FullMsg
    End If

    If m_OutputFile And Len(m_FilePath) > 0 Then
        WriteToFile FullMsg
    End If

    If m_OutputMsgBox And Level >= m_MsgBoxMinLevel Then
        MsgBox FullMsg, vbOKOnly, "Logger"
    End If

    Dim Listener As ILogListener
    For Each Listener In m_Listeners
        Listener.LogReceived Level, FullMsg
    Next
End Sub

Private Sub OpenFile()
    If Len(m_FilePath) = 0 Then Exit Sub
    m_FileNum = FreeFile
    Open m_FilePath For Append As m_FileNum
End Sub

Private Sub CloseFile()
    If m_FileNum > 0 Then
        Close m_FileNum
        m_FileNum = 0
    End If
End Sub

Private Sub WriteToFile(ByVal Msg As String)
    If m_FileNum = 0 Then
        OpenFile
    End If
    If m_FileNum > 0 Then
        Print #m_FileNum, "[" & Now & "] " & Msg
    End If
End Sub