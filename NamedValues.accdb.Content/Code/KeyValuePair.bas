Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_varValue As Variant
Private m_strKey As String



' ___ Public Members ___



Public Property Let Value(ByVal Value As Variant)
   m_varValue = Value
End Property
Public Property Get Value() As Variant
   Value = m_varValue
End Property

Public Property Let Key(ByVal Value As String)
   m_strKey = Value
End Property
Public Property Get Key() As String
   Key = m_strKey
End Property