Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mp_ValueSeperator As String
Private mp_KeyValueSeperator As String

Private m_Values As VBA.Collection



' ___ Initialize/Terminate ___



Private Sub Class_Initialize()
   Set m_Values = New VBA.Collection
   mp_ValueSeperator = ";"
   mp_KeyValueSeperator = "="
End Sub

Private Sub Class_Terminate()
   Set m_Values = Nothing
End Sub



' ___ Public Properties ___



Public Property Get Self() As NamedValues
   Set Self = Me
End Property

Public Property Let ValueSeperator(ByVal Value As String)
   If Len(Value) > 0 Then
      mp_ValueSeperator = Value
   End If
End Property
Public Property Get ValueSeperator() As String
   ValueSeperator = mp_ValueSeperator
End Property

Public Property Let KeyValueSeperator(ByVal Value As String)
   If Len(Value) > 0 Then
      mp_KeyValueSeperator = Value
   End If
End Property
Public Property Get KeyValueSeperator() As String
   KeyValueSeperator = mp_KeyValueSeperator
End Property

Public Property Get Item(ByVal Key As String) As String
   Item = m_Values(Key).Value
End Property

Public Property Get ItemOrDefault(ByVal Key As String, Optional ByVal Default As Variant = Null) As Variant
   If Exists(Key) Then
      ItemOrDefault = m_Values(Key).Value
   Else
      ItemOrDefault = Default
   End If
End Property

Public Property Let AsString(ByVal NewString As String)
   Dim PosLastValueSep As Integer
   Dim PosValueSep As Integer
   Dim PosKeyValueSep As Integer
   Dim KeyStart As Integer
   Dim kvp As KeyValuePair
   
   Set m_Values = Nothing
   Set m_Values = New Collection
   
   PosLastValueSep = 1 - Len(mp_ValueSeperator)
   PosKeyValueSep = InStr(1, NewString, mp_KeyValueSeperator)
   Do While PosKeyValueSep > 0
      Set kvp = New KeyValuePair
      KeyStart = PosLastValueSep + Len(mp_ValueSeperator)
      kvp.Key = Mid$(NewString, KeyStart, PosKeyValueSep - KeyStart)
      PosValueSep = InStr(PosKeyValueSep, NewString, mp_ValueSeperator)
      If PosValueSep > 0 Then
         kvp.Value = Mid$(NewString, PosKeyValueSep + Len(mp_KeyValueSeperator), PosValueSep - PosKeyValueSep - Len(mp_KeyValueSeperator))
      Else
         kvp.Value = Mid$(NewString, PosKeyValueSep + Len(mp_KeyValueSeperator))
      End If
      m_Values.Add kvp, kvp.Key
      Set kvp = Nothing
      If PosValueSep > 0 Then
         PosLastValueSep = PosValueSep
         PosKeyValueSep = InStr(PosLastValueSep + 1, NewString, mp_KeyValueSeperator)
      Else
         PosKeyValueSep = 0
      End If
   Loop
End Property
Public Property Get AsString() As String
   Dim strTemp As String
   Dim kvp As KeyValuePair
   
   For Each kvp In m_Values
      strTemp = strTemp & kvp.Key & mp_KeyValueSeperator & kvp.Value & mp_ValueSeperator
   Next kvp
   If Len(strTemp) > 0 Then
      AsString = Left$(strTemp, Len(strTemp) - Len(mp_ValueSeperator))
   Else
      AsString = ""
   End If
End Property



' ___ Public Methods ___



Public Sub Add(ByVal Key As String, ByVal Value As String)
   Dim vak As KeyValuePair
   
   Set vak = New KeyValuePair
   vak.Key = Key
   vak.Value = Value
   m_Values.Add vak, vak.Key
   Set vak = Nothing
End Sub

Public Function Exists(ByVal Key As String) As Boolean
   Dim strDummy As String
   
   On Error Resume Next
   strDummy = m_Values(Key).Value
   Exists = (Err.Number = 0)
   Err.Clear
End Function