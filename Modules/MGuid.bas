Attribute VB_Name = "MGuid"
Option Explicit

Private Type GUID_TYPE
    lData1 As Long
    lData2 As Long
    lData3 As Long
    bData4(7) As Byte
End Type

Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (oGuid As GUID_TYPE) As LongPtr
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (oGuid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr

'Формируем Guid
Public Property Get Generate() As String
    Dim oGuid       As GUID_TYPE
    Dim sGuid       As String
    Dim retValue    As LongPtr
    
    Const guidLength As Long = 39
    
    retValue = CoCreateGuid(oGuid)
    
    If retValue = 0 Then
        sGuid = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(oGuid, StrPtr(sGuid), guidLength)
        If retValue = guidLength Then
            Generate = Mid(sGuid, InStr(sGuid, "{") + 1, InStr(sGuid, "}") - InStr(sGuid, "{") - 1)
        End If
    End If
    
End Property
