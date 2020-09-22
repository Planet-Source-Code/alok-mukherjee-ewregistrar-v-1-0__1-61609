Attribute VB_Name = "mRegisterTypeLib"
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Enum eSYSKIND
   SYS_WIN16 = 0&
   SYS_WIN32 = 1&
   SYS_MAC = 2&
End Enum

Private Declare Function LoadTypeLib Lib "oleaut32.dll" ( _
    pFileName As Byte, pptlib As Object) As Long
Private Declare Function RegisterTypeLib Lib "oleaut32.dll" ( _
    ByVal ptlib As Object, szFullPath As Byte, _
    szHelpFile As Byte) As Long
Private Declare Function UnRegisterTypeLib Lib "oleaut32.dll" ( _
      libID As GUID, _
      ByVal wVerMajor As Integer, _
      ByVal wVerMinor As Integer, _
      ByVal lCID As Long, _
      ByVal tSysKind As eSYSKIND _
   ) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (lpsz As Byte, pclsid As GUID) As Long

Public Function RegTypelib(sLib As String, ByVal bState As Boolean) As Long
Dim suLib() As Byte
Dim errOK As Long
Dim tlb As Object
   If bState Then
      ' Basic automatically translates strings to Unicode Byte arrays
      ' but doesn't null-terminate, so you must do it yourself
      suLib = sLib & vbNullChar
      ' Pass first byte of array
      errOK = LoadTypeLib(suLib(0), tlb)
      If errOK = 0 Then
         errOK = RegisterTypeLib(tlb, suLib(0), 0)
         
      End If
      RegTypelib = errOK
   Else
      Dim cTLI As TypeLibInfo
      Dim tGUID As GUID, sCLSID As String
      Dim iMajor As Integer, iMinor As Integer
      Dim lCID As Long

      Set cTLI = TLI.TypeLibInfoFromFile(sLib)
      sCLSID = cTLI.GUID
      iMajor = cTLI.MajorVersion
      iMinor = cTLI.MinorVersion
      lCID = cTLI.lCID
      Set cTLI = Nothing
      
      suLib = sCLSID & vbNullChar
      errOK = CLSIDFromString(suLib(0), tGUID)
      If errOK = 0 Then
         errOK = UnRegisterTypeLib(tGUID, iMajor, iMinor, lCID, SYS_WIN32)
         RegTypelib = errOK
      End If
      
   End If
   
End Function

Public Function UIRegisterTypeLib(ByVal sLib As String, ByVal bState As Boolean, ByVal bShowMessage As Boolean) As Boolean
Dim errNo As Long
Dim sPre As String

   errNo = RegTypelib(sLib, bState)
   If bShowMessage Then
      If bState Then
         sPre = "Registration "
      Else
         sPre = "Un-Registration "
      End If
      If errNo = 0 Then
        EWregMain.lstJobDone.AddItem sLib & " " & sPre & " Successful."
      Else
        EWregMain.lstJobDone.AddItem sLib & _
          " " & sPre & " Failed ( Error : " & _
          Hex$(errNo) & " )."
      End If
   End If
   UIRegisterTypeLib = (errNo = 0)
End Function
