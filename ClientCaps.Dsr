VERSION 5.00
Begin {17016CEE-E118-11D0-94B8-00A0C91110ED} WebClass1 
   ClientHeight    =   5445
   ClientLeft      =   750
   ClientTop       =   1425
   ClientWidth     =   7320
   _ExtentX        =   12912
   _ExtentY        =   9604
   MajorVersion    =   0
   MinorVersion    =   8
   StateManagementType=   1
   ASPFileName     =   ""
   DIID_WebClass   =   "{12CBA1F6-9056-11D1-8544-00A024A55AB0}"
   DIID_WebClassEvents=   "{12CBA1F5-9056-11D1-8544-00A024A55AB0}"
   TypeInfoCookie  =   0
   BeginProperty WebItems {193556CD-4486-11D1-9C70-00C04FB987DF} 
      WebItemCount    =   0
   EndProperty
   NameInURL       =   "WebClass1"
End
Attribute VB_Name = "WebClass1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Public Sub DisplayServerVariables(Optional whichVariable As Variant)
    Dim index As Integer
    Response.Write "<TABLE>"
    With Request
    If IsMissing(whichVariable) Then
         For index = 1 To .ServerVariables.Count
            Response.Write "<TR><TD>"
            Response.Write Request.ServerVariables.Key(index)
            Response.Write "</TD><TD>"
            Response.Write Request.ServerVariables.Item(index)
            Response.Write "</TD></TR>"
         Next
    Else
Response.Write "<TR><TD>"
Response.Write whichVariable
Response.Write "</TD><TD>"
Response.Write Request.ServerVariables(whichVariable)
Response.Write "</TD></TR>"
Response.Write "</TABLE>"
    End If
Response.Write "</TABLE>"
    End With
End Sub

Public Sub WebClass_Start()
    DisplayServerVariables
End Sub

Public Function GetBrowserID() As String
    ' This will return the abbreviated form of the browser name
    ' such as IE for Internet Explorer or NS for Netscape Navigator
    Dim objBrowser As Object
    Set objBrowser = Server.CreateObject(MSWC.BrowserType)
    GetBrowserID = CStr(objBrowser.Browser)
End Function

Public Function GetBrowserVersion() As Single
    ' This will return the version number of the browser as type
    ' Single.
    Dim objBrowser As Object
    Set objBrowser = Server.CreateObject(MSWC.BrowserType)
    GetBrowserVersion = CSng(objBrowser.Version)
End Function

