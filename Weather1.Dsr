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
   TypeInfoCookie  =   5
   BeginProperty WebItems {193556CD-4486-11D1-9C70-00C04FB987DF} 
      WebItemCount    =   2
      BeginProperty WebItem1 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "ReportCommit"
         DISPID          =   1280
         Template        =   "ReportCommit1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{A6302456-4F01-11D2-8BF0-70D350C10000}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "C:\Program Files\Microsoft Visual Studio\VB98\Chapter13\ReportCommit.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   1
            BeginProperty Event0 {193556D3-4486-11D1-9C70-00C04FB987DF} 
               Name            =   "onRequestStates"
               DISPID          =   1280
               Type            =   1
               OriginalHREF    =   ""
               TagType         =   -269488145
               BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
                  AttribCount     =   1
                  BeginProperty Attrib0 {FA6A55FC-458A-11D1-9C71-00C04FB987DF} 
                     TagType         =   0
                     Attribute       =   "HREF"
                     State           =   3
                     TagName         =   "Hyperlink2"
                     OriginalURL     =   "about:ChooseStatePlaceholder"
                     Parent          =   ""
                     Template        =   "ReportCommit"
                     BoundEvent      =   "onRequestStates"
                     BoundItem       =   ""
                     Suffix          =   ""
                     UsesAnonymousName=   -1
                     TagNumber       =   2
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem2 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "WeatherSummary"
         DISPID          =   1281
         Template        =   "Weather1Template1.htm"
         Token           =   "WS:"
         DIID_WebItemEvents=   "{A63023F4-4F01-11D2-8BF0-70D350C10000}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "C:\Program Files\Microsoft Visual Studio\VB98\Chapter13\Weather1Template.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
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

Public State As String
    
Private Sub ReportCommit_onRequestStates()
Dim conn As Connection      'This declares the connection
        Dim rs As Recordset     'This declares the record set
    Dim States As Dictionary
    Dim State As String

    ' Create the initial header HTML and output it to the client
Response.Write "<HTML><HEAD><TITLE>Get New State</TITLE></HEAD>"
    Response.Write "<BODY>"
    Response.Write "<H1>States</H1><P>Click on a state to display all_"
               of it 's cities for editing.</P>"
    ' Flush the output to insure that something's seen on the client
    Response.Flush
    'Open the connection
Set conn = Server.Create("ADODB.Connection")
conn.Open = "WeatherReports"
'Create a new dictionary to hold the list of states
Set States = New Dictionary
'Retrieve a list of all states ordered alphabetically
        Set rs = conn.Execute("SELECT State FROM ReportData ORDER BY State ASC;")
    'Iterate through the list, and remove duplicates.
    rs.MoveFirst
    While Not rs.EOF
        ' Get the state from the current record
    State = rs!State
    ' if the state is not within the dictionary, then add it
    ' otherwise ignore it
        If Not States.Exist(State) Then
            States.Add State, State
        End If
        ' Move to the next record
        rs.MoveNext
    Wend
    ' Connection to database no longer needed.
    rs.Close
    conn.Close
    ' Output a form with a combo box showing all the states,
    ' as well as a submit button.
    Response.Write "<FORM id=stateForm name=stateForm action='" + URLFor(WeatherSubmit) + "' method='POST'>"
    Response.Write "Please select a State:"
Response.Write "<select name=""State"" id=""State"">"
For Each State In States
    ' Retrieve the session level State variable and see if
    ' it belongs to the listed state. Make that option the
    ' selected one if it does.

If Session("State") = State Then
Response.Write "<OPTION value='" + State + "' selected>" + State
        Else
            Response.Write "<OPTION value='" + State + "'>+State"
        End If
    Next
    Response.Write "</SELECT>"
    Response.Write "<INPUT TYPE='SUBMIT'>"
    Response.Write "</FORM>"
    Response.Write "</BODY></HTML>"
End Sub

Private Sub WebClass_Start()
    ' For purposes of illustration, predefine a state. This will change.
State = "Washington"
ReportCommit.WriteTemplate
End Sub


Sub WeatherSummary_ProcessTag(ByVal TagName As String, TagContents As String, SendTags As Boolean)
    Dim conn As Connection      'This declares the connection
    Dim rs As Recordset     'This declares the record set
    Dim fIndex As Integer   'Declaration for a field index
    Dim buffer As String
    Select Case TagName
        Case "WS:STATE"
            TagContents = State
        Case "WS:WEATHERTABLE"
            Set conn = New Connection
            'Set conn = Server.Create("ADODB.Connection")
            conn.Open "Weather"
            Set rs = conn.Execute("SELECT * FROM Weather WHERE State='" + State + "';")
            buffer = ""
            buffer = buffer + "<TABLE>"
            buffer = buffer + "<THEAD>"
            For fIndex = 0 To rs.Fields.Count - 1
                buffer = buffer + "<TH>" + rs.Fields(fIndex).Name + "</TH>"
            Next
            buffer = buffer + "</THEAD><TBODY>"
            rs.MoveFirst
            While Not rs.EOF
                buffer = buffer + "<TR>"
                For fIndex = 0 To rs.Fields.Count - 1
                    buffer = buffer + "<TD>" + CStr(rs.Fields(fIndex)) + "</TD>"
                Next
                buffer = buffer + "</TR>"
                rs.MoveNext
            Wend
            buffer = buffer + "</TBODY></TABLE>"
            rs.Close
            TagContents = buffer
    End Select
End Sub

Private Sub WeatherUpdate(WeatherRecord As Dictionary)
    Dim conn As Connection      'This declares the connection
    Dim rs As Recordset     'This declares the record set

    Set conn = New Connection
    conn.Open "Weather"
    Set rs = conn.Execute("SELECT * FROM ReportData WHERE State='" + WeatherRecord("State") + "' and City='" + WeatherRecord("City") + "';")
    rs!HiF = WeatherRecord("HiF")
    rs!LOF = WeatherRecord("LoF")
    rs!Skies = WeatherRecord("Skies")
    rs!Forecast = WeatherRecord("Forecast")
    rs.Update
    rs.Close
    conn.Close
End Sub

Private Sub ReportCommit_Respond()
    Dim ReportData As Dictionary
    Set ReportData = Session("ReportData")
    WeatherUpdate ReportData
    ReportCommit.WriteTemplate
End Sub


