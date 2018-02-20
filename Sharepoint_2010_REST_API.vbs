Option Explicit
' ---------------------------------------------------------------------------------------------------------
' MANUEL OPDATERING AF INTRANET
' ---------------------------------------------------------------------------------------------------------
' Author:		KeikoWare, Kim Eik Ortvald
' Create date:	2018-02-15
' Description:	Opdatering af SharepointLister via SharePoint 2010 REST API, 
' 				Afvikles maneult og opdaterer ALLE listerNavnene i variablen arrSharepointLists
' ---------------------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------------------------------------------
'	list update logic
' ---------------------------------------------------------------------------------------------------------
'	input list name
'	getListItems
'	getNewListData
'	If getNewData = Success AND getListItems = Success Then
'		foreach(item){
'			deleteListItem(item)
'		}
'		foreach(newItem){
'			putListItem(newItem)
'		}
'	End If
' ---------------------------------------------------------------------------------------------------------

' Initialisering af scriptet
Dim con, usr, pwd, listJSONData
Dim xmlDoc, objNode, objNodeList, objIdNode, strId, strJSONdata, arrJSONdata
Dim strWebUrl, strListName, arrSharepointLists, strListItem, arrUrlData
Dim strError, strSuccess ' Globale variables which indicate succes or error on every webservice call
Dim intListCnt, intTotalCnt, intListCntDel, intListCntErr, intListCntIns

' ---------------------------------------------------------------------------------------------------------
' Variablen med alle de SharePoint lister der skal opdateres!
' ---------------------------------------------------------------------------------------------------------

arrSharepointLists = Array( "http://sharepoint.company.net/sites/InternalIntranet/SubSite/_vti_bin/listdata.svc/AllEmployees", _ 
							"http://sharepoint.company.net/sites/InternalIntranet/SubSite/_vti_bin/listdata.svc/Cars", _
							"http://sharepoint.company.net/sites/InternalIntranet/SubSite/_vti_bin/listdata.svc/Absence")
' ---------------------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------------------------------------------
' Sharepoint credentials - must be editor of the sharepoint lists
' ---------------------------------------------------------------------------------------------------------
Dim useWindowsLogin
useWindowsLogin = true
usr = Base64Decode("a2lvcg==")
pwd = Base64Decode("SzMxazBXYXJlQDEyMDE4")
' ---------------------------------------------------------------------------------------------------------

' Her afgøres det om der skal skrives til consollen? 
If InStr(Ucase(WScript.FullName), "CSCRIPT.EXE") Then 
	con = true
Else
	con = false
End If
' Herunder itereres over listen med SharepointLister der skal opdateres
intTotalCnt = 0
For Each strListItem in arrSharepointLists
	
	arrUrlData = split(strListItem,"_vti_bin/listdata.svc/")
	strlistName = arrUrlData(1)
	strWebUrl   = arrUrlData(0)
	
	If con = true then 	wscript.Echo "Henter Data fra " + strlistName
	intListCnt = 0
	intListCntDel = 0
	intListCntErr = 0
	listJSONData = ""
	If getSQLdata(strListName) then	
		If con = true then wscript.echo "JSON data downloaded succesfull"
		If getArrayOfListItems(strWebUrl, strListName) then
			If con = true then wscript.echo "Sharepoint XML downloaded succesfull"
			Set xmlDoc = CreateObject("MSXML2.DOMDocument") 
			xmlDoc.loadXML(strSuccess)
			Set objNodeList = xmlDoc.SelectNodes("/feed/entry/content/m:properties/d:Id")
			If con = true then wscript.echo "Sletter " & (objNodeList.length) & " gamle poster ..."

			For Each objNode In objNodeList
				intListCnt = intListCnt + 1
				strId = " - no child nodes found - "
				If not objNode is nothing Then 
					strId = objNode.text 
					If deleteListItem(strWebUrl, strListName, strId) then
						intListCntDel = intListCntDel + 1
					Else
						intListCntErr = intListCntErr + 1
					End If
				End If
			Next
			If con = true then wscript.echo "Deleted Sharepoint list items: " & intListCnt & " (deleted: " & intListCntDel & " - failed: " & intListCntErr & ")"

			intListCnt = 0
			intListCntIns = 0
			intListCntErr = 0
			listJSONData = replace(replace(listJSONData,"[{",""),"}]","")
			arrJSONdata = split(listJSONData,"},{")
			If con = true then wscript.echo "Tilføjer " & (ubound(arrJSONdata) + 1) & " nye poster ..."
			For Each strJSONData in arrJSONdata
				intListCnt = intListCnt + 1
'				If con = true then wscript.echo  "Tilføjer ... {" & strJSONData & "}"
				If createListItem(strWebUrl, strListName, "{" & strJSONData & "}") Then
					intListCntIns = intListCntIns + 1
				Else
					intListCntErr = intListCntErr + 1
					wscript.echo "{" & strJSONData & "}"
					wscript.quit
				End If
			Next
			If con = true then wscript.echo "Inserted Sharepoint list items: " & intListCnt & " (inserted: " & intListCntIns & " - failed: " & intListCntErr & ")"
		Else
			If con = true then wscript.echo strError
		End If
	Else
		If con = true then wscript.echo strError
	End If
	If con = true then wscript.echo " ---- "
Next
wscript.echo "All Done"
wscript.Quit()

' *****************************************************************************
' ***************************** SCRIPT END ************************************
' *****************************************************************************

' *****************************************************************************
'                       SHAREPOINT UPDATING FUNCTIONS
' *****************************************************************************

' getArrayOfLists
Function getArrayOfLists(strWebUrl)
    Dim url, http, objResponse
	url = webUrl + "/_vti_bin/listdata.svc"
	Set http =  CreateObject("WinHttp.WinHttpRequest.5.1")
	http.open "GET", url, false
    If useWindowsLogin Then 
		' Use windows authentication
		http.SetAutoLogonPolicy 0
	Else
		' Use UserName and PassWord
		http.SetCredentials usr, pwd, 0
	End If
	http.send
	If http.status = 200 then
		strError = ""
		strSuccess = http.responseTEXT
		getArrayOfLists = true
	Else
		strError = http.status
		getArrayOfLists = false
	End if
	Set http = Nothing
End Function

' getArrayOfListItems
Function getArrayOfListItems(webUrl, listName)
    Dim url, http, objResponse
	url = webUrl + "/_vti_bin/listdata.svc/" + listName
'	Set http = createObject("Microsoft.XMLHTTP")
	Set http =  CreateObject("WinHttp.WinHttpRequest.5.1")
	http.open "GET", url, false
    If useWindowsLogin Then 
		' Use windows authentication
		http.SetAutoLogonPolicy 0
	Else
		' Use UserName and PassWord
		http.SetCredentials usr, pwd, 0
	End If
	http.send
	If http.status = 200 then
		strError = ""
		strSuccess = http.responseTEXT
		getArrayOfListItems = true
	Else
		strError = http.status
		getArrayOfListItems = false
	End if
	Set http = Nothing
End Function

' getListItemById
Function getListItemById(webUrl, listName, itemId)
    Dim url, http, objResponse
	url = webUrl & "/_vti_bin/listdata.svc/" & listName & "(" & itemId & ")"
	Set http =  CreateObject("WinHttp.WinHttpRequest.5.1")
	http.open "GET", url, false
    If useWindowsLogin Then 
		' Use windows authentication
		http.SetAutoLogonPolicy 0
	Else
		' Use UserName and PassWord
		http.SetCredentials usr, pwd, 0
	End If
	http.send
	If http.status = 200 then
		strError = ""
		strSuccess = http.responseTEXT
		getListItemById = true
	Else
		strError = http.status & " " & url
		getListItemById = false
	End if
	Set http = Nothing
End Function

' updateListItemById
Function updateListItemById(webUrl, listName, itemId, itemProperties)
	If getListItemById(webUrl, listName, itemId) = true Then
		Dim url, http, objResponse
		url = webUrl & "/_vti_bin/listdata.svc/" & listName & "(" & itemId & ")"
		Set http =  CreateObject("WinHttp.WinHttpRequest.5.1")
		http.open "POST", url, false
		http.setRequestHeader "X-HTTP-Method", "MERGE"
        http.setRequestHeader "If-Match", "*"
		If useWindowsLogin Then 
			' Use windows authentication
			http.SetAutoLogonPolicy 0
		Else
			' Use UserName and PassWord
			http.SetCredentials usr, pwd, 0
		End If
		http.send itemProperties
		If http.status = 200 then
			strError = ""
			strSuccess = http.responseTEXT
			updateListItemById = true
		Else
			strError = http.status & " " & url
			updateListItemById = false
		End if
		Set http = Nothing
	End if
End Function

' createListItem
Function createListItem(webUrl, listName, itemProperties)
	Dim url, http, objResponse
	url = webUrl & "/_vti_bin/listdata.svc/" & listName
	Set http =  CreateObject("WinHttp.WinHttpRequest.5.1")
	http.open "POST", url, false
	http.setRequestHeader "Content-type", "application/json"
    If useWindowsLogin Then 
		' Use windows authentication
		http.SetAutoLogonPolicy 0
	Else
		' Use UserName and PassWord
		http.SetCredentials usr, pwd, 0
	End If
	http.send itemProperties 
	If http.status = 201 then
		strError = ""
		strSuccess = http.responseTEXT
		createListItem = true
'		wscript.echo "Data oprettet korrekt"
	Else
		If http.status < 500 Then 
			strError = http.responseTEXT
		Else
			strError = "Items could not be created (" & http.status & ")" & strError
		End If
		createListItem = false
		wscript.echo "Data IKKE sendt: " & strError
	End if
	Set http = Nothing
End Function

' deleteListItem
Function deleteListItem(webUrl, listName, itemId)
	If getListItemById(webUrl, listName, itemId) = true Then
		Dim url, http, objResponse
		url = webUrl & "/_vti_bin/listdata.svc/" & listName & "(" & itemId & ")"
		strError = ""
		Set http =  CreateObject("WinHttp.WinHttpRequest.5.1")
		http.open "POST", url, false
		http.setRequestHeader "X-HTTP-Method", "DELETE"
        http.setRequestHeader "If-Match", "*"
		If useWindowsLogin Then 
			' Use windows authentication
			http.SetAutoLogonPolicy 0
		Else
			' Use UserName and PassWord
			http.SetCredentials usr, pwd, 0
		End If
		http.send "" 
		If http.status = 204 then
			strSuccess = http.responseTEXT
			deleteListItem = true
		Else
			If http.status < 500 Then strError = http.responseTEXT
			strError = "Item could not be deleted " & itemId & " (" & http.status & ")" & strError
			deleteListItem = false
		End if
		Set http = Nothing
	Else
		strError = "Item does not exist " & itemId
		deleteListItem = false
	End If
End Function

Function getSQLData(listName)
	Dim url, http, objResponse
	url = "http://phb2b01.aarsleff.com/intranet/spList/data/?list=" + strListName
	Set http =  CreateObject("WinHttp.WinHttpRequest.5.1")
	http.open "GET", url, false
    If useWindowsLogin Then 
		' Use windows authentication
		http.SetAutoLogonPolicy 0
	Else
		' Use UserName and PassWord
		http.SetCredentials usr, pwd, 0
	End If
	http.send
	listJSONData = http.responseTEXT
	If http.status = 200 then
		strError = ""
		getSQLData = true
	Else
		If http.status < 500 Then strError = http.responseTEXT
		strError = "JSON data is not available (" & http.status & ")" & strError
		getSQLData = false
	End if
	Set http = Nothing
End Function

' *****************************************************************************
' 									HELP FUNCTIONS
' *****************************************************************************
Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function Base64Decode(ByVal vCode)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.text = vCode
    Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
    Set oNode = Nothing
    Set oXML = Nothing
End Function

'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string 
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To text/string
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the output text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And get text/string data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function