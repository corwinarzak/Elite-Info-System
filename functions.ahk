Encrypt(mystring)
{

}

Decrypt(allstr)
{

}

URLDownloadToVar(url) {
    if url <> ""
    {
        hObject:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
        hObject.Open("GET", url)
        hObject.Send()
        response := hObject.ResponseText
        hObject :=
        return response
    }
    else
    {
        ErrorGui("ERROR! Exiting...")
        ExitApp
    }
}

showAndBreak(sthk, ophk, wbhk, awhk, sec=5) 
{
    msstart := A_TickCount
    Loop
    {
        if (GetKeyState(sthk, "p") or GetKeyState(ophk, "p") or GetKeyState(wbhk, "p") or GetKeyState(awhk, "p"))
            break
        now := A_TickCount-msstart
        msec := (sec * 1000)
        if (now > msec)
            break
    }
    
}

showAndBreakNoKey(sec=5) 
{
    msstart := A_TickCount
    Loop
    {
        now := A_TickCount-msstart
        msec := (sec * 1000)
        if (now > msec)
            break
    }
    
}

showAndBreakOnKey(sthk, ophk, wbhk, awhk) 
{
    Loop
    {
        if (GetKeyState(sthk, "p") or GetKeyState(ophk, "p") or GetKeyState(wbhk, "p") or GetKeyState(awhk, "p"))
            break
    }
    
}

ErrorGui(string="Error while processing. Try again.")
{
    Gui, Destroy
            
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 40,20
    Gui, font, s18 cwhite
    Gui, Add, Text, center cWhite, %string%
    WinSet, TransColor, EE0000 175
    Gui, show, autosize

    showAndBreak(status_hotkey, operations_hotkey, wing_enter_hotkey, all_wings_hotkey)
    
    Gui, Destroy
    Suspend, Off
    Return
}

GetAccessCode(client_id, client_secret, refresh_token) 
{
    StringReplace, client_id, client_id, %A_SPACE%,, All
    StringReplace, client_secret, client_secret, %A_SPACE%,, All
    StringReplace, refresh_token, refresh_token, %A_SPACE%,, All
    aURL := "https://www.googleapis.com/oauth2/v3/token"
    aPostData := "client_id=" client_id "&client_secret=" client_secret "&refresh_token=" refresh_token "&grant_type=refresh_token"
    oHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    oHTTP.Open("POST", aURL , False)
    oHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    oHTTP.Send(aPostData)
    data1 := oHTTP.ResponseText
    parsedAToken := JSON.Load(data1, true)
    parsedAccessToken := parsedAToken.access_token
    oHTTP :=
    return parsedAccessToken
}

userLoginEnjin(website, email, pass)
{
	
	Random, rand , , 99999
	
	result := []
	
	request = 
	(
	{ 
		"jsonrpc": "2.0", 
		"id": "%rand%", 
		"method": "User.login", 
		"params": { 
			"email": "%email%", 
			"password": "%pass%" 
		} 
	}
	)

	apiURL := website "/api/v1/api.php"
	apiHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	apiHTTP.Open("POST", apiURL, false)
	apiHTTP.SetRequestHeader("Content-Type", "application/json")
	apiHTTP.Send(request)
	content := apiHTTP.ResponseText
	contentdata := JSON.Load(content, true)
    result["userid"] := contentdata.result.user_id
	result["username"] := contentdata.result.username
	result["session"] := contentdata.result.session_id
	apiHTTP :=
	return, result
}

userLoginShivtr(website, email, pass)
{
		
	result := []
	
	request = 
	(
	{ 
		"user": { 
			"email": "%email%", 
			"password": "%pass%" 
		} 
	}
	)

	apiURL := website "/users/sign_in.json"
	apiHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	apiHTTP.Open("POST", apiURL, false)
	apiHTTP.SetRequestHeader("Content-Type", "application/json")
	apiHTTP.Send(request)
	content := apiHTTP.ResponseText
	contentdata := JSON.Load(content, true)
    result["userid"] := contentdata.user_session.id
	result["username"] := contentdata.user_session.name
	apiHTTP :=
	return, result
}

checkSession(website, sessionid)
{
	
	Random, rand , , 99999
	
	result := []
	
	request = 
	(
	{ 
		"jsonrpc": "2.0", 
		"id": "%rand%", 
		"method": "User.checkSession", 
		"params": { 
			"session_id": "%sessionid%" 
		} 
	}
	)

	apiURL := website "/api/v1/api.php"
	apiHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	apiHTTP.Open("POST", apiURL, false)
	apiHTTP.SetRequestHeader("Content-Type", "application/json")
	apiHTTP.Send(request)
	content := apiHTTP.ResponseText
	contentdata := JSON.Load(content, true)
	result["identity"] := contentdata.result.hasIdentity
    result["username"] := contentdata.result.username
	apiHTTP :=
	return, result
}

formatstamp(timestamp)
{
    current := 1970
    current += timestamp, Seconds
    FormatTime, formattedtime, %current%, yyyy-MMM-dd HH:mm:ss
    return formattedtime
}

FormatSeconds(NumberOfSeconds)  
{
    time = 19700101000000 
    time += %NumberOfSeconds%, seconds
    FormatTime, mmnoss, %time%, mm
    return NumberOfSeconds//3600 ":" mmnoss  
}

loadXML(ByRef data)
{
  o := ComObjCreate("MSXML2.DOMDocument.6.0")
  o.async := false
  o.loadXML(data)
  return o
}

uriEncode(str)
{ 
	b_Format := A_FormatInteger
	data := ""
	SetFormat,Integer,H
	Loop,Parse,str
		if ((Asc(A_LoopField)>0x7f) || (Asc(A_LoopField)<0x30) || (asc(A_LoopField)=0x3d))
			data .= "%" . ((StrLen(c:=SubStr(ASC(A_LoopField),3))<2) ? "0" . c : c)
		Else
			data .= A_LoopField
	SetFormat,Integer,%b_format%
	return data
}

uriDecode(str)
{ 
	Loop,Parse,str,`%
		txt := (A_Index=1) ? A_LoopField : txt chr("0x" substr(A_LoopField,1,2)) SubStr(A_LoopField,3)
	return txt
}

getWorksheets(SheetKey, accessToken)
{
    WS_url := "https://spreadsheets.google.com/feeds/worksheets/" SheetKey "/private/full?access_token=" accessToken
    fileName := URLDownloadToVar(WS_url)
    StringReplace, fileName, fileName, <feed xmlns='http://www.w3.org/2005/Atom' xmlns:openSearch='http://a9.com/-/spec/opensearchrss/1.0/' xmlns:gs='http://schemas.google.com/spreadsheets/2006'>, <feed xmlns:openSearch='http://a9.com/-/spec/opensearchrss/1.0/' xmlns:gs='http://schemas.google.com/spreadsheets/2006'>
    xmlObj := loadXML(fileName)
    n := xmlObj.SelectNodes("//entry/link[1]/@href")

    xmlWorksheets := []
    while, Node := n.item[A_Index-1]
    {   
        xmlWorksheets[A_index] := Node.text
    }
    return xmlWorksheets
}

getWSnames(Sheetkey, accessToken)
{   
    WS_url := "https://spreadsheets.google.com/feeds/worksheets/" SheetKey "/private/full?access_token=" accessToken
    fileName := URLDownloadToVar(WS_url)
    StringReplace, fileName, fileName, <feed xmlns='http://www.w3.org/2005/Atom' xmlns:openSearch='http://a9.com/-/spec/opensearchrss/1.0/' xmlns:gs='http://schemas.google.com/spreadsheets/2006'>, <feed xmlns:openSearch='http://a9.com/-/spec/opensearchrss/1.0/' xmlns:gs='http://schemas.google.com/spreadsheets/2006'>

    xmlObj := loadXML(fileName)
    n0 := xmlObj.SelectNodes("./node()//entry/title")
    
    WSnames := []
    while, Node := n0.item[A_Index-1]
    {   
        WSnames[A_index] := Node.text
    }
    return WSnames
}

searchCMDR(Sheetkey, accessToken, name:="")
{
    
        Worksheets := getWorksheets(Sheetkey, accessToken)
        
        statusres := []
        result := []
        idx := 0
        colcnt := 0
        StringUpper, uname, name
        cmdnm := uriEncode(uname)
        loop, % Worksheets.length()
        {   
            idx++
            WS_url2 := Worksheets[A_index] "?access_token=" accessToken "&sq=name=%22" cmdnm "%22"
            fileName2 := URLDownloadToVar(WS_url2)
            StringReplace, fileName2, fileName2, <feed xmlns='http://www.w3.org/2005/Atom' xmlns:openSearch='http://a9.com/-/spec/opensearchrss/1.0/' xmlns:gsx='http://schemas.google.com/spreadsheets/2006/extended'>, <feed xmlns:openSearch='http://a9.com/-/spec/opensearchrss/1.0/' xmlns:gsx='http://schemas.google.com/spreadsheets/2006/extended'>
            
            xmlObj2 := loadXML(fileName2)
            
            n1 := xmlObj2.SelectNodes("./node()/title")
            if n1.length() > 0
            {
                while, Node := n1.item[A_Index-1]
                {   
                    statusres[A_Index] := Node.text
                }
            }
            
            n2 := xmlObj2.SelectNodes("//entry/*")
            
            if n2.length() > 0
            {
                while, Node := n2.item[A_Index-1]
                {   
                    if statusres[A_Index]
                    {
                        result["status"] := statusres[A_Index]
                    }
                    result["index"] := idx
                    if InStr(Node.nodeName, "gsx:")
                    {
                        colcnt++
                        inputnodname := Node.nodeName
                        StringTrimLeft, nodName, inputnodname, 4
                        result["column"colcnt]:= nodName
                        result["content"colcnt] := Node.text
                    }
                    
                }   
            }
            
        }
        result["colcnt"] := colcnt
        
    return result
}

getWorksheetsData(Sheetkey, accessToken)
{
    
    Worksheets := getWorksheets(Sheetkey, accessToken)
    result := [] 
    loop, % Worksheets.length()
    {   
        WS_url := Worksheets[A_index] "?access_token=" accessToken
        fileData := URLDownloadToVar(WS_url)
        StringReplace, fileData, fileData, <feed xmlns='http://www.w3.org/2005/Atom' xmlns:openSearch='http://a9.com/-/spec/opensearchrss/1.0/' xmlns:gsx='http://schemas.google.com/spreadsheets/2006/extended'>, <feed xmlns:openSearch='http://a9.com/-/spec/opensearchrss/1.0/' xmlns:gsx='http://schemas.google.com/spreadsheets/2006/extended'>
        result[A_index] := fileData
    }
        
    return result
}

getWorksheetsTitleNum(filevar)
{
    xmlObj2 := loadXML(filevar)
    res := []    
    n1 := xmlObj2.SelectNodes("./node()/title")
    if n1.length() > 0
    {
        while, Node := n1.item[A_Index-1]
        {   
            res["title"] := Node.text
        }
    }
    n2 := xmlObj2.SelectNodes("./node()/entry")
    if n2.length() > 0
    {
        while, Node := n2.item[A_Index-1]
        {   
            res["entries"] := A_index
        }
    }
    return res
}

getOpEntries(filevar, loopidx)
{
    xmlObj2 := loadXML(filevar)
    res := []
    res[loopidx] := []
    res[loopidx]["label"] := []
    res[loopidx]["content"] := []
    idx := 0
    n1 := xmlObj2.SelectNodes("./node()/entry[" loopidx "]/*")
    if n1.length() > 0
    {
        while Node := n1.item[A_Index-1]
        {   
            if InStr(Node.nodeName, "gsx:")
            {
                idx++
                inputnodname := Node.nodeName
                StringTrimLeft, nodName, inputnodname, 4
                res[loopidx]["label"][idx] := nodName
                res[loopidx]["content"][idx] := Node.text
                res[loopidx]["fields"] := idx
            }
        }
    }
    return res
}

getWingsEntries(filevar, loopidx)
{
    xmlObj2 := loadXML(filevar)
    res := []
    res[loopidx] := []
    res[loopidx]["label"] := []
    res[loopidx]["content"] := []
    idx := 0
    
    n := xmlObj2.SelectNodes("./node()/entry[" loopidx "]/link[2]/@href")
    while, Node := n.item[A_Index-1]
    {   
        res[loopidx]["edit"] := Node.text
    }
    
    n1 := xmlObj2.SelectNodes("./node()/entry[" loopidx "]/*")
    if n1.length() > 0
    {
        while Node := n1.item[A_Index-1]
        {   
            if InStr(Node.nodeName, "gsx:")
            {
                idx++
                inputnodname := Node.nodeName
                StringTrimLeft, nodName, inputnodname, 4
                res[loopidx]["label"][idx] := nodName
                res[loopidx]["content"][idx] := Node.text
                res[loopidx]["fields"] := idx
            }
        }
    }
    return res
}