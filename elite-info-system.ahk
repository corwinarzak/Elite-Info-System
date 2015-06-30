#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

#SingleInstance,Force
#include JSON.ahk
#include functions.ahk
#include CLR.ahk

appversion := "1.0"
appname := "Elite Info System"
appauthor := "Corwin Arzak"

ComObjError(false)

if FileExist( "settings.dll" ) {
    asm := CLR_LoadLibrary("settings.dll")
    obj := CLR_CreateObject(asm, "Settings")
    e_client_id := obj.gcc()
    e_client_secret := obj.gcs()
    e_refresh_token := obj.grt()
    e_org_name := obj.gon()
    e_org_website := obj.gow()
    e_status_sheet_key := obj.gss()
    e_ops_sheet_key := obj.gos()
    e_wings_sheet_key := obj.gws()

    client_id := Decrypt(e_client_id)
    client_secret := Decrypt(e_client_secret)
    refresh_token := Decrypt(e_refresh_token)
    org_name := Decrypt(e_org_name)
    org_website := Decrypt(e_org_website)
    status_sheet_key := Decrypt(e_status_sheet_key)
    ops_sheet_key := Decrypt(e_ops_sheet_key)
    wings_sheet_key := Decrypt(e_wings_sheet_key)
}
else
{
    ErrorGui("setting.dll file not found. Exiting...")
    ExitApp
}

myAccessToken := GetAccessCode(client_id, client_secret, refresh_token)

if FileExist( "settings.ini" ) {
    
    IniRead, status_hotkey, settings.ini, General, status_hotkey
    IniRead, operations_hotkey, settings.ini, General, operations_hotkey
    
    IniRead, wing_enter_hotkey, settings.ini, General, wing_enter_hotkey
    IniRead, all_wings_hotkey, settings.ini, General, all_wings_hotkey
    
    IniRead, login_type, settings.ini, General, login_type
    
    IniRead, e_session_id, settings.ini, Session, session_id
    
    if (e_session_id) && (e_session_id <> "ERROR")
    {
        session_id := Decrypt(e_session_id)
    }
    
    if status_sheet_key <> ""
    {
        statusWS := getWSnames(status_sheet_key, myAccessToken)
    }
    
    if ops_sheet_key <> ""
    {
        opsWS := getWSnames(ops_sheet_key, myAccessToken)
    }
    
    if wings_sheet_key <> ""
    {
        wingsWS := getWSnames(wings_sheet_key, myAccessToken)
    }
    
    if % statusWS.length()
    {
        loop, % statusWS.length()
        {   
            IniRead, statusWS_bcg_%A_index%, settings.ini, Worksheets, statusWS_bcg_%A_index%
            IniRead, statusWS_txt_%A_index%, settings.ini, Worksheets, statusWS_txt_%A_index%
        }
    }
    if % opsWS.length()
    {
        loop, % opsWS.length()
        {
            IniRead, opsWS_bcg_%A_index%, settings.ini, Worksheets, opsWS_bcg_%A_index%
            IniRead, opsWS_txt_%A_index%, settings.ini, Worksheets, opsWS_txt_%A_index%
        }
    }
    if % wingsWS.length()
    {
        loop, % wingsWS.length()
        {
            IniRead, wingsWS_bcg_%A_index%, settings.ini, Worksheets, wingsWS_bcg_%A_index%
            IniRead, wingsWS_txt_%A_index%, settings.ini, Worksheets, wingsWS_txt_%A_index%
        }
    }
    IniRead, nowTimeStamp, settings.ini, Wings, last_check
}
else
{
    ErrorGui("setting.ini file not found. Exiting...")
    ExitApp
}

if !nowTimeStamp
{
    nowTimestamp := A_NowUTC
    nowTimestamp -= 19700101000000,seconds
}

StringUpper, uOrgName, org_name

if (login_type = "Enjin_API")
{
    if session_id
    {   
        is_session_valid := checkSession(org_website, session_id)
        
        is_identified := is_session_valid["identity"]
        
        if is_identified
        {
            isLogged := 1
            logusername := is_session_valid["username"]
            goto, StartApp
        }
        else
        {
            goto, LoginEnjin
        }
    }
    else
    {
        goto, LoginEnjin
    }
}

if (login_type = "Shivtr_API")
{
    goto, LoginShivtr
}

if (login_type = "WithoutLogin")
{
    isLogged := 1
    goto, StartApp
}

LoginEnjin:
    
    Gui, Destroy
    
    Suspend, On
    
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 20,40
    Gui, Font, s18 bold
    Gui, Add, Text, xm-2 cYellow section, Login with Enjin API
    Gui, Font, s10 cLime normal
    Gui, Add, Text, xp+2 yp+62, %uOrgName%
    Gui, Font, s10 cWhite
    Gui, Add, Text, xs cWhite, email:
    Gui, Font, s10 cBlack
    Gui, Add, Edit, xs yp+20 w220 vEmailAdr,
    Gui, Font, s10 cWhite
    Gui, Add, Text, xs yp+42 cWhite, password:
    Gui, Font, s10 cBlack
    Gui, Add, Edit, xs yp+20 w220 vUsrPass Password,
    Gui, Font, s14 cBlack
    Gui, Add, Button, xs yp+56 vButtonLogin gButtonLoginEnjin default, Login
    Gui, Font, s6
    Gui, Add, Text, xs+2,
    Gui, Font, s9 c555555 normal
    Gui, Add, Text, xp+2 yp, %appname% v.%appversion% by %appauthor%
    WinSet, TransColor, EE0000 175
    Gui, Show, Center autosize, Login
    GuiControl, focus, EmailAdr  
    return

ButtonLoginEnjin:
    Suspend, On
    Gui, Submit
    
    Gui, Destroy
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 40,20
    Gui, font, s18 cwhite
    Gui, Add, Text, center cWhite, .... trying to login ....
    WinSet, TransColor, EE0000 175
    Gui, show, autosize
    
    LoginResponse := userLoginEnjin(org_website, EmailAdr, UsrPass)
    
    if LoginResponse["session"]
    {
        logusername := LoginResponse["username"]
        session_id := LoginResponse["session"]
        sessionid := Encrypt(session_id)
        IniWrite, %sessionid%, settings.ini, Session, session_id
        isLogged := 1
        ErrorGui("Welcome " logusername "!")
        goto, StartApp
    }
    else
    {
        ErrorGui("Login failed. Try Again.")
        Gui, Destroy
        Suspend, Off
        goto, LoginEnjin
    }
    
    return
    
LoginShivtr:
    
    Gui, Destroy
    
    Suspend, On
    
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 20,40
    Gui, Font, s18 bold
    Gui, Add, Text, xm-2 cYellow section, Login with Shivtr API
    Gui, Font, s10 cLime normal
    Gui, Add, Text, xp+2 yp+62, %uOrgName%
    Gui, Font, s10 cWhite
    Gui, Add, Text, xs cWhite, email:
    Gui, Font, s10 cBlack
    Gui, Add, Edit, xs yp+20 w220 vEmailAdr,
    Gui, Font, s10 cWhite
    Gui, Add, Text, xs yp+42 cWhite, password:
    Gui, Font, s10 cBlack
    Gui, Add, Edit, xs yp+20 w220 vUsrPass Password,
    Gui, Font, s14 cBlack
    Gui, Add, Button, xs yp+56 vButtonLogin gButtonLoginShivtr default, Login
    Gui, Font, s6
    Gui, Add, Text, xs+2,
    Gui, Font, s9 c555555 normal
    Gui, Add, Text, xp+2 yp, %appname% v.%appversion% by %appauthor%
    WinSet, TransColor, EE0000 175
    Gui, Show, Center autosize, Login
    GuiControl, focus, EmailAdr  
    return

ButtonLoginShivtr:
    Suspend, On
    Gui, Submit
    
    Gui, Destroy
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 40,20
    Gui, font, s18 cwhite
    Gui, Add, Text, center cWhite, .... trying to login ....
    WinSet, TransColor, EE0000 175
    Gui, show, autosize
    
    LoginResponse := userLoginShivtr(org_website, EmailAdr, UsrPass)
    
    if LoginResponse["userid"]
    {
        logusername := LoginResponse["username"]
        isLogged := 1
        ErrorGui("Welcome " logusername "!")
        goto, StartApp
    }
    else
    {
        ErrorGui("Login failed. Try Again.")
        Gui, Destroy
        Suspend, Off
        goto, LoginShivtr
    }
    
    return

StartApp:

gosub, CheckWings
gosub, HelpScreen
Hotkey, %status_hotkey%, CmdrCheckPrompt
Hotkey, %operations_hotkey%, OperationsList
Hotkey, %wing_enter_hotkey%, CmdrWingsPrompt
Hotkey, %all_wings_hotkey%, CheckWingsList
Hotkey, F1, HelpScreen
SetTimer, CheckWings, 60000
return

HelpScreen:

    Gui, Destroy
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 20,20
    Gui, Font, s22 bold
    Gui, Add, Text, xm-2 cWhite section, CMDR-Check Help Screen
    Gui, Font, s10 cLime normal
    Gui, Add, Text, xp+2 yp+32 , list of key commands
    
    Gui, Font, s22 bold
    Gui, Add, Text, xm cYellow section, %status_hotkey%
    Gui, Font, s11 normal
    Gui, Add, Text, ys cWhite w300, Commander check `ncheck commander status from official status lists
    
    Gui, Font, s22 bold
    Gui, Add, Text, xm cYellow section, %operations_hotkey%
    Gui, Font, s11 normal
    Gui, Add, Text, ys cWhite w300, List of operations `nopens list of operations. By clicking on operation name you can see operation details
    
    Gui, Font, s22 bold
    Gui, Add, Text, xm cYellow section, %wing_enter_hotkey%
    Gui, Font, s11 normal
    Gui, Add, Text, ys cWhite w300, Wing Broadcast entry `nsend wing broadcast to other members who are currently using EIC CMDR-Check app
    
    Gui, Font, s22 bold
    Gui, Add, Text, xm cYellow section, %all_wings_hotkey%
    Gui, Font, s11 normal
    Gui, Add, Text, ys cWhite w300, List of all active Wing Broadcasts `nopens list of current wing broadcasts
    
    Gui, Font, s9 c555555 normal
    Gui, Add, Text, xp+2 yp+48, %appname% v.%appversion% by %appauthor%
    
    WinSet, TransColor, EE0000 175
    Gui, Show, Center autosize, CMDRCheck

    return  

OperationsList:
    Gui, Destroy
    
    Suspend, On

    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 40,20
    Gui, font, s18 cwhite
    Gui, Add, Text, center cWhite, .... Retrieving Active Operations ....
    WinSet, TransColor, EE0000 175
    Gui, show, autosize
    
    myAccessToken := GetAccessCode(client_id, client_secret, refresh_token)
    operations := getWorksheetsData(ops_sheet_key, myAccessToken)

    Gui, Destroy

    Gui, +e0x80 +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    
    if operations
    {  
        loop, % operations.length()
        {   
            indexcnt := A_index
            msgalertcolor := opsWS_bcg_%indexcnt%
            msxtxtcolor := opsWS_txt_%indexcnt%
            operationTitleNum%indexcnt% := getWorksheetsTitleNum(operations[A_index])
            Gui, font, s18 normal
            Gui, Add, Text, ym c%msgalertcolor% section, % operationTitleNum%A_index%["title"]
            Gui, font, s10 cEEEEEE normal
            Gui, Add, ListView, -LV0x8 LV0x100 IconSmall r25 w250 vOps_%A_index% gLv -WantF2 -Hdr -E0x200 -Multi -grid BackgroundTrans +ReadOnly AltSubmit, OperationsLists
            LV_ModifyCol(1, 220)
            ImageListID := IL_Create(2)  ; Create an ImageList to hold 10 small icons.
            LV_SetImageList(ImageListID)  ; Assign the above ImageList to the current ListView.
            IL_Add(ImageListID, "shell32.dll", 50)
            IL_Add(ImageListID, "shell32.dll", 14)
                           
            loop, % operationTitleNum%A_index%["entries"]
            {   
                fieldscnt := A_index
                allentries := getOpEntries(operations[indexcnt], fieldscnt)
                iconstr :=
                loop, % allentries[fieldscnt]["fields"]
                {
                    if A_index > 1
                    { 
                        if % allentries[fieldscnt]["content"][A_index]
                        {
                            iconstr := "Icon2"
                        }
                        else
                        {
                            iconstr := ""
                        }
                    }
                }
                
                curValue := allentries[fieldscnt]["content"][1]
                LV_Add(iconstr, curValue)
            }
			
        }
    }
    else
    {
        ErrorGui()
    }

      ; Auto-size each column to fit its contents.
    
    Gui, margin, 50,20 
    WinSet, TransColor, EE0000 175
    Gui, show, autosize, OperationsScreen
    Suspend, Off
    return
    
    Lv:
    if (a_guievent = "normal") 
    {   
        
        StringSplit, tempOpsIdx, A_GuiControl, _,
        sel_opsidx := tempOpsIdx2
        sel_op := A_EventInfo
        LV_Modify(sel_op, "-Select")
        sel_entries := getOpEntries(operations[sel_opsidx], sel_op)
        sel_op_entries := sel_entries[sel_op]["fields"]
        
        if sel_op_entries
        {
            loop, %sel_op_entries%
            {
                if A_index > 1
                {
                    curValue := sel_entries[sel_op]["content"][A_index]
                    if curValue
                    {
                       goto, ShowOp 
                    }
                    else 
                    {
                        stillEmpty := 1
                    }
                }
            }
            if stillEmpty
            {   
                
                #Persistent
                Tooltip, % "no additional data for " sel_entries[sel_op]["content"][1]
                SetTimer, RemoveToolTip, 1500
                return

                RemoveToolTip:
                SetTimer, RemoveToolTip, Off
                ToolTip
                return 
            }
        }
    }
    return
    
    ShowOp:
        Suspend, On

        idxcnt := sel_opsidx
        allentriesidx := sel_op
        allentries := getOpEntries(operations[idxcnt], allentriesidx)
        msgalertcolor := opsWS_bcg_%idxcnt%
        msxtxtcolor := opsWS_txt_%idxcnt%
        Gui, Destroy
        Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
        Gui, Color, 000000
        Gui, margin, 16,30
        
        StringUpper, OpCat, % operationTitleNum%idxcnt%["title"]
        
        Gui, font, s10 c%msgalertcolor% bold
        Gui, Add, text, xm+16, % OpCat
        Gui, font, s15 c%msxtxtcolor% bold
        Gui, Add, text, xp yp+20, % allentries[allentriesidx]["content"][1]
        
        Gui, margin, 30,10
        Gui, font, s4 normal
        Gui, Add, text, ,
        Gui, font, s10 c%msxtxtcolor% normal
        loop, % allentries[allentriesidx]["fields"]
        {
            if A_index > 1
            {
                if % allentries[allentriesidx]["content"][A_index]
                {
                    if SubStr(allentries[allentriesidx]["content"][A_index], 1 , 4) = "http"
                    {
                        oper_url := % allentries[allentriesidx]["content"][A_index]
                        Gui, font, bold
                        Gui, Add, text, xm w100 section, % allentries[allentriesidx]["label"][A_index]
                        Gui, font, normal underline
                        Gui, Add, text, ys w550 gGoToForum, % oper_url
                    }
                    else
                    {   
                        Gui, font, bold
                        Gui, Add, text, xm cWhite w100 section, % allentries[allentriesidx]["label"][A_index]
                        Gui, font, normal
                        Gui, Add, text, ys cWhite w550, % allentries[allentriesidx]["content"][A_index]
                    }
                }
            }
        }
        
        Gui, margin, 30,10
        Gui, Add, Text, x40, 
        WinSet, TransColor, FFFFFF 175
        Gui, show, autosize, OperationScreen
        Suspend, Off
        return
    
    GoToForum:
    this_link := A_GuiControl
    Run %this_link%
    showAndBreak(status_hotkey, operations_hotkey, wing_enter_hotkey, all_wings_hotkey) 
    Gui, Destroy
    Suspend, Off
    Exit

return

CmdrCheckPrompt:
    
    Gui, Destroy
    
    Suspend, On
    
    myAccessToken := GetAccessCode(client_id, client_secret, refresh_token)
    
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 20,40
    Gui, Font, s25 bold
    Gui, Add, Text, xm-2 cYellow section, CMDR CHECK
    Gui, Font, s10 cLime normal
    Gui, Add, Text, xp+2 yp+38, %uOrgName%
    Gui, Font, s14 cWhite
    Gui, Add, Text, xm cWhite, Enter CMDR name:
    Gui, Font, s16 cBlack
    Gui, Add, Edit, xm yp+26 vTargetName  
    Gui, Font, s9 c555555 normal
    Gui, Add, Text, xp+2 yp+48, %appname% v.%appversion% by %appauthor%
    Gui, Add, Button, vButtonHidden gButtonCheckCommander default, Check Commander
    guicontrol, hide, ButtonHidden
    WinSet, TransColor, EE0000 175
    Gui, Show, Center autosize, CMDRCheck
    GuiControl, focus, TargetName
    Suspend, Off
    return  

    ButtonCheckCommander:
        Suspend, On
        Gui, Submit  
        
        if (TargetName = "")
            {
                Gui, Destroy
                CStatus := 
                Suspend, Off
                return
            }
        else          
        
            Gui, Destroy

            Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
            Gui, Color, 000000
            Gui, margin, 40,30
            Gui, font, s18 cwhite
            Gui, Add, Text, center cWhite, .... Searching for CMDR %TargetName% ....
            WinSet, TransColor, EE0000 175
            Gui, show, autosize

            if status_sheet_key <> ""
            { 
                CStatus := SearchCMDR(status_sheet_key, myAccessToken, TargetName)
                index := CStatus["index"]
                msgtxt := CStatus["status"]
                msgalertcolor = % statusWS_bcg_%index%
                msxtxtcolor = % statusWS_txt_%index%
            }
            else
            {
                ErrorGui()
                ExitApp
            }
            
            if (!CStatus["status"])
            {
                msgalertcolor = 000000
                msxtxtcolor = ffffff
                msgtxt = Unknown
            }
            
            Gui, Destroy
            Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
            Gui, Color, %msgalertcolor%
            Gui, margin, 30,30
            Gui, font, s18 normal
            Gui, Add, Text, c%msxtxtcolor%, CMDR %TargetName%
            Gui, font, s10 bold
            Gui, Add, Text, c%msxtxtcolor%, STATUS:
            Gui, font, s28 bold
            Gui, Add, Text, xm yp+18 c%msxtxtcolor%, %msgtxt%
            Gui, margin, 30,0
            Gui, font, s8 bold
            Gui, Add, Text, xm,

            Gui, font, s11
            loop, % CStatus["colcnt"]
            {
                if A_index > 1
                {
                    if CStatus["content"A_Index]
                    {
                        Gui, font, bold
                        Gui, Add, text, xm c%msxtxtcolor% w90 section, % CStatus["column"A_Index]
                        Gui, font, normal
                        Gui, Add, text, ys c%msxtxtcolor%, % CStatus["content"A_Index]
                    }
                }
            }
            
            WinSet, TransColor, 000000 175
            Gui, margin, 30,30
            Gui, show, autosize
            showAndBreak(status_hotkey, operations_hotkey, wing_enter_hotkey, all_wings_hotkey, 8)
            Gui, Destroy
            Suspend, Off 
            return

CmdrWingsPrompt:
    
    Gui, Destroy
    
    Suspend, On
    
    myAccessToken := GetAccessCode(client_id, client_secret, refresh_token)
    
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 20,40
    Gui, Font, s25 bold
    Gui, Add, Text, xm-2 cYellow section, Wing broadcast
    Gui, Font, s10 cLime normal
    Gui, Add, Text, xp+2 yp+42, %uOrgName%
    Gui, Font, s10 cWhite
    Gui, Add, Text, xs, Your name:
    if logusername 
    {
        Gui, Font, s12 cWhite
        Gui, Add, Edit, xs yp+20 vLeaderName +readonly -E0x200, % logusername
    }
    else
    {
        Gui, Font, s12 cblack
        Gui, Add, Edit, xs yp+20 vLeaderName,
    }
    Gui, Font, s10 cWhite
    Gui, Add, Text, xs yp+42 cWhite, System:
    Gui, Font, s12 cBlack
    Gui, Add, Edit, xs yp+20 w150 vSystemName,
    Gui, Font, s10 cWhite
    Gui, Add, Text, xs yp+42 cWhite, Mission type:
    Gui, Font, s12 cBlack
    Gui, Add, DropDownList, xm yp+20 w120 h20 vWingType r5, PvE|PvP|PvE/PvP||Powerplay|Operation
    Gui, Font, s10 cWhite
    Gui, Add, Text, xs yp+42 cWhite, Duration (in hours):
    Gui, Font, s12 cBlack
    Gui, Add, DropDownList, xm yp+20 w80 h20 vWingDuration r6, 1|2||3|4|5|6
    Gui, Font, s14 cBlack
    Gui, Add, Button, xs yp+56 vButtonHidden gButtonWingBroadcast default, Broadcast Wing
    Gui, Font, s6
    Gui, Add, Text, xs+2,
    Gui, Font, s9 c555555 normal
    Gui, Add, Text, xp+2 yp, %appname% v.%appversion% by %appauthor%
    WinSet, TransColor, EE0000 175
    Gui, Show, Center autosize, WingBroadcast
    GuiControl, focus, SystemName
    Suspend, Off
    return

ButtonWingBroadcast:
    Suspend, On
    Gui, Submit
    
    Gui, Destroy
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 40,20
    Gui, font, s18 cwhite
    Gui, Add, Text, center cWhite, .... Sending wings broadcast ....
    WinSet, TransColor, EE0000 175
    Gui, show, autosize
    
    broadcastTimestamp := A_NowUTC
    broadcastTimestamp -= 19700101000000,seconds
    
    WingsWSLinks := getWorksheets(wings_sheet_key, myAccessToken)
    xmlentry := "<entry xmlns='http://www.w3.org/2005/Atom' xmlns:gsx='http://schemas.google.com/spreadsheets/2006/extended'><gsx:timestamp type='text'>" broadcastTimestamp "</gsx:timestamp><gsx:name>" LeaderName "</gsx:name><gsx:system>" SystemName "</gsx:system><gsx:type>" WingType "</gsx:type><gsx:duration>" WingDuration "</gsx:duration></entry>"
    
    posturl := WingsWSLinks[1] "?access_token=" myAccessToken  
    o2HTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    o2HTTP.Open("POST", posturl , False)
    o2HTTP.SetRequestHeader("Content-Type", "application/atom+xml")
    o2HTTP.Send(xmlentry)
    o2HTTP :=
    Gui, Destroy
    Suspend, Off
    return

CheckWings:
    
    myAccessToken := GetAccessCode(client_id, client_secret, refresh_token)
    WingsList := getWorksheetsData(wings_sheet_key, myAccessToken)
    
    if WingsList
    {   
        newentrycnt := 0
        
        loop, % WingsList.length()
        {   
            indexcnt := A_index
            operationTitleNum%indexcnt% := getWorksheetsTitleNum(WingsList[A_index])
            broadcastnum := operationTitleNum%A_index%["entries"]
            loop, % broadcastnum
            {
                
                fieldscnt := A_index
                allentries := getOpEntries(WingsList[indexcnt], fieldscnt)
                curtimestamp := allentries[fieldscnt]["content"][1]
                if (curtimestamp > nowtimestamp)
                {   
                    newentrycnt++
                    if newentrycnt = 1
                    {
                        Gui, 7: +Owner +AlwaysOnTop -border -caption -ToolWindow
                        Gui, 7: Color, 000000
                        Gui, 7: margin, 20,20
                        Gui, 7: Font, s18 bold
                        Gui, 7:Add, Text, xm-2 cYellow section, New wing broadcast:
                    }
                    if newentrycnt < 6
                    {
                        loop, % allentries[fieldscnt]["fields"]
                        {
                            
                                if A_index = 1 
                                {
                                    Gui, 7: font, s2 normal
                                    Gui, 7:Add, Text, xs cWhite h12,
                                    Gui, 7: font, s10 bold
                                    Gui, 7:Add, text, xs yp+24 cWhite h24 section, Broadcast time:
                                    Gui, 7: font, s10 normal
                                    bcsttime := formatstamp(allentries[fieldscnt]["content"][A_index])
                                    Gui, 7:Add, text, ys cWhite h24, % bcsttime
                                }
                                else
                                {
                                    sufix := (A_index = 5) ? "h" : ""
                                    Gui, 7: font, s10 bold
                                    Gui, 7:Add, text, xs yp+24 cWhite h24 section, % allentries[fieldscnt]["label"][A_index]
                                    Gui, 7: font, s10 normal
                                    Gui, 7:Add, text, ys cWhite h24, % allentries[fieldscnt]["content"][A_index] " " sufix
                                }
                            foundnewwing := 1
                        }
                    }
                    else
                    {   
                        morenew := A_index
                    }
                }
            }
            morenum := morenew - 5
            if morenum
            {   
                Gui, 7: font, s12 normal
                Gui, 7:Add, Text, xs cWhite h24,
                Gui, 7:Add, text, xs yp+24 cWhite h24 section, % "press F7 to see " morenum " more"
            }
        }
        nowTimestamp := A_NowUTC
        nowTimestamp -= 19700101000000,seconds
        if foundnewwing
        {
            Gui, 7:Show, x30 y30 autosize, BroadcastAlert
            WinSet, TransColor, EE0000 175, BroadcastAlert
            SetTimer, CheckWingsGUIClose, 10000
        }
        foundnewwing := 0
    }
    IniWrite, %nowTimestamp%, settings.ini, Wings, last_check
    return

CheckWingsGUIClose:
    gui, 7:Destroy
    SetTimer, CheckWingsGUIClose, Off
    return

CheckWingsList:

    Gui, Destroy
    
    Suspend, On
    
    if wings_sheet_key <> ""
    {
        msgalertcolor = % wingsWS_bcg_1
        msxtxtcolor = % wingsWS_txt_1
    }
    else
    {
        ErrorGui()
        ExitApp
    }
    
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 40,20
    Gui, font, s18 cwhite
    Gui, Add, Text, center cWhite, .... Scanning broadcasts ....
    WinSet, TransColor, EE0000 175
    Gui, Show, Center autosize, ScanWingBroadcast
    
    myAccessToken := GetAccessCode(client_id, client_secret, refresh_token)
    WingsList := getWorksheetsData(wings_sheet_key, myAccessToken)
    
    if WingsList
    {  
        Gui, Destroy
        Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
        Gui, Color, %msgalertcolor%
        Gui, margin, 20,20
        Gui, Font, s18 cYellow bold
        Gui, Add, Text, xm-2 section, All wings broadcasts
        DeleteWing := []
        loop, % WingsList.length()
        {   
            indexcnt := A_index
            operationTitleNum%indexcnt% := getWorksheetsTitleNum(WingsList[A_index])
            foundanywing := 0
            loop, % operationTitleNum%A_index%["entries"]
            {   
                fieldscnt := A_index
                allentries := getWingsEntries(WingsList[indexcnt], fieldscnt)
                curtimestamp := allentries[fieldscnt]["content"][1]
                delusername := allentries[fieldscnt]["content"][2]
                
                loop, % allentries[fieldscnt]["fields"]
                {
                    foundanywing := 1
                    if A_index = 1 
                    {
                        Gui, font, s10 c%msxtxtcolor% bold
                        Gui, Add, text, xm h24 section, Broadcast time:
                        Gui, font, s10 c%msxtxtcolor% normal
                        bcsttime := formatstamp(allentries[fieldscnt]["content"][A_index])
                        curentrytime := allentries[fieldscnt]["content"][A_index]
                        Gui, Add, text, xp yp+18 h24, % bcsttime
                    }
                    else
                    {
                        sufix := (A_index = 5) ? "h" : ""
                        Gui, font, s10 c%msxtxtcolor% bold
                        Gui, Add, text, ys h24, % allentries[fieldscnt]["label"][A_index]
                        Gui, font, s10 c%msxtxtcolor% normal
                        Gui, Add, text, yp+18 h24, % allentries[fieldscnt]["content"][A_index] " " sufix
                    }
                }
                currnowtime := A_NowUTC
                currnowtime -= 19700101000000, seconds
                curduration := (allentries[fieldscnt]["content"][5] * 3600) + curentrytime
                curremainingtime := FormatSeconds((curduration - currnowtime))
                Gui, font, s10 c%msxtxtcolor% bold
                Gui, Add, text, ys h24, remaining time
                Gui, font, s10 c%msxtxtcolor% normal
                Gui, Add, text, yp+18 h24, % curremainingtime " h"
                if (delusername = logusername)
                {
                    DeleteWing[fieldscnt] := allentries[fieldscnt]["edit"]
                    Gui, font, s10 c%msxtxtcolor% bold
                    Gui, Add, text, ys h24, delete
                    Gui, font, s10 cRed bold
                    Gui, Add, text, yp+18 h24 vDel_%fieldscnt% gButtonWingDelete, x
                }
                
            }
        }
        if foundanywing
        {
            WinSet, TransColor, 000000 175
            Gui, Show, Center autosize, AllWings
        }
        else
        {
            Gui, Destroy
            Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
            Gui, Color, 000000
            Gui, margin, 40,20
            Gui, font, s18 cwhite
            Gui, Add, Text, center cWhite, .... no active broadcasts at the moment ....
            WinSet, TransColor, EE0000 175
            Gui, Show, Center autosize, NotWingBroadcast
            showAndBreak(status_hotkey, operations_hotkey, wing_enter_hotkey, all_wings_hotkey)
            Suspend, Off
            Gui, Destroy
        }
        foundanywing := 0
    }
    else
    {
        ErrorGui()
    }
    Suspend, Off
    return
    
ButtonWingDelete:
    Suspend, On
    Gui, Submit
    
    StringSplit, tempdel, A_GuiControl, _,
    tmpidx = % tempdel2
    delurl := DeleteWing[tmpidx]
    
    Gui, Destroy
    Gui, +lastFound +Owner +AlwaysOnTop -border -caption -ToolWindow
    Gui, Color, 000000
    Gui, margin, 40,20
    Gui, font, s18 cwhite
    Gui, Add, Text, center cWhite, .... Deleting wings broadcast ....
    WinSet, TransColor, EE0000 175
    Gui, show, autosize
    
    posturl := delurl "?access_token=" myAccessToken  
    o2HTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    o2HTTP.Open("DELETE", posturl)
    o2HTTP.Send()
    o2HTTP :=
    Gui, Destroy
    Suspend, Off
    return

ButtonCancel:
    Gui, Destroy
    CStatus := ""
    Operations := ""
    Suspend, Off
    Return

GuiClose:
GuiCancel:
GuiEscape:

    Gui, Destroy
    CStatus := ""
    Operations := ""
    foundanywing := 0
    foundnewwing := 0
    Suspend, Off 

    if !isLogged
    {
        Gui, Destroy
        CStatus := ""
        Operations := ""
        foundanywing := 0
        foundnewwing := 0
        Suspend, Off
        
        if (login_type = "Enjin_API")
        {
            goto, LoginEnjin
        }

        if (login_type = "Shivtr_API")
        {
            goto, LoginShivtr
        }
    }
    
    Return
    