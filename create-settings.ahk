#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance,Force
#include JSON.ahk
#include functions.ahk
#include CLR.ahk

if FileExist( "settings.ini" ) 
{
    goto, WorksheetSettings
} 

Gui, 1:Add, Picture, x+10 ym+20 w480 h50 , eis-head.png

Gui, 1:Add, GroupBox, xs y+20 w1005 h103 Section, GENERAL INFO
Gui, 1:Add, Text, xp+20 yp+30 w280 h20, Group/Organisation name: 
Gui, 1:Add, Text, xp+490 yp w100 h20 , Status Hotkey:
Gui, 1:Add, Text, xp+110 yp w100 h20 , Operations Hotkey:
Gui, 1:Add, Text, xp+110 yp w100 h20 , Wing Enter Hotkey:
Gui, 1:Add, Text, xp+110 yp w100 h20 , All Wings Hotkey:
Gui, 1:Add, Edit, xs+20 yp+15 w480 h20 vOrgName,
Gui, 1:Add, DropDownList, xp+490 yp w100 h20 vStatusHotkey r11, F2|F3|F4|F5|F6|F7|F8||F9|F10|F11|F12
Gui, 1:Add, DropDownList, xp+110 yp w100 h20 vOpsHotkey r11, F2|F3|F4|F5|F6|F7|F8|F9||F10|F11|F12
Gui, 1:Add, DropDownList, xp+110 yp w100 h20 vWingHotkey r11, F2|F3|F4|F5|F6||F7|F8|F9|F10|F11|F12
Gui, 1:Add, DropDownList, xp+110 yp w100 h20 vAllWingsHotkey r11, F2|F3|F4|F5|F6|F7||F8|F9|F10|F11|F12

Gui, 1:Add, GroupBox, xs y+40 w1005 h103 Section, LOGIN INFO
Gui, 1:Add, Text, xp+20 yp+30 w280 h20, Organisation Website (with http://): 
Gui, 1:Add, Edit, xp yp+15 w480 h20 vOrgWebsite disabled,
Gui, 1:Add, radio, xp+520 yp w120 h20 vLoginEnjin, Enjin API login 
Gui, 1:Add, radio, xp+120 yp w120 h20 vLoginShivtr, Shivtr API login
Gui, 1:Add, radio, xp+120 yp w120 h20 vWithoutLogin Checked, without login


Gui, 1:Add, GroupBox, xs w498 h105 Section, CMDR STATUS SPREADSHEET
Gui, 1:Add, Text, xp+15 yp+30 w220 h20 , Spreadsheet key:
Gui, 1:Add, Edit, xp yp+20 w465 h20 vStatusSheetKey, 

Gui, 1:Add, GroupBox, ys w498 h325, GOOGLE OAUTH 2.0 ACCESS SETTINGS
Gui, 1:Add, Text, xp+20 yp+30 w210 h20 , Client ID:
Gui, 1:Add, Edit, yp+20 w460 h20 vClientId, 
Gui, 1:Add, Text, yp+30 w230 h20 , Client Secret:
Gui, 1:Add, Edit, yp+20 w230 h20 vClientSecret, 
Gui, 1:Add, Button, yp+40 w460 h30 vGetAuthCodeBtn gGetAuthCode, Get Authorization Code
Gui, 1:Add, Text, yp+40 w400 h20 , (after copying Authorization code from popup window it should appear in this box)
Gui, 1:Add, Edit, yp+15 w460 h20 vAuthCode, 
Gui, 1:Add, Button, yp+40 w460 h30 vGetRefreshTokenBtn gGetRefreshToken, Request Refresh Token
Gui, 1:Add, Text, yp+40 w420 h20 , (after requesting Refresh token it should appear in this box)
Gui, 1:Add, Edit, yp+15 w460 h20 vRefreshToken +ReadOnly, 

Gui, 1:Add, GroupBox, xs yp-180 w500 h105 Section, OPERATIONS SPREADSHEET
Gui, 1:Add, Text, xp+15 yp+30 w220 h20 , Spreadsheet key:
Gui, 1:Add, Edit, xp yp+20 w465 h20 vOpsSheetKey, 

Gui, 1:Add, GroupBox, xs w498 h105 Section, WINGS BROADCAST SPREADSHEET
Gui, 1:Add, Text, xp+15 yp+30 w220 h20 , Spreadsheet key:
Gui, 1:Add, Edit, xp yp+20 w465 h20 vWingsSheetKey, 

Gui, 1:Add, text, xm w10 h10,
Gui, 1:Add, Button, xm w1030 h50 Center vSubmitBtn gCreateSettings, CREATE SETTINGS FILE AND GO TO WORKSHEETS SETTINGS
Gui, 1:Add, text, xm w10 h10,
Gui, 1:Add, Text, xm w690 h20 , This program does not interact with Elite Dangerous application in any level. It serve as custom GUI for interacting with Google Spreadsheets.
Gui, 1:Add, Text, xm w450 h20 , developed by Corwin Arzak in AutoHotKey scripting language.

SetTimer, EnableButtonsSetup, 1
Gui, 1:Show, w1050 h780, Settings
GuiControl, 1:Disable, GetAuthCodeBtn
GuiControl, 1:Disable, GetRefreshTokenBtn
GuiControl, 1:Disable, SubmitBtn
return

GetAuthCode:
    gui_title := "Get Authorization Code"
    GuiControlGet, urlClientId,, ClientId
    authURL := "https://accounts.google.com/o/oauth2/auth?scope=https://spreadsheets.google.com/feeds&redirect_uri=urn:ietf:wg:oauth:2.0:oob&response_type=code&client_id=" urlClientId
    WEB := ComObjCreate("Shell.Explorer")
    WEB.Visible := true
    Gui +LastFound
    Gui, 2:Color, FFFFFF
    Gui, 2:Add, ActiveX, w1000 h700 x5 y55 vWEB, Shell.Explorer
    WEB.Navigate(authURL)
    Gui,2:show, w1010 h760 x5 y5, %gui_title%
    Gui2_ID := WinExist()
    GroupAdd, Auth_Gui, ahk_id %Gui2_ID%
    clipboard =  
    ClipWait
    GuiControl, 1:, AuthCode, %clipboard%
    If pwb !=
      ObjRelease(WEB)
    Gui, 2:destroy
    GuiControl, 1:+ReadOnly, ClientId
    GuiControl, 1:+ReadOnly, ClientSecret
    return
    #if WinActive(gui_title)
        ^c::
        send {AppsKey}c
        GuiControl, 1:, AuthCode, %clpbrd%
        If pwb !=
          ObjRelease(WEB)
        Gui, 2:destroy
        GuiControl, 1:+ReadOnly, ClientId
        GuiControl, 1:+ReadOnly, ClientSecret
    return

SwitchStatus:
    StringSplit, enabledstatus, A_GuiControl, _,
    wsflagops := "WSflagops"
    wsflagstatus := "WSflagstatus"
    if (%enabledstatus1% == wsflagstatus)
    {   
        GuiControl, 1:Enable, WS_color%enabledstatus2%
        GuiControl, 1:Enable, WS_color%enabledstatus2%_value
        GuiControl, 1:Enable, WS_text%enabledstatus2%
        GuiControl, 1:Enable, WS_text%enabledstatus2%_value
    }
    if (%enabledstatus1% == wsflagops) 
    {   
        GuiControl, 1:Disable, WS_color%enabledstatus2%
        GuiControl, 1:Disable, WS_color%enabledstatus2%_value
        GuiControl, 1:Disable, WS_text%enabledstatus2%
        GuiControl, 1:Disable, WS_text%enabledstatus2%_value
    }
    return

GetRefreshToken:
    GuiControlGet, urlClientId,, ClientId
    GuiControlGet, urlClientSecret,, ClientSecret
    GuiControlGet, urlAuthCode,, AuthCode
    rURL := "https://www.googleapis.com/oauth2/v3/token"
    rPostData := "code=" urlAuthCode "&client_id=" urlClientId "&client_secret=" urlClientSecret "&redirect_uri=urn:ietf:wg:oauth:2.0:oob&access_type=offline&grant_type=authorization_code"
    o2HTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    o2HTTP.Open("POST", rURL , False)
    o2HTTP.SetRequestHeader("User-Agent", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)")
    o2HTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    o2HTTP.Send(rPostData)
    data2 := o2HTTP.ResponseText
    parsedRToken := JSON.Load(data2, true)
    parsedRefreshToken := parsedRToken.refresh_token
    GuiControl, 1:, RefreshToken, %parsedRefreshToken%
    return
 
CreateSettings: 
    
    Gui, Submit

    if FileExist( "settings.ini" ) 
    {
        FileDelete, settings.ini
    }
    
    if (WithoutLogin = 1)
    {
        LoginType := "WithoutLogin"
    }
    
    if (LoginEnjin = 1)
    {
        LoginType := "Enjin_API"
    }
    
    if (LoginShivtr = 1)
    {
        LoginType := "Shivtr_API"
    }
    
    IniWrite, %StatusHotkey%, settings.ini, General, status_hotkey
    IniWrite, %OpsHotkey%, settings.ini, General, operations_hotkey
    IniWrite, %WingHotkey%, settings.ini, General, wing_enter_hotkey
    IniWrite, %AllWingsHotkey%, settings.ini, General, all_wings_hotkey
    IniWrite, %LoginType%, settings.ini, General, login_type
    
    e_ClientId := Encrypt(ClientId)
    e_ClientSecret := Encrypt(ClientSecret)
    e_RefreshToken := Encrypt(RefreshToken)
    e_OrgName := Encrypt(OrgName)
    e_OrgWebsite := Encrypt(OrgWebsite)
    e_StatusSheetKey := Encrypt(StatusSheetKey)
    e_OpsSheetKey := Encrypt(OpsSheetKey)
    e_WingsSheetKey := Encrypt(WingsSheetKey)
    
    c# =
    (
        using System;
        class Settings {
            public string gcc(string argum="%e_ClientId%") {
                return argum;
            }
            public string gcs(string argum="%e_ClientSecret%") {
                return argum;
            }
            public string grt(string argum="%e_RefreshToken%") {
                return argum;
            }
            public string gon(string argum="%e_OrgName%") {
                return argum;
            }
            public string gow(string argum="%e_OrgWebsite%") {
                return argum;
            }
            public string gss(string argum="%e_StatusSheetKey%") {
                return argum;
            }
            public string gos(string argum="%e_OpsSheetKey%") {
                return argum;
            }
            public string gws(string argum="%e_WingsSheetKey%") {
                return argum;
            }
        }
    )
    CLR_CompileC#(c#, "System.dll", 0, "settings.dll")
    
    if FileExist("settings.ini") && FileExist("settings.dll")
    {
        goto, WorksheetSettings
    }
    return

EnableButtonsSetup:
    
    GuiControlGet, wlogin,, WithoutLogin
    GuiControlGet, cid,, ClientId
    GuiControlGet, csec,, ClientSecret
    GuiControlGet, autc,, AuthCode
    GuiControlGet, reft,, RefreshToken
    GuiControlGet, oname,, OrgName
    GuiControlGet, skey,, StatusSheetKey
    
    if (!wlogin)
    {
        GuiControl, 1:Enable, OrgWebsite
    }
    else
    {
        GuiControl, 1:Disable, OrgWebsite
    }
    
    If (cid && csec)
        GuiControl, 1:Enable, GetAuthCodeBtn
    else
        GuiControl, 1:Disable, GetAuthCodeBtn
    
    If (cid && csec && autc)
    {
        GuiControl, 1:Enable, GetRefreshTokenBtn
        GuiControl, 1:Disable, GetAuthCodeBtn
    }
    else
    {
        GuiControl, 1:Disable, GetRefreshTokenBtn
    }
    
    If (cid && csec && autc && reft)
    {
        GuiControl, 1:Disable, GetRefreshTokenBtn
        GuiControl, 1:Disable, GetAuthCodeBtn
        GuiControl, 1:+ReadOnly, AuthCode
    }
    
    If cid && csec && reft && oname && skey 
    {   
        GuiControl, 1:Enable, SubmitBtn
        SetTimer, EnableButtonsSetup, Off
    }
    else 
    {
        GuiControl, 1:Disable, SubmitBtn
    }
    
    return
    

ButtonColorPicker:
    color := ChooseColor(%A_GuiControl%)
	Return


WorksheetSettings:

    Gui, destroy
    SetTimer, EnableButtonsSetup, Off
    
    if FileExist( "settings.dll" ) {
        asm := CLR_LoadLibrary("settings.dll")
        obj := CLR_CreateObject(asm, "Settings")
        e_client_id := obj.gcc()
        e_client_secret := obj.gcs()
        e_refresh_token := obj.grt()
        e_org_name := obj.gon()
        e_status_sheet_key := obj.gss()
        e_ops_sheet_key := obj.gos()
        e_wings_sheet_key := obj.gws()
        
        client_id := Decrypt(e_client_id)
        client_secret := Decrypt(e_client_secret)
        refresh_token := Decrypt(e_refresh_token)
        org_name := Decrypt(e_org_name)
        status_sheet_key := Decrypt(e_status_sheet_key)
        ops_sheet_key := Decrypt(e_ops_sheet_key)
        wings_sheet_key := Decrypt(e_wings_sheet_key)
    }
    else
    {
        ErrorGui("settings.dll file not found. Exiting...")
        ExitApp
    }

    myAccessToken := GetAccessCode(client_id, client_secret, refresh_token)
    
    if FileExist( "settings.ini" ) {

        IniRead, status_hotkey, settings.ini, General, status_hotkey
        IniRead, operations_hotkey, settings.ini, General, operations_hotkey
        IniRead, wing_enter_hotkey, settings.ini, General, wing_enter_hotkey
        IniRead, all_wings_hotkey, settings.ini, General, all_wings_hotkey
        IniRead, login_type, settings.ini, General, login_type
        
        if status_sheet_key <> ""
        {
            statusWS := getWSnames(status_sheet_key, myAccessToken)
        }
        
        if ops_sheet_key <> ""
        {
            opsWS := getWSnames(ops_sheet_key, myAccessToken)
        }
        ;msgbox, % wings_sheet_key
        if wings_sheet_key <> ""
        {
            wingsWS := getWSnames(wings_sheet_key, myAccessToken)
        }
        
        IniRead, testWS, settings.ini, Worksheets,
        
        if testWS
        {
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
        }
    }
    else
    {
        ErrorGui("settings.ini file not found. Exiting...")
        ExitApp
    }
    
    StringUpper, uOrgName, org_name

    Gui, destroy
    Gui, 1:Add, Picture, x+10 ym+20 w480 h50 , eis-head.png
    
    Gui, 1:Add, GroupBox, xm y+20 w510 h440 , STATUS WORKSHEETS
    Gui, 1:Add, Text, xp+20 yp+35 w220 h20 section, status message for user:
    Gui, 1:Add, text, ys w82 h20, back color:
    Gui, 1:Add, text, ys w85 h20, text color:

    if % statusWS.length()
    {
        loop, % statusWS.length()
        {
            Gui, 1:Add, edit, xs w220 h20 readonly section, % statusWS[A_index]
            Gui, 1:Add, edit, ys w60 h20 vVal_statusWS_bcg_%A_index%, % statusWS_bcg_%A_index%
            Gui, 1:Add, button, x+0 yp w20 h20 vstatusWS_bcg_%A_index% gButtonColorPicker, ...
            Gui, 1:Add, edit, ys w60 h20 vVal_statusWS_txt_%A_index%, % statusWS_txt_%A_index%
            Gui, 1:Add, button, x+0 yp w20 h20 vstatusWS_txt_%A_index% gButtonColorPicker, ...
        }
    }
    
    Gui, 1:Add, GroupBox, ym+90 w510 h328 section, OPERATION WORKSHEETS
    Gui, 1:Add, Text, xp+20 yp+35 w220 h20 section, operations list name:
    Gui, 1:Add, text, ys w82 h20, back color:
    Gui, 1:Add, text, ys w85 h20, text color:

    if % opsWS.length()
    {
        loop, % opsWS.length()
        {
            Gui, 1:Add, edit, xs w220 h20 readonly section, % opsWS[A_index]
            Gui, 1:Add, edit, ys w60 h20 vVal_opsWS_bcg_%A_index%, % opsWS_bcg_%A_index%
            Gui, 1:Add, button, x+0 yp w20 h20 vopsWS_bcg_%A_index% gButtonColorPicker, ...
            Gui, 1:Add, edit, ys w60 h20 vVal_opsWS_txt_%A_index%, % opsWS_txt_%A_index%
            Gui, 1:Add, button, x+0 yp w20 h20 vopsWS_txt_%A_index% gButtonColorPicker, ...
        }
    }

    Gui, 1:Add, GroupBox, xs-20 y+228 w510 h105 section, WINGS BROADCAST WORKSHEETS
    Gui, 1:Add, Text, xp+20 yp+35 w220 h20 section, wings list name:
    Gui, 1:Add, text, ys w82 h20, back color:
    Gui, 1:Add, text, ys w85 h20, text color:

    if % wingsWS.length()
    {
        loop, % wingsWS.length()
        {
            Gui, 1:Add, edit, xs w220 h20 readonly section, % wingsWS[A_index]
            Gui, 1:Add, edit, ys w60 h20 vVal_wingsWS_bcg_%A_index%, % wingsWS_bcg_%A_index%
            Gui, 1:Add, button, x+0 yp w20 h20 vwingsWS_bcg_%A_index% gButtonColorPicker, ...
            Gui, 1:Add, edit, ys w60 h20 vVal_wingsWS_txt_%A_index%, % wingsWS_txt_%A_index%
            Gui, 1:Add, button, x+0 yp w20 h20 vwingsWS_txt_%A_index% gButtonColorPicker, ...
        }
    }

    Gui, 1:Add, text, xm w10 h10,
    Gui, 1:Add, Button, xm w1030 h50 Center vSubmitBtnUpdate gUpdateWorksheetSettings, UPDATE SETTINGS FILE
    Gui, 1:Add, text, xm w10 h10,
    Gui, 1:Add, Text, xm w690 h20 , This program does not interact with Elite Dangerous application in any level. It serve as custom GUI for interacting with Google Spreadsheets.
    Gui, 1:Add, Text, xm w450 h20 , developed by Corwin Arzak in AutoHotKey scripting language.
    
    Gui, 1:Show, w1050 h680, Settings
    GuiControl, 1:Disable, GetAuthCodeBtn
    GuiControl, 1:Disable, GetRefreshTokenBtn
    GuiControl, 1:Enable, SubmitBtnUpdate
    return
 
UpdateWorksheetSettings: 
    
    Gui, Submit
    
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
    
    if FileExist( "settings.ini" ) {
        if % statusWS.length()
        {
            loop, % statusWS.length()
            {   
                if % Val_statusWS_bcg_%A_index%
                {
                    IniWrite, % Val_statusWS_bcg_%A_index%, settings.ini, Worksheets, statusWS_bcg_%A_index%
                }
                else
                {
                    Val_statusWS_bcg_%A_index% := "555555"
                    IniWrite, % Val_statusWS_bcg_%A_index%, settings.ini, Worksheets, statusWS_bcg_%A_index%
                }
                if % Val_statusWS_txt_%A_index%
                {
                    IniWrite, % Val_statusWS_txt_%A_index%, settings.ini, Worksheets, statusWS_txt_%A_index%
                }
                else
                {
                    Val_statusWS_txt_%A_index% := "DDDDDD"
                    IniWrite, % Val_statusWS_txt_%A_index%, settings.ini, Worksheets, statusWS_txt_%A_index%
                }
            }
        }
        if % opsWS.length()
        {
            loop, % opsWS.length()
            {
                
                
                if % Val_opsWS_bcg_%A_index%
                {
                    IniWrite, % Val_opsWS_bcg_%A_index%, settings.ini, Worksheets, opsWS_bcg_%A_index%
                }
                else
                {
                    Val_opsWS_bcg_%A_index% := "000000"
                    IniWrite, % Val_opsWS_bcg_%A_index%, settings.ini, Worksheets, opsWS_bcg_%A_index%
                }
                if % Val_opsWS_txt_%A_index%
                {
                    IniWrite, % Val_opsWS_txt_%A_index%, settings.ini, Worksheets, opsWS_txt_%A_index%
                }
                else
                {
                    Val_opsWS_txt_%A_index% := "ffffff"
                    IniWrite, % Val_opsWS_txt_%A_index%, settings.ini, Worksheets, opsWS_txt_%A_index%
                }
            }
        }
        if % wingsWS.length()
        {
            loop, % wingsWS.length()
            {
                
                
                if % Val_wingsWS_bcg_%A_index%
                {
                    IniWrite, % Val_wingsWS_bcg_%A_index%, settings.ini, Worksheets, wingsWS_bcg_%A_index%
                }
                else
                {
                    Val_wingsWS_bcg_%A_index% := "000000"
                    IniWrite, % Val_wingsWS_bcg_%A_index%, settings.ini, Worksheets, wingsWS_bcg_%A_index%
                }
                if % Val_wingsWS_txt_%A_index%
                {
                    IniWrite, % Val_wingsWS_txt_%A_index%, settings.ini, Worksheets, wingsWS_txt_%A_index%
                }
                else
                {
                    Val_wingsWS_txt_%A_index% := "ffffff"
                    IniWrite, % Val_wingsWS_txt_%A_index%, settings.ini, Worksheets, wingsWS_txt_%A_index%
                }
            }
        }
    }
    
    if FileExist( "settings.ini" ) 
    {
        msgbox, File successfully written!
        ExitApp
    }
    return

GuiClose:
GuiCancel:
GuiEscape:
    ExitApp
    
ChooseColor(Color=0xF, hWnd=0x0, Flags=0x2 )  { ; CC_FULLOPEN := 0x2
    VarSetCapacity(CC,36+64,0), NumPut(36,CC), NumPut(hWnd,CC,4), NumPut(Color,CC,12)
    NumPut(&CC+36,CC,16), NumPut(Flags,CC,20), DllCall( "comdlg32\ChooseColorA", Str,CC ) 
    Hex:="123456789ABCDEF0",   RGB:=&CC+11 
    Loop 3  
        HexColorCode .=  SubStr(Hex, (*++RGB >> 4), 1) . SubStr(Hex, (*RGB & 15), 1) 
    ;Return HexColorCode  
    GuiControl, 1:, Val_%A_GuiControl%, %HexColorCode%
}