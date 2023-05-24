/*
    Author:   Clint Lister
    CoAuthor: 
    Language: English
    Platform: Windows 10Timer
    Company: Federal-Mogul Glasgow, KY
    Script function: Generate labels for 5-digit batch codes and record data in SQL

*/

SendMode Input
SetWorkingDir %A_ScriptDir%
#NoEnv


#include <DBA>
#include <Portlib>
#include <Ramac_INC>

#SingleInstance Force

; **************************************************************** CONFIGURATION
Configuration:
{
    Version = 23.05.23.1  ; Version = YY.MM.DD.rev
    inifile := "Data\ADBShear.ini" ;ini Configuration file location
    datafile := "Data\ShearGuiData.ini"
    iniRead, dbServer, %inifile%, Database, Server
    iniRead, numOfSlots, %iniFile%, BasicData, Num Of Slots
    rdPart := "R&D"
    
    i := 1
    while i <=numOfSlots
    {
        Order%i%_Active := false
        i++
    }
}



; ****************************************************************  GuiDefinition
{
    GuiDefinition:
    
    gosub, GetPartNoList
    GUIControl,,partNoList,%partNoList%
    GUIControl,,rdPart,%rdPart%
    
    GUI, 1:New, +OwnDialogs +resize +MinSize +LastFound, WindowTitle
    GUI, font, cFFFFFF
    GUI, Add, Picture, x720 y25 w130 h35, %A_WorkingDir%\Data\Tenneco-white.bmp
    GUI, Add, Text, x740 y70 vtest, Version: %Version%
    GUI, font, s12
    GUI, Add, Text, x10 y50, Part #: 
    GUI, Add, Text, x330 y50 w150 vcommonName
    GUI, Add, Text, x110 y80 w200 vformula
    GUI, Add, Text, x10 y80, Formula: 
    GUI, Add, Text, x10 y110, Tech:
    
    i:=1
    delta := 60
    while i <=numOfSlots
    {
        GUI,font,s12 cWhite
        yVar := (i-1)*delta+180
        GUI, Add, Text, x10 y%yVar%, Order#:
        
        yVar := (i-1)*delta+210
        GUI, Add, Text, x10 y%yVar%, StampCode:
        
        GUI,font,s10 c3578E5
        yVar := (i-1)*delta+180
        GUI, Add, Text, x315 y%yVar%, Pad1
        yVar := (i-1)*delta+208
        GUI, Add, Text, x315 y%yVar%, Pad2
        
        GUI,font,s12 cBlack
        yVar := (i-1)*delta+177
        GUI, Add, Edit, vOrder_%i% r1 w200 x110 y%yVar%
        GUIControl, disable, Order_%i%
        ;GUI, Add, Button, h25 w200 x110 y%yVar% gAddOrder vOrderButton_%i%, % Order_%i%
        ;GUIControl, disable, OrderButton_%i%
        yVar := (i-1)*delta+ 205
        GUI, Add, Edit, UpperCase r1 w200 x110 y%yVar% vStampCode_%i%
        GUIControl,disable,StampCode_%i%
        
        yVar := (i-1)*delta+234
        GUI, Add, Text, x10 y%yVar% w800 0x10
        
        yVar := (i-1)*delta+180
        GUI, Add, Radio, r1 w15 vscratch1OK_%i% x745 y%yVar%
        GUI, Add, Radio, r1 w15 vscratch1NG_%i% x775 y%yVar%
        GUIControl,disable, scratch1OK_%i%
        GUIControl,disable, scratch1NG_%i%
        
         yVar := (i-1)*delta+210
        GUI, Add, Radio, Group r1 w15 vscratch2OK_%i% x745 y%yVar%
        GUI, Add, Radio, r1 w15 vscratch2NG_%i% x775 y%yVar%
         GUIControl,disable, scratch2OK_%i%
        GUIControl,disable, scratch2NG_%i%
        
        i++
    }
    
    i:=1
    delta := 60
    while i <=numOfSlots
    {
        yVar := (i-1)*delta+177
        GUI, Add, Edit, Number r1 w60 vshear1L_%i% x360 y%yVar%
        yVar := (i-1)*delta+205
        GUI, Add, Edit, Number r1 w60 vshear2L_%i% x360 y%yVar%
        GUIControl,disable, shear1L_%i%
        GUIControl,disable, shear2L_%i%
        
        i++
    }
    
    i:=1
    delta := 60
    while i <=numOfSlots
    {
        yVar := (i-1)*delta+177
        GUI, Add, Edit, Number r1 w60 vshear1R_%i% x420 y%yVar%
        yVar := (i-1)*delta+205
        GUI, Add, Edit, Number r1 w60 vshear2R_%i% x420 y%yVar%
        GUIControl,disable, shear1R_%i%
        GUIControl,disable, shear2R_%i%
        
        i++
    }
    
    i:=1
    delta := 60
    while i <=numOfSlots
    {
        yVar := (i-1)*delta+177
        GUI, Add, Edit, Number r1 w50 vret1E_%i% x500 y%yVar%
        GUI, Add, Edit, Number r1 w50 vret1T_%i% x550 y%yVar%
        GUI, Add, Edit, r1 w60 vsg1_%i% x635 y%yVar%
        yVar := (i-1)*delta+205
        GUI, Add, Edit, Number r1 w50 vret2E_%i% x500 y%yVar%
        GUI, Add, Edit, Number r1 w50 vret2T_%i% x550 y%yVar%
        GUI, Add, Edit, r1 w60 vsg2_%i% x635 y%yVar%
        GUIControl,disable,ret1E_%i%
        GUIControl,disable,ret1T_%i%
        GUIControl,disable,sg1_%i%
        GUIControl,disable,ret2E_%i%
        GUIControl,disable,ret2T_%i%
        GUIControl,disable,sg2_%i%
        
        i++
    }
    
    
    GUI, font, c000000 s12
    GUI, Color, 000000, F0F0F0
    GUI, Add, DropDownList, w200 h200 vpartNo x110 y47 gGetSpecs, %partNoList%|%rdPart%
    GUI, Add, Edit, Limit2 Uppercase r1 w40 x55 y107 vtech


    GUI, font, s14 bold italic c3578E5
    GUI, Add, Text, x10 y10, ADB Shear Testing
    GUi, font, normal
    GUI,font,s12 italic c3578E5
    GUI, Add, Text, x360 y120, Shear
    GUI, Add, Text, x500 y120, Retention
    GUI, Add, Text, x625 y120, Sp Gravity
    GUI, Add, Text, x720 y120, Scratch Test
    GUI, font, s10 norm
    GUI, Add, Text, x360 y150, Left
    GUI, Add, Text, x420 y150, Right
    GUI, Add, Text, x500 y150, Edge
    GUI, Add, Text, x550 y150, Tot
    GUI, Add, Text, x740 y150, OK
    GUI, Add, Text, x770 y150, NG
    
    GUI, font, s10 cGreen
    GUI, Add, Text, x360 y100 w100 vshearMinLb
    GUI, Add, Text, x500 y100 w50 vshearRetEdge
    GUI, Add, Text, x550 y100 w50 vshearRetTotal
    GUI, Add, Text, x625 y100 w100 vsgRange, %sgMin% - %sgMax%
    
    
    GUI,font,s14 normal
    GUI, Add, button, x600 y30 w100 h60 vSendSQL gSendData, Send to SQL
    GUIControl, disable, SendSQL
    GUI,font,s10 normal
    GUI, Add, button, x110 y110 w100 h30, Reset Form
    
    
    GUI, Show, X100 Y100, ADB Shear Testing
    
    gosub, InitializeGUI

    return
}


; ****************************************************************  HotKeys
^F1::
{   
    AboutTxt := "ADB Shear Testing`n"
    AboutTxt := AboutTxt . "Verison: " . version . "`n`n"
    AboutTxt := AboutTxt . "Database: " . dbserver . "`n"
    MsgBox, , ABOUT, % AboutTxt
    return
}


; ****************************************************************  Functions

InitializeGUI:
{
    gosub, ReadDataFile
   
    GUIControl,, tech, %tech%
    
   if(partNo != "")
    {
       GUIControl, choose, partNo, %partNo%
        GUIControl,, formula, %formula%
        GUIControl,, commonName, %commonName%  
        gosub, GetSpecs
    }

    gosub, ReadDataFile
    GUIControl, enable, SendSQL
    
    i := 1
    while(i <= numOfSlots and Order%i%_Active = 1)
    {
        EnableOrder(i)
        GUIControl,,Order_%i%, % Order_%i%
        GUIControl,,StampCode_%i%, % StampCode_%i%
        GUIControl,,shear1L_%i%, % shear1L_%i%
        GUIControl,,shear1R_%i%, % shear1R_%i%
        GUIControl,,shear2L_%i%, % shear2L_%i%
        GUIControl,,shear2R_%i%, % shear2R_%i%
        GUIControl,,ret1E_%i%, % ret1E_%i%
        GUIControl,,ret1T_%i%, % ret1T_%i%
        GUIControl,,ret2E_%i%, % ret2E_%i%
        GUIControl,,ret2T_%i%, % ret2T_%i%
        GUIControl,,sg1_%i%, % sg1_%i%
        GUIControl,,sg2_%i%, % sg2_%i%
        GUIControl,, scratch1OK_%i%,% scratch1OK_%i%
        GUIControl,, scratch1NG_%i%,% scratch1NG_%i%
        GUIControl,, scratch2OK_%i%,% scratch2OK_%i%
        GUIControl,, scratch2NG_%i%,% scratch2NG_%i%
        GUIControl, disable, Order_%i%
        GUIControl, disable, StampCode_%i%
            
        i++
    }
            
    If(i <= numOfSlots and partNo != "")
    {
        GUIControl, enable, Order_%i%
        GUIControl, enable, StampCode_%i%
    }
    return
}

ButtonResetForm:
{
    MsgBox,4132,,All data on screen now will be deleted.  Are you sure you want to continue?
    IfMsgBox No
        return
    
    gosub, ResetForm
    return
}
    
ResetForm:
{
    rdPart := "R&D"
    GUIControl,,partNo,|%partNoList%|%rdPart%
    GUIControl,choose, partNo, 0
    GUIControl, enable, partNo
    formula :=
    commonName :=
    GUIControl,,formula,
    GUIControl,,commonName,
    GUIControl,,shearMinLb,
    GUIControl,,shearRetEdge,
    GUIControl,,shearRetTotal,
    GUIControl,, sgRange,
    
    gosub, ResetRed

    i := 1
    while(i <= numOfSlots)
    {
        DisableOrder(i)
        
        i++
    }
    
    rdActive := False
    
    GUI, submit, NoHide
    
    return
}

#IfWinActive ADB Shear Testing
Enter::
NumpadEnter::
{
    startChar := 0
    GUIControlGet, currentCont, FocusV
    If InStr(currentCont,"StampCode")
    {
        startChar := 11
        gosub, AddOrder
        return
    }
    
    If InStr(currentCont, "Order")
    {
        startChar := 7
        gosub, AddOrder
        return
    }
    
    return
}

AddOrder:
{
   GUI, +OwnDialogs
   GUI, submit, NoHide
   slot := SubStr(currentCont, startChar) 

    If rdActive
    {
        goto, SkipChecks
    }

    If(StrLen(Order_%slot%) != 8)
    {
        MsgBox,4112,,% Order_%slot% . " is not a valid Order #.  Try Again."
        GUIControl,,Order_%slot%,
        GUIControl,focus,Order_%slot%
        return
    }
    
    If Order_%slot% is not integer
    {
        MsgBox,4112,,% Order_%slot% . " is not a valid Order #.  Try Again."
        GUIControl,,Order_%slot%,
        GUIControl,focus,Order_%slot%
        return
    }
        
    stampCodeA := SubStr(StampCode_%slot%, 1,4)
    stampCodeB := SubStr(StampCode_%slot%, 5,1)
    
    If(StrLen(StampCode_%slot%) = 5)
    {
        If stampCodeA is digit
        {
            if stampCodeB is alpha
            {
                goto, StampOK
            }
        }
    }
        MsgBox,4112,,% StampCode_%slot% . " is not a valid Stamp Code.  Try Again."
        GUIControl,,StampCode_%slot%
        GUIControl,focus,StampCode_%slot%
        return

    StampOK:
    stampCodeErr := GetStampCode(Order_%slot%, slot, dbServer,formula, StampCode_%slot%)
    If(stampCodeErr < 0)
    {
        Order_%slot% := ""
        GUIControl,,OrderButton_%slot%,
        return
    }
    
    GUIControl,disable,Order_%slot%
    GUIControl,disable,StampCode_%slot%
    
    SkipChecks:
    EnableOrder(slot)
    
    Order%slot%_Active := True
    
    i := slot + 1
    If i <= numOfSlots
    {
        GUIControl,enable,Order_%i%
        GUIControl,focus,Order_%i%
        GUIControl,enable,StampCode_%i%
    }
        
    return
}

EnableOrder(slot)
{
    GUIControl,enable,shear1L_%slot%
    GUIControl,enable,shear1R_%slot%
    GUIControl,enable,shear2L_%slot%
    GUIControl,enable,shear2R_%slot%
    GUIControl,enable,ret1E_%slot%
    GUIControl,enable,ret1T_%slot%
    GUIControl,enable,ret2E_%slot%
    GUIControl,enable,ret2T_%slot%
    GUIControl,enable,sg1_%slot%
    GUIControl,enable,sg2_%slot%
    GUIControl,enable,scratch1OK_%slot%
    GUIControl,enable,scratch1NG_%slot%
    GUIControl,enable,scratch2OK_%slot%
    GUIControl,enable,scratch2NG_%slot%
    
    return
}

DisableOrder(slot)
{
    GUIControl,disable,shear1L_%slot%
    GUIControl,disable,shear1R_%slot%
    GUIControl,disable,shear2L_%slot%
    GUIControl,disable,shear2R_%slot%
    GUIControl,disable,ret1E_%slot%
    GUIControl,disable,ret1T_%slot%
    GUIControl,disable,ret2E_%slot%
    GUIControl,disable,ret2T_%slot%
    GUIControl,disable,sg1_%slot%
    GUIControl,disable,sg2_%slot%
    GUIControl,disable,scratch1OK_%slot%
    GUIControl,disable,scratch1NG_%slot%
    GUIControl,disable,scratch2OK_%slot%
    GUIControl,disable,scratch2NG_%slot%
    GUIControl,disable,Order_%slot%
    GUIControl,disable,StampCode_%slot%
    GUIControl,,shear1L_%slot%,
    GUIControl,,shear1R_%slot%,
    GUIControl,,shear2L_%slot%,
    GUIControl,,shear2R_%slot%,
    GUIControl,,ret1E_%slot%,
    GUIControl,,ret1T_%slot%,
    GUIControl,,ret2E_%slot%,
    GUIControl,,ret2T_%slot%,
    GUIControl,,sg1_%slot%,
    GUIControl,,sg2_%slot%,
    GUIControl,,scratch1OK_%slot%,0
    GUIControl,,scratch1NG_%slot%,0
    GUIControl,,scratch2OK_%slot%,0
    GUIControl,,scratch2NG_%slot%,0
    GUIControl,,Order_%slot%,
    GUIControl,,StampCode_%slot%
    
    Order%slot%_Active := False
    
    GUI, submit, NoHide
    
    return
}

GetPartNoList:
{
    try {
        currentDB := openDatabase("ADO",dbServer,"AirDiscScrap","autohotkey","autohotkey")
    } catch e{
        MsgBox,4112, Error, % "Failed to create connection. Check your Connection string and DB Settings!`n`n" ExceptionDetail(e)
    }
    
    sqlStr := "EXECUTE [AirDiscScrap].[dbo].[GetPartNoList] "
    
    try{
        Recordset := currentDB.OpenRecordSet(sqlStr)
    } catch e {
        msgBox, 4112, Error, % "No Part #s found!"
    }
    
    try{
        while !Recordset.getEOF()
        {
            partNoList .= RecordSet["PartNo"]
            RecordSet.MoveNext()
            if !Recordset.getEOF()
            {
                partNoList .= "|"
            }
        }
    } catch e {
        msgBox, 16, Error, % "Error when retrieving part # list from SQL"
        return
    }

        Recordset.close()
        currentDb.close()
        
        return
}

GetSpecs:
{
    GUI,+OwnDialogs,
    GUIControl,,formula,
    GUIControl,,commonName,
    GUI,submit,NoHide
    rdActive := False
    
    If(partNo = rdPart)
    {
        gosub, SetupRD
        rdActive := True
        return
    }
    
    rdPart := "R&D"
    GUIControl,,partNo,|%partNoList%|%rdPart%
    GUIControl,choose, partNo, %partNo%
    
    try {
        currentDB := openDatabase("ADO",dbServer,"AirDiscScrap","autohotkey","autohotkey")
    } catch e{
        MsgBox,4112, Error, % "Failed to create connection. Check your Connection string and DB Settings!`n`n" ExceptionDetail(e)
        gosub, ResetPartNo
        return
    }
    
    sqlStr := "SELECT TOP 1 [SAP_Part_Number],[ProdSpec],[CommonName],[Formula],[ShearAmb_Min_lb],[ShearAmb_Min_daN],[ShearAmb_Ret_Edge],[ShearAmb_Ret_Total],[SG_Min],[SG_Max],[PadArea],[TheoSG] FROM [AirDiscScrap].[dbo].[Shear_Data] WHERE SAP_Part_Number = '" . partNo . "'"
    
    try{
        rs := currentDB.OpenRecordSet(sqlStr)
    } catch e {
        msgBox, 4112, Error, % "Could not open recordset from SQL."
        gosub, ResetPartNo
        return
    }
    
    try{
            prodSpec := rs["ProdSpec"]
            commonName := rs["CommonName"]
            formula := rs["Formula"]
            shearMinLb  := rs["ShearAmb_Min_lb"]
            shearMinDaN := rs["ShearAmb_Min_daN"]
            shearRetEdge := rs["ShearAmb_Ret_Edge"]
            shearRetTotal := rs["ShearAmb_Ret_Total"]
            sgMin := Format("{:.2f}", rs["SG_Min"])
            sgMax := Format("{:.2f}", rs["SG_Max"])
            padArea := Format("{:.1f}", rs["PadArea"])
            theoSG := rs["TheoSG"]
        }catch e {
        msgBox, 16, Error, % "Error when retrieving specs from SQL table."
        gosub, ResetPartNo
        return
        }
        
        GUIControl,,commonName,%commonName%
        GUIControl,,formula,%formula%
        GUIControl,,shearMinLb,%shearMinLb%
        GUIControl,,shearRetEdge,%shearRetEdge%
        GUIControl,,shearRetTotal,%shearRetTotal%
        GUIControl,,sgRange,%sgMin% - %sgMax%

        Recordset.close()
        currentDb.close()
        
        GUIControl,enable, Order_1
        GUIControl,focus, Order_1
        GUIControl,enable, StampCode_1
        GUIControl,disable,partNo
    
    return
}


ResetPartNo:
{
    rdPart := "R&D"
    GUIControl,,partNo,|%partNoList%|%rdPart%
    GUIControl,choose, partNo, 0
    GUIControl, enable, partNo
    GUIControl,,formula,
    GUIControl,,commonName,
    GUIControl,,shearMinLb,
    GUIControl,,shearRetEdge,
    GUIControl,,shearRetTotal,
    GUIControl,, sgRange,
    return
}

GetStampCode(orderNo,slot,dbServer,formula, stampCode)
{
    ;get the batch num associated with this order num
    try {
        currentDB := openDatabase("ADO",dbServer,"ramac","autohotkey","autohotkey")
    } catch e{
        MsgBox,4112, Error, % "Failed to create connection to ramac SQL table. Contact Engineering!`n`n" ExceptionDetail(e)
        return, -1
    }
    
    sqlStr := "SELECT TOP 1 [OrderNumber],[BatchID],[RunTime] FROM [ramac].[dbo].[ADB_Batch_Reference] WHERE OrderNumber = '" . orderNo . "'"
    
    try{
        rs := currentDB.OpenRecordSet(sqlStr)
    } catch e {
        msgBox, 4112, Error, % "Could not open ADB Batch Reference recordset from SQL."
        return, -1
    }
    
    try{
            db_stampCode := SubStr(rs["BatchID"],5,4) . SubStr(rs["BatchID"],10,1)
            batchNo := rs["BatchID"]
    }catch e {
        msgBox, 4112, Error, % orderNo . " does not exist in ADB Batch Reference Table."
        return, -1
    }
    
    ; if batch assigned to the order in ADB Batch Reference is not the same as that entered by the operator, cancel entry
    If(db_stampCode != stampCode)
    {
        MsgBox,4144,,% "Order " . Order_%slot% . " was assigned Stamp Code`n`n" . db_stampCode . "`n`nin the database.`n`nThis does not match what was entered."  
        ;Check the pad you're testing to be sure it matches what you entered.`nClick OK to continue with the Stamp Code you've entered.`nClick Cancel to go back and fix."
        ;IfMsgBox Cancel
            return, -1
    }
    
    ;find the formula associated with this batch num
    try {
        currentDB := openDatabase("ADO",dbServer,"BizWare","autohotkey","autohotkey")
    } catch e{
        MsgBox,4112, Error, % "Failed to create connection to BizWare SQL table. Contact Engineering!`n`n" ExceptionDetail(e)
        ResetSlot(slot)
        return, -1
    }
    
    sqlStr := "SELECT TOP 1 [BatchID],[Formula] FROM [BizWare].[dbo].[Compound5_Batch_Data] WHERE BatchID = '" . batchNo . "'"
    
    try{
        rs := currentDB.OpenRecordSet(sqlStr)
    } catch e {
        msgBox, 4112, Error, % "Could not open Compound5_Batch_Data recordset from SQL."
        ResetSlot(slot)
        return, -1
    }
    
    try{
            batchFormula := rs["Formula"]
    }catch e {
        msgBox, 4112, Error, % batchNo . " does not exist in Compound 5 Batch Data."
        ResetSlot(slot)
        return, -1
    }
    
    If(batchFormula != formula)
    {
        MsgBox,4113,,% "Order " . Order_%slot% . " was assigned batch " . batchNo . " which is formula`n`n" . batchFormula . "`n`nThis does not match the part number entered which is formula`n`n" . formula . "`n`nCheck the pad you're testing to be sure it matches what you entered.`nClick OK to continue with the Stamp Code you've entered.`nClick Cancel to go back and fix."
        IfMsgBox Cancel
            return, -2
    }

    GUIControl,,stampCode_%slot%, % stampCode_%slot%

    rs.close()
    currentDb.close()
    
    return, 0
}

SetupRD:
{
    InputBox, rdPart,,Type or Scan a part # for this R&&D check.
    If ErrorLevel
    {
        goSub, ResetPartNo
        return
    }
    
    GUIControl,,partNo,|%partNoList%|%rdPart%
    GUIControl,choose, partNo, %rdPart%
    commonName := "R&&D"
    GUIControl,,commonName,%commonName%
    
    InputBox, formula,,Type or Scan a formula for this R&&D check.
    If ErrorLevel
    {
        gosub, ResetPartNo
        return
    }
    
    GUIControl,,formula,%formula%
    
    GUIControl,enable,Order_1
    GUIControl,enable, StampCode_1
    GUIControl,disable,partNo
    
    
    return
}

ResetSlot(slot)
{
    GUIControl,,Order_%slot%,
    GUIControl,,stampCode_%slot%,
    
    GUI,submit,NoHide
    return
}

SendData:
{
    GUI,submit,NoHide
    GUI, +OwnDialogs
    ; reset all entered data to black text before rechecking for out of tolerance
    gosub, ResetRed
    ; check entered data against tolerances
    validData := False
    gosub, ValidateData
    
    ; if bad data types found or user chooses to edit typos, exit the thread
    If !validData
        return
    
     try {
        currentDB := openDatabase("ADO",dbServer,"SPC","autohotkey","autohotkey")
    } catch e{
        MsgBox,4112, Error, % "Failed to create connection to SPC database. Check your Connection string and DB Settings!`n`n" ExceptionDetail(e)
    }
   
    i = 1
    while i <= numOfSlots and Order%i%_Active
    {
        j := 1
        while j<=2
        {
            If scratch%j%OK_%i%
                scratchString := "OK"
            else
                scratchString := "NG"
            
            sqlStr := "EXECUTE  [SPC].[dbo].[SaveShearData]  @PartNo = '" . partNo . "', @StampCode = '" . StampCode_%i% . "', @OrderNo = '" . Order_%i% . "', @Oper = '" . tech . "', @ShearMin = '" . shearMinLb . "', @ShearL = '" . shear%j%L_%i% . "', @ShearR = '" . shear%j%R_%i% . "', @ShearMinDaN = '" . shearMinDaN . "', @RetEdgeMin = '" . shearRetEdge . "',  @RetEdge= '" . ret%j%E_%i% . "', @RetTotMin = '" . shearRetTotal . "', @RetTot = '" . ret%j%T_%i% . "', @SGMin = '" . sgMin . "', @SGMax = '" . sgMax . "', @SG = '" . sg%j%_%i% . "', @Scratch = '" . scratchString . "', @PadArea = '" . padArea . "', @TheoSG = '" . theoSG . "', @Formula = '" . formula . "';"
            
            try
            {
                Recordset := currentDB.OpenRecordSet(sqlStr)

            } catch e {
                msgBox, 4112, Error, Error when inserting data.
                msgBox, 4112, Error, NO DATA RECORDED!
                return
            }
             j++
        }
         i++
    }
       
    MsgBox,4112,, Data Successfully Recorded!
    
    
        Recordset.close()
        currentDb.close()

    gosub, ResetForm
    
    return
}


ValidateData:
{
    ;set all out of tol bits to false
    shearOut := false
    retOut := false
    sgOut := false
    scratchOut := false
    
    ;ensure a part # was set
    If(partNo = "" or commonName = "" or formula = "")
    {
        MsgBox,4112,,"No Part No selected."
        return
    }
    
    ;check entered values for type and out of tolerance, cycle through each active slot
    i := 1
    while i<=numOfSlots
    {
        ; only perform these checks if the slot has been activated
        If Order%i%_Active
        {
            ; ensure that no alpha characters were put in the SG results
             sgBadType := false
        
            temp1 := sg1_%i%
            temp2 := sg2_%i%
            If temp1 is not number
            {
                msgbox,,, % temp1
                sgBadType := true
                TurnRed("sg1_" . i)
            }
            If temp2 is not number
            {
                sgBadType := true
                TurnRed("sg2_" . i)
            }
        
            If(sgBadType)
            {
                MsgBox, 4112,,% "One or more SG results for " . Order_%i% . " include non-numeric charcters.  Fix these and try again."
                return
            }
            
            If((!scratch1OK_%i% and !scratch1NG_%i%) or (!scratch2OK_%i% and !scratch2NG_%i%))
            {
                MsgBox, 4112,, % "No Scratch result selected for Order " . Order_%i% . "`nSelect one and try again."
                return
            }
        
            ; set arrays to cycle through and check against tolerance limit.  dataArray is the values entered.  ctrlArray is the name of the GUI control that holds those values.  Must be in same order.
            dataArray := [shear1L_%i%, shear1R_%i%, shear2L_%i%, shear2R_%i%, ret1E_%i%, ret1T_%i%, ret2E_%i%, ret2T_%i%, sg1_%i%, sg2_%i%, scratch1OK_%i%, scratch1NG_%i%, scratch2OK_%i%, scratch2NG_%i%]
            ctrlArray := []
            ctrlArray[1] := "shear1L_" . i
            ctrlArray[2] := "shear1R_" . i
            ctrlArray[3] := "shear2L_" . i
            ctrlArray[4] := "shear2R_" . i
            ctrlArray[5] := "ret1E_" . i
            ctrlArray[6] := "ret1T_" . i
            ctrlArray[7] := "ret2E_" . i
            ctrlArray[8] := "ret2T_" . i
            ctrlArray[9] := "sg1_" . i
            ctrlArray[10] := "sg2_" . i
            ctrlArray[11] := "scratch1OK_" . i
            ctrlArray[12] := "scratch1NG_" . i
            ctrlArray[13] := "scratch2OK_" . i
            ctrlArray[14] := "scratch2NG_" . i
            
            ; go through each data entry, check what kind it is and check against the appropriate tol.  If out, set the text to red.
            For j, element in ctrlArray
            {
                If InStr(ctrlArray[j], "shear")
                {
                    If dataArray[j] < shearMinLb
                    {
                        TurnRed(ctrlArray[j])
                        shearOut := true
                    }
                }
                
                If InStr(ctrlArray[j], "ret")
                {
                    If SubStr(ctrlArray[j], 5, 1) = "E"
                    {
                        If dataArray[j] < shearRetEdge
                        {
                            TurnRed(ctrlArray[j])
                            retOut := true
                        }
                    }
                    If SubStr(ctrlArray[j], 5, 1) = "T"
                    {
                        If dataArray[j] < shearRetTotal
                        {
                            TurnRed(ctrlArray[j])
                            retOut := true
                        }
                    }
                }
                
                If InStr(ctrlArray[j], "sg")
                {
                    If((dataArray[j] < sgMin) or (dataArray[j] > sgMax))
                    {
                        TurnRed(ctrlArray[j])
                        sgOut := true
                    }
                }
                
                If(j = 11 or j = 13)
                {
                    If dataArray[j] = 0 
                    scratchOut := true
                }
            }
        }
        i++
    }
        
        If(shearOut or retOut or sgOut or scratchOut)
        {
            MsgBox,4113,, Some entered data is out of spec.  Check for NG selected on Scratch Tests and any red highlighted values.`n`nIf you need to correct a typo, click Cancel.`nIf data is correct, click OK.   AND CONTACT QA MANAGER!!!!
            IfMsgBox Cancel
                return
        }
    
    validData := true
    
    return
}

TurnRed(datapt)
{
    GUI, font, cE00000 s12
    GUIControl,font, %datapt%
    
    return
}

ResetRed:
{
    GUI, font, cBlack s12
    
    i := 1
    while i <= numOfSlots
    {
        GUIControl, font, shear1L_%i%
        GUIControl, font, shear1R_%i%
        GUIControl, font, shear2L_%i%
        GUIControl, font, shear2R_%i%
        GUIControl, font, ret1E_%i%
        GUIControl, font, ret1T_%i%
        GUIControl, font, ret2E_%i%
        GUIControl, font, ret2T_%i%
        GUIControl, font, sg1_%i%
        GUIControl, font, sg2_%i%
        
        i++
    }
    return
}
        
        

GUIClose:
{
    gosub, WriteDataFile
    ExitApp
}

ReadDataFile:
{
    iniRead, partNo, %datafile%, BasicData, PartNo
    iniRead, formula, %datafile%, BasicData, Formula
    iniRead, commonName, %datafile%, BasicData, Common Name
    iniRead, tech,%datafile%, BasicData, Tech
    
    i := 1
    while i <= numOfSlots
    {
        iniRead, Order%i%_Active, %datafile%, Order%i%, Active,0
        iniRead, Order_%i%, %datafile%, Order%i%, OrderNo
        iniRead, StampCode_%i%, %datafile%, Order%i%, StampCode
        iniRead, shear1L_%i%, %datafile%, Order%i%, Shear L Pad 1
        iniRead, shear1R_%i%, %datafile%, Order%i%, Shear R Pad 1
        iniRead, shear2L_%i%, %datafile%, Order%i%, Shear L Pad 2
        iniRead, shear2R_%i%, %datafile%, Order%i%, Shear R Pad 2
        iniRead, ret1E_%i%, %datafile%, Order%i%, Ret Edge Pad 1
        iniRead, ret1T_%i%, %datafile%, Order%i%, Ret Total Pad 1
        iniRead, ret2E_%i%, %datafile%, Order%i%, Ret Edge Pad 2
        iniRead, ret2T_%i%, %datafile%, Order%i%, Ret Total Pad 2
        iniRead, sg1_%i%, %datafile%, Order%i%, SG Pad 1
        iniRead, sg2_%i%, %datafile%, Order%i%, SG Pad 2
        iniRead, scratch1OK_%i%, %datafile%, Order%i%, Scratch OK Pad 1
        iniRead, scratch1NG_%i%, %datafile%, Order%i%, Scratch NG Pad 1
        iniRead, scratch2OK_%i%, %datafile%, Order%i%, Scratch OK Pad 2
        iniRead, scratch2NG_%i%, %datafile%, Order%i%, Scratch NG Pad 2
        i++
    }
    
    return
}

WriteDataFile:
{
    GUI, submit, NoHide
    iniWrite, %partNo%, %datafile%, BasicData, PartNo
    iniWrite, %formula%, %datafile%, BasicData, Formula
    iniWrite, %commonName%, %datafile%, BasicData, Common Name
    iniWrite, %tech%, %datafile%, BasicData, Tech
    
    i := 1
    while i <= numOfSlots
    {
        iniWrite, % Order%i%_Active, %datafile%, Order%i%, Active
        iniWrite, % Order_%i%, %datafile%, Order%i%, OrderNo
        iniWrite, % StampCode_%i%, %datafile%, Order%i%, StampCode
        iniWrite, % shear1L_%i%, %datafile%, Order%i%, Shear L Pad 1
        iniWrite, % shear1R_%i%, %datafile%, Order%i%, Shear R Pad 1
        iniWrite, % shear2L_%i%, %datafile%, Order%i%, Shear L Pad 2
        iniWrite, % shear2R_%i%, %datafile%, Order%i%, Shear R Pad 2
        iniWrite, % ret1E_%i%, %datafile%, Order%i%, Ret Edge Pad 1
        iniWrite, % ret1T_%i%, %datafile%, Order%i%, Ret Total Pad 1
        iniWrite, % ret2E_%i%, %datafile%, Order%i%, Ret Edge Pad 2
        iniWrite, % ret2T_%i%, %datafile%, Order%i%, Ret Total Pad 2
        iniWrite, % sg1_%i%, %datafile%, Order%i%, SG Pad 1
        iniWrite, % sg2_%i%, %datafile%, Order%i%, SG Pad 2
        iniWrite, % scratch1OK_%i%, %datafile%, Order%i%, Scratch OK Pad 1
        iniWrite, % scratch1NG_%i%, %datafile%, Order%i%, Scratch NG Pad 1
        iniWrite, % scratch2OK_%i%, %datafile%, Order%i%, Scratch OK Pad 2
        iniWrite, % scratch2NG_%i%, %datafile%, Order%i%, Scratch NG Pad 2
        
        i++
    }
    return
}





