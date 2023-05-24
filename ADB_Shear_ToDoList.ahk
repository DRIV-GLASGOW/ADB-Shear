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
    listfile := "Data\ADBShearList.txt"
    iniRead, dbServer, %inifile%, Database, Server
    iniRead, lastUpdated, %inifile%, BasicData, LastUpdated
    firstPass := True
}



; ****************************************************************  GuiDefinition

GuiDefinition:
{
    GUI, 1:New, +OwnDialogs +resize, WindowTitle
    GUI, Color, 000000
    GUI, font, s14
    GUI, Add, Button, x10 y10 w150 h50 gGetList, Update List
    GUI, font, cWhite
    GUI, Add, Text, x175 y20 w400 vlastUpdated
    GUIControl,,lastUpdated,%lastUpdated%
    GUI, font, s12 c129500
    GUI, Add, Text, x10 y75, A ticket with this batch has run through finishing and needs a shear test done
    GUI, font, c707070
    GUI, Add, Text, x10 y95, A batch was made in Compound but no parts available for shear test
    
    GUI, Show, X510 Y100 w800 h300, ADB Shear Testing List
    
    If firstPass 
    {
        goSub, GetList
    }

    return
}


; ****************************************************************  HotKeys
^F1::
{   
    AboutTxt := "ADB Shear Testing To Do List`n"
    AboutTxt := AboutTxt . "Verison: " . version . "`n`n"
    AboutTxt := AboutTxt . "Database: " . dbserver . "`n"
    MsgBox, , ABOUT, % AboutTxt
    return
}


; ****************************************************************  Functions

GetList:
{
    GUI, +OwnDialogs
    formula := []
    julianDay := []
    finished := []
    
    try {
        currentDB := openDatabase("ADO",dbServer,"SPC","autohotkey","autohotkey")
    } catch e{
        MsgBox,4112, Error, % "Failed to create connection to SPC database. Check your Connection string and DB Settings!`n`n" ExceptionDetail(e)
    }
    
    sqlStr := "SELECT [Formula],[JulianDay],[ThroughFinish] FROM [SPC].[dbo].[Shear_NotCompleted] ORDER BY JulianDay asc, Formula"
    
    try{
        Recordset := currentDB.OpenRecordSet(sqlStr)
    } catch e {
        msgBox, 4112, Error, % "Error when opening the recordset."
    }
    
    i := 1
    try{
        while !Recordset.getEOF()
        {
            formula[i] := RecordSet["Formula"]
            julianDay[i] := RecordSet["JulianDay"]
            finished[i] := RecordSet["ThroughFinish"]
            i++
            RecordSet.MoveNext()
        }
    } catch e {
        msgBox, 16, Error, % "Error when pulling data from Shear_NotComplete table."
        return
    }
    
    numOfRec := i - 1

        Recordset.close()
        currentDb.close()

    gosub, WriteFile
    gosub, UpdateGUI
    
    FormatTime, lastUpdated,, MMM dd, yyyy hh:mm:ss tt
    GUIControl,, lastUpdated, %lastUpdated%
    
    return
}

WriteFile:
{
    FileDelete, %listfile%
    i := 1
    while i <= numOfRec
    {
        FileAppend, % formula[i] . "," . julianDay[i] . "," . finished[i] . ",`n", %listfile%
        i++
    }
    return
}

UpdateGUI:
{
    GUI, Destroy
    firstPass := False
    gosub, GuiDefinition
    
    GUI, font, s14
    xStart := 10
    xDel := 200
    xOff := 15
    yStart := 130
    yDel := 25
    column := 1
    maxRows := 15
    row := 0
    maxCol := 5
    yVarMax := 0
    overMax := False
    
    i := 1
    while(formula[i] != "" and column <= maxCol)
    {
        If(julianDay[i] != julianDay[i-1])
        {
            row += 1
            if(row >= maxRows)
            {
                column ++
                row := 1
            }
            
            if(column > maxCol)
            {
                overMax := True
                column --
                goto, OverMaxCol
            }

            xVar := xStart +(column - 1)*xDel
            yVar := yStart + (row-1)*yDel
            GUI, font, Underline Italic cWhite
            GUI, Add, Text, x%xVar% y%yVar%, % julianDay[i]
            row ++
        }
        
        If finished[i]
            GUI, font, norm c129500
        else
            GUI, font, norm c707070
        
        xVar := xStart + (column - 1)*xDel + xOff
        yVar := yStart + (row-1)*yDel
        GUI, Add, Text,  x%xVar% y%yVar%, % formula[i]
        row++
        
        if(yVar > yVarMax)
            yVarMax := yVar
        
        i++
    }
    
    OverMaxCol:
    hVar := yVarMax+ 100
    wVar := column*xDel + xStart
    if wVar < 800
        wVar := 800
    
    GUI, Show, X500 Y100 w%wVar% h%hVar%, ADB Shear Testing List
    
    if overMax
    {
        MsgBox,4112,, Too many records were returned.  Contact Engineering!
        yVar := yVarMax + 2*yDel
        GUI, font, cRed
        GUI, Add, Text, x10 y%yVar%, NOT ALL RECORDS RETURNED.  CONTACT ENGINEERING!!!
    }
    return
}

GUIClose:
{
    GUI, Submit, NoHide
    iniWrite, %lastUpdated%, %inifile%, BasicData, LastUpdated
    ExitApp
    
    return
}





