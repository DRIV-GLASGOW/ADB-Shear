;################### Serial_Port_Console_ReadFile.ahk ###################
; 8/11/09
; Assembled, tested (on WinXP), and debugged by aobrien with help from
;   other AHK forum members (especially Lexikos) and threads.
;
; This script is based upon dll structures that are built into the
; Microsoft Windows environment. This script is ugly (not very readable),
; because, it was written for demonstration purposes and I didn't want
; to complicate things by using #Include files.
;
; The most useful subroutines are listed below. Take them and write your
; application around them.
;     RS232_FileHandle:= RS232_Initialize(RS232_Settings) -- Get the filehandle
;     RS232_Close(RS232_FileHandle)
;     RS232_Read(RS232_FileHandle,"0xFF",RS232_Bytes_Received) -- 0xFF Size of receive buffer. This 
;                                returns HEX data (in ASCII form) received from the serial port.
;     Example: Read_Data := RS232_Read(RS232_FileHandle,"0xFF",RS232_Bytes_Received) ;if the RX buffer 
;              contained 0x11, 0x22, 0x00, 0x33, 0x44 then Read_Data will contain 1122003344
;
;     RS232_Write(RS232_FileHandle,Hex_Data) -- Comma delimited hex data. If I wanted to
;                               send "Hello World" I would do the following:
;     ;ASCII DATA=  H    e    l    l    o   SP    W    o    r    l    d
;     Hex_Data = 0x48,0x65,0x6C,0x6C,0x6F,0x20,0x57,0x6F,0x72,0x6C,0x64
;     RS232_Write(RS232_FileHandle,Hex_Data)
;
; Instructions:
; 1) Modify the RS232 port settings (under the User Variables heading)
;    to your needs and save the file.
;
; 2) Launch this script to connect to the RS232 COM Port.
;
; 3) CTRL-F1 to close the RS232 COM port and exit the receive loop.
;
; Script Behavior/Notes:
; * The script is designed to use a text editor (Notepad) to place the
;     received RS232 COM port characters.
; * When you attempt to type into the designated text editor the script
;     will capture the character and send it out the RS232 COM port. This is
;     accomplished with the Hotkey Assignments section.
; * Currently the script is written to only send/receive ASCII characters,
;     however, it would be REALLY EASY to modify the script so that it
;     will output/input data - something that HyperTerminal can't do.
; * When you first launch the script it will open Notepad and save it
;     using the Console_Path variable and a predetermined file name.
;
; !!!The Notepad text file MUST be saved so that the words
;      "COM1_Console.txt - Notepad" appear as the window
;      title, because the script will want to change to the window with
;      that name when it receives a character on the RS232 COM port.
;
;########################################################################
; ADAPTED TO INCLUDE FORMAT BY MICKEY MCCLELLAN 4/22/17




;;-- CUT  INCLUDE VARIABLES IN ACTUAL CODE
;########################################################################
;###### User Variables ##################################################
;########################################################################
;~ RS232_Port     = COM1
;~ RS232_Baud     = 115200
;~ RS232_Parity   = N
;~ RS232_Data     = 8
;~ RS232_Stop     = 1
;~ Console_Path =

;########################################################################
;###### Script Variables ################################################
;########################################################################
;~ RS232_Settings   = %RS232_Port%:baud=%RS232_Baud% parity=%RS232_Parity% data=%RS232_Data% stop=%RS232_Stop% dtr=Off

;########################################################################
;###### Function Templates###############################################
;########################################################################
;~ RS232_Initialize(RS232_Settings)
;~ RS232_Close(RS232_FileHandle)
;~ RS232_Write(RS232_FileHandle,Message)
;~ RS232_Read(RS232_FileHandle,Num_Bytes,ByRef RS232_Bytes_Received)
;;-- END CUT  INCLUDE VARIABLES IN ACTUAL CODE


;########################################################################
;###### Initialize RS232 COM Subroutine #################################
;########################################################################
RS232_Initialize(RS232_Settings)
{
  ;###### Extract/Format the RS232 COM Port Number ######
  ;7/23/08 Thanks krisky68 for finding/solving the bug in which RS232 COM Ports greater than 9 didn't work.
  StringSplit, RS232_Temp, RS232_Settings, `:
  RS232_Temp1_Len := StrLen(RS232_Temp1)  ;For COM Ports > 9 \\.\ needs to prepended to the COM Port name.
  If (RS232_Temp1_Len > 4)                   ;So the valid names are
    RS232_COM = \\.\%RS232_Temp1%             ; ... COM8  COM9   \\.\COM10  \\.\COM11  \\.\COM12 and so on...
  Else                                          ;
    RS232_COM = %RS232_Temp1%

  ;8/10/09 A BIG Thanks to trenton_xavier for figuring out how to make COM Ports greater than 9 work for USB-Serial Dongles.
  StringTrimLeft, RS232_Settings, RS232_Settings, RS232_Temp1_Len+1 ;Remove the COM number (+1 for the semicolon) for BuildCommDCB.
  ;MsgBox, RS232_COM=%RS232_COM% `nRS232_Settings=%RS232_Settings%

  ;###### Build RS232 COM DCB ######
  ;Creates the structure that contains the RS232 COM Port number, baud rate,...
  VarSetCapacity(DCB, 28)
  BCD_Result := DllCall("BuildCommDCB"
       ,"str" , RS232_Settings ;lpDef
       ,"UInt", &DCB)        ;lpDCB
  If (BCD_Result <> 1)
  {
    ;~ MsgBox, There is a problem with Serial Port communication. `nFailed Dll BuildCommDCB, BCD_Result=%BCD_Result%
    return -1
  }

  ;###### Create RS232 COM File ######
  ;Creates the RS232 COM Port File Handle
  RS232_FileHandle := DllCall("CreateFile"
       ,"Str" , RS232_COM     ;File Name         
       ,"UInt", 0xC0000000   ;Desired Access
       ,"UInt", 3            ;Safe Mode
       ,"UInt", 0            ;Security Attributes
       ,"UInt", 3            ;Creation Disposition
       ,"UInt", 0            ;Flags And Attributes
       ,"UInt", 0            ;Template File
       ,"Cdecl Int")

  If (RS232_FileHandle < 1)
  {
    ;~ MsgBox, There is a problem with Serial Port communication. `nFailed Dll CreateFile, RS232_FileHandle=%RS232_FileHandle% `nLast Error:  %A_LastError%`nThe script will now exit.
    ;~ Exit
    errorcode := %A_LASTERROR% * -1
    return errorcode
  }

  ;###### Set COM State ######
  ;Sets the RS232 COM Port number, baud rate,...
  SCS_Result := DllCall("SetCommState"
       ,"UInt", RS232_FileHandle ;File Handle
       ,"UInt", &DCB)          ;Pointer to DCB structure
  If (SCS_Result <> 1)
  {
    ;~ MsgBox, There is a problem with Serial Port communication. `nFailed Dll SetCommState, SCS_Result=%SCS_Result%
    RS232_Close(RS232_FileHandle)
    return -1
  }

  ;###### Create the SetCommTimeouts Structure ######
  ReadIntervalTimeout        = 0xffffffff
  ReadTotalTimeoutMultiplier = 0x00000000
  ReadTotalTimeoutConstant   = 0x00000000
  WriteTotalTimeoutMultiplier= 0x00000000
  WriteTotalTimeoutConstant  = 0x00000000

  VarSetCapacity(Data, 20, 0) ; 5 * sizeof(DWORD)
  NumPut(ReadIntervalTimeout,         Data,  0, "UInt")
  NumPut(ReadTotalTimeoutMultiplier,  Data,  4, "UInt")
  NumPut(ReadTotalTimeoutConstant,    Data,  8, "UInt")
  NumPut(WriteTotalTimeoutMultiplier, Data, 12, "UInt")
  NumPut(WriteTotalTimeoutConstant,   Data, 16, "UInt")

  ;###### Set the RS232 COM Timeouts ######
  SCT_result := DllCall("SetCommTimeouts"
     ,"UInt", RS232_FileHandle ;File Handle
     ,"UInt", &Data)         ;Pointer to the data structure
  If (SCT_result <> 1)
  {
    ;~ MsgBox, There is a problem with Serial Port communication. `nFailed Dll SetCommState, SCT_result=%SCT_result% 
    RS232_Close(RS232_FileHandle)
    return -1
  }
  
  Return %RS232_FileHandle%
}

;########################################################################
;###### Close RS23 COM Subroutine #######################################
;########################################################################
RS232_Close(RS232_FileHandle)
{
  ;###### Close the COM File ######
  CH_result := DllCall("CloseHandle", "UInt", RS232_FileHandle)
  If (CH_result <> 1)
    ;~ MsgBox, Failed Dll CloseHandle CH_result=%CH_result%

  Return
}

;########################################################################
;###### Write to RS232 COM Subroutines ##################################
;########################################################################
RS232_Write(RS232_FileHandle,Message)
{
  SetFormat, Integer, DEC

  ;Parse the Message. Byte0 is the number of bytes in the array.
  ;~ StringSplit, Byte, Message, `,
  StringSplit, Byte, Message
  Data_Length := Byte0
  ;MsgBox, Data_Length=%Data_Length% b1=%Byte1% b2=%Byte2% b3=%Byte3% b4=%Byte4%

  ;Set the Data buffer size, prefill with 0xFF.
  VarSetCapacity(Data, Byte0, 0xFF)

  ;Write the Message into the Data buffer
  i=1
  Loop %Byte0%
  {
    NumPut(Asc(Byte%i%), Data, (i-1) , "UChar")
    ;MsgBox, %i%
    i++
  }
  ;MsgBox, Data string=%Data%

  ;###### Write the data to the RS232 COM Port ######
  WF_Result := DllCall("WriteFile"
       ,"UInt" , RS232_FileHandle ;File Handle
       ,"UInt" , &Data          ;Pointer to string to send
       ,"UInt" , Data_Length    ;Data Length
       ,"UInt*", Bytes_Sent     ;Returns pointer to num bytes sent
       ,"Int"  , "NULL")
  ;~ If (WF_Result <> 1 or Bytes_Sent <> Data_Length)
    ;~ MsgBox, Failed Dll WriteFile to RS232 COM, result=%WF_Result% `nData Length=%Data_Length% `nBytes_Sent=%Bytes_Sent%
}

;########################################################################
;###### Read from RS232 COM Subroutines #################################
;########################################################################
RS232_Read(RS232_FileHandle,Num_Bytes,ByRef RS232_Bytes_Received)
{
  SetFormat, Integer, HEX

  ;Set the Data buffer size, prefill with 0x55 = ASCII character "U"
  ;VarSetCapacity won't assign anything less than 3 bytes. Meaning: If you
  ;  tell it you want 1 or 2 byte size variable it will give you 3.
  Data_Length  := VarSetCapacity(Data, Num_Bytes, 0x55)
  ;MsgBox, Data_Length=%Data_Length%

  ;###### Read the data from the RS232 COM Port ######
  ;~ MsgBox, RS232_FileHandle=%RS232_FileHandle% `nNum_Bytes=%Num_Bytes%
  Read_Result := DllCall("ReadFile"
       ,"UInt" , RS232_FileHandle   ; hFile
       ,"Str"  , Data             ; lpBuffer
       ,"Int"  , Num_Bytes        ; nNumberOfBytesToRead
       ,"UInt*", RS232_Bytes_Received   ; lpNumberOfBytesReceived
       ,"Int"  , 0)               ; lpOverlapped

  ;~ MsgBox, RS232_FileHandle=%RS232_FileHandle% `nRead_Result=%Read_Result% `nBR=%RS232_Bytes_Received% ,`nData=%Data%
  If (Read_Result <> 1)
  {
    ;~ MsgBox, There is a problem with Serial Port communication. `nFailed Dll ReadFile on RS232 COM, result=%Read_Result% - The Script Will Now Exit. `nLast Error: %A_LastError%
    RS232_Close(RS232_FileHandle)
    return -1
  }

  ;###### Format the received data ######
  ;This loop is necessary because AHK doesn't handle NULL (0x00) characters very nicely.
  ;Quote from AHK documentation under DllCall:
  ;     "Any binary zero stored in a variable by a function will hide all data to the right
  ;     of the zero; that is, such data cannot be accessed or changed by most commands and
  ;     functions. However, such data can be manipulated by the address and dereference operators
  ;     (& and *), as well as DllCall itself."
  i = 0
  Data_HEX =
  Loop %RS232_Bytes_Received%
  {
    ;First byte into the Rx FIFO ends up at position 0

    Data_HEX_Temp := NumGet(Data, i, "UChar") ;Convert to HEX byte-by-byte
    StringTrimLeft, Data_HEX_Temp, Data_HEX_Temp, 2 ;Remove the 0x (added by the above line) from the front

    ;If there is only 1 character then add the leading "0'
    Length := StrLen(Data_HEX_Temp)
    If (Length =1)
      Data_HEX_Temp = 0%Data_HEX_Temp%

    i++

    ;Put it all together
    Data_HEX := Data_HEX . Data_HEX_Temp
  }
  ;MsgBox, Read_Result=%Read_Result% `nRS232_Bytes_Received=%RS232_Bytes_Received% ,`nData_HEX=%Data_HEX%

  SetFormat, Integer, DEC
  Data := Data_HEX

  Return Data
}
