;~ SeaMAX Libary
;~ Author: Mickey McClellan
;~ Date: 04/07/2017
;~ This Library will facilitate the use of the SeaLevel USB Relay. It has been designed for the Model 8112 Relay

;~ Constants
SIZE_CHAR := 1
SIZE_SHORT := 2
SIZE_INT := 4
SIZE_UINT := 4
SIZE_INT64 := 8
SIZE_UINT64 := 8
SIZE_FLOAT := 4
SIZE_DOUBLE := 8
SIZE_PTR := %A_PtrSize%

;~ DeviceConfig Structure per API Manual
;~ typedef struct configuration
;~ {
;~ int model;
;~ int commType;
;~ int baudrate;
;~ int parity;
;~ int firmware;
;~ } DeviceConfig;
DevCFG := 0
smhandle := 0
stateCoils := 0

Is64bit(){
    Return A_PtrSize = 8 ? 1 : 0
}
InitializeVariables() {
	;~ AHK doesn't not have structures, so NumGet and NumPuts must be used to access data via starting address and offset.
	;~ Declare Variable sizes for subsequent DLLCalls
	global smhandle
	global stateCoils
	global DevCFG
	if(is64bit())
	{
		;~ msgbox, ,,64bit
		VarSetCapacity(DevCFG, SIZE_UIN64*5)  ; set size of Device Configuraiton structure.  5 ints = 40.  
		VarSetCapacity(smhandle, SIZE_UINT64)	; set size ov smhandle variable Uint64 = 8 bytes
	} else {
		;~ msgbox,,,32bit
		VarSetCapacity(DevCFG, SIZE_UINT*5)  ; set size of Device Configuraiton structure.  5 ints = 40.  
		VarSetCapacity(smhandle, SIZE_UINT)	; set size ov smhandle variable Uint64 = 8 bytes
	}
	VarSetCapacity(stateCoils, 1) ; the 8112 only has 4 relays so 1 byte stores all the information required
}

InitializeSeaMax(ConnStr="SeaDAC Lite 0") { ; Defaults to first SeaDAC Lite connected to USB
	InitializeVariables()
	global DeviceConfiguration
	IfNotExist, Dll\Seamax.dll
		return -1
	
	
	;~ msgbox,,Debug,% "SDD Result: " . Result
	;~ msgbox,,Debug,LoadingLibrary
	;~ Load DLL in memory and keep it from being unloaded.  DllCall automatically loads and Frees the Libary if this isn't used.	
	if(is64bit())
	{
		hModule := DllCall("LoadLibrary", "Str", "Dll\\x64\SeaMax.dll", "Ptr") 
	} else {
		hModule := DllCall("LoadLibrary", "Str", "Dll\SeaMax.dll", "Ptr") 
	}
	;~ msgbox,,Debug,% "hModule: " . hModule
	If (hModule<= 0 ) {
		hModule := DllCall("LoadLibrary", "Str", "SeaMax.dll", "Ptr") 
		If (hModule<= 0 ) 
		{
			return -2
		}
	} 
	;~ msgbox,,Debug,SM_OPEN
	if(is64bit())
	{
		Result := DllCall("SeaMax.dll\SM_Open", "Uint64*" , smhandle ,"AStr", ConnStr)		
	} else {
		Result := DllCall("SeaMax.dll\SM_Open", "Uint64*" , smhandle ,"AStr", ConnStr)		
	}
	;~ msgbox,,Debug,% "Result: " . Result
	;~ msgbox,,Debug,% "SMHANDLE: " . smhandle
	If (Result != 0 ) {
		
		return -3
	} 
	;~ msgbox, ,,TEST1
	If (smhandle <= 0 ) {
		return %smhandle% - 1000
	} 
	;~ msgbox, ,,TEST2
	
	;~ msgbox,,Debug,GetSM_Config
	;~ Dev_CFG := GetSeaMaxConfig(smhandle)
	;~ msgbox,,Debug,% "CFG: " . Dev_CFG
	return %smhandle%
}
CloseSeaMax(hndl) {
	if(is64bit())
	{
	Result := DllCall("Dll\x64\SeaMax.dll\SM_Close", "Uint64" , hndl)
	} else {
	Result := DllCall("Dll\x64\SeaMax.dll\SM_Close", "Uint" , hndl)
	}
	return Result
}
ShutdownSeaMax(hndl) {
	CloseSeaMax(hndl)
	DLLCall("FreeLibrary", "Ptr", hModule) ;Free the memory used by the library
}
GetSeaMaxConfig(hndl) {
	handle := hndl
	;~ msgbox,,,% handle
	if(is64bit())
	{
		global SIZE_UINT64
		VarSetCapacity(DevCFG, SIZE_UINT64*5)  ; set size of Device Configuraiton structure.  5 ints = 40.  
		VarSetCapacity(handle, SIZE_UINT64)	; set size of smhandle variable Uint64 = 8 bytes
		Result := DllCall("SeaMAX\SM_GetDeviceConfig", "Uint64", handle, Ptr, &DevCFG)
	} else {
		global SIZE_UINT
		VarSetCapacity(DevCFG, SIZE_UINT*5)  ; set size of Device Configuraiton structure.  5 ints = 40.  
		VarSetCapacity(handle, SIZE_UINT)	; set size ov smhandle variable Uint64 = 8 bytes
		Result := DllCall("SeaMAX\SM_GetDeviceConfig", "Uint", handle, Ptr, &DevCFG)
	}
	
	If (Result < 0 ) {
		MsgBox,,Error, An error occured getting the device configuration. Error: %A_LastError% , %ErrorLevel%  RESULT: %Result%
	} 
	return DevCFG

}
ReadCoils(hndle){
	global stateCoils
	if(is64bit())
	{
	Result := DllCall("SeaMAX\SM_ReadDigitalOutputs", "Uint64", hndle, "Uint64", 0, "Uint64" , 4 , Ptr, &stateCoils)
	} else {
	Result := DllCall("SeaMAX\SM_ReadDigitalOutputs", "Uint", hndle, "Uint", 0, "Uint" , 4 , Ptr, &stateCoils)
	}
	return NumGet(stateCoils,0,"Char")
}
intUpdateCoils(hndle,value) {
	global stateCoils
	if ((value > 15) || (value < 0)) {
		return -255 ;; Invalid value
	}
	NumPut(value, stateCoils, 0, "Char")
	if(is64bit())
	{
		Result := DllCall("SeaMAX\SM_WriteDigitalOutputs", "Uint64", hndle, "Uint64", 0, "Uint64" , 4 , Ptr, &stateCoils)
	} else {
		Result := DllCall("SeaMAX\SM_WriteDigitalOutputs", "Uint", hndle, "Uint", 0, "Uint" , 4 , Ptr, &stateCoils)
	}
	return Result
}
UpdateCoils(hndle,Coil1, Coil2, Coil3, Coil4){
	global stateCoils
	if (Coil1 != 0) {
		Coil1 := 1
	}
	if (Coil2 != 0) {
		Coil2 := 1*2
	}
	if (Coil3 != 0) {
		Coil3 := 1*4
	}
	if (Coil4 != 0) {
		Coil4 := 1*8
	}
		
	coils := Coil1 + Coil2 + Coil3 + Coil4
	;~ stateCoils = Coil1 | Coil2 | Coil3 | Coil4
	;~ msgbox,,,% "Coils:" coils
	if(is64bit())
	{
		NumPut(coils, stateCoils, 0, "Char")
		Result := DllCall("SeaMAX\SM_WriteDigitalOutputs", "Uint64", hndle, "Uint64", 0, "Uint64" , 4 , Ptr, &stateCoils)
	} else {
		NumPut(coils, stateCoils, 0, "Char")
		Result := DllCall("SeaMAX\SM_WriteDigitalOutputs", "Uint", hndle, "Uint", 0, "Uint" , 4 , Ptr, &stateCoils)
	}
	
	return Result
}
GetSeaMaxModel(){
	if(is64bit())
	{
		return NumGet(DevCFG, 0, "Uint64") 
	} else {
		return NumGet(DevCFG, 0, "Uint") 
	}
			
	; Get the first int  
	; This is the only paramter returned in SeaMAX Lite products
}
GetSeaMaxCommType(){
	if(is64bit())
	{
		return NumGet(DevCFG, 8, "Uint64") 	
	} else {
		return NumGet(DevCFG, 8, "Uint") 	
	}
	
	; Get the second int
}
GetSeaMaxbaudrate(){
	if(is64bit())
	{
		return NumGet(DevCFG, 16, "Uint64") 		
	} else {
		return NumGet(DevCFG, 16, "Uint") 		
	}
	
	; Get the third int
}
GetSeaMaxparity(){
	if(is64bit())
	{
		return NumGet(DevCFG, 24, "Uint64") 
	} else {
		return NumGet(DevCFG, 24, "Uint32") 
	}
			
	; Get the fourth int
}
GetSeaMaxfirmware(){
	if(is64bit())
	{
		return NumGet(DevCFG, 32, "Uint64") 	
	} else {
		return NumGet(DevCFG, 32, "Uint") 	
	}
		
	; Get the fifth int
}

GetSeaMaxCoil1(hndle){
	ReadCoils(hndle)
	return stateCoils && 1
}
GetSeaMaxCoil2(hndle){
	ReadCoils(hndle)
	return stateCoils && 2
}
GetSeaMaxCoil3(hndle){
	ReadCoils(hndle)
	return stateCoils && 4
}
GetSeaMaxCoil4(hndle){
	ReadCoils(hndle)
	return stateCoils && 8
}
SetSeaMaxCoil(hndle, coil,state) {
	if ((coil > 4) || (coil < 1)){
		return -255
	}
	if (state != 0) {
		state := 1 << (coil - 1)
	}
	x := (1 << (coil-1))
	;~ msgbox,,, % x
	currentcoils := ReadCoils(hndle)
	coilmask := 15 - (1 << (coil-1))
	NewCoils := currentcoils  & coilmask
	NewCoils := NewCoils | state
	;~ MsgBox,,, % "CurrentCoils: " . currentcoils . "`nCoilMask: " . coilmask . "`nState: " . state . "`nNewCoils: " . NewCoils
	
	intUpdateCoils(hndle,NewCoils)
}