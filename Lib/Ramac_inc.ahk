#Include <DBA>

;~ /* Functions for interfacing with RAMAC */


{ ;~ getValue(Object db,String SQL, String Index)
;~ db = Database Resource
;~ String SQL = SQL Query String (i.e. "SELECT * FROM ramac_data WHERE 1;")
;~ Index = Name of index of column desired
}
getValue(db,SQL,Index=0) {
	
	returnValue := ""
	try {
		ret := db.OpenRecordSet(SQL)
		
		if (!ret.EOF) {
			returnValue := ret[Index]
		}
		ret.close()
	} catch e {
		; error message
		MsgBox,16, Error, % "Failed get the value!`n`n" ExceptionDetail(e)
	}
	return returnValue
}
;~ 
{ ;~ openDatabase(String dbType, String Server, String dBase, String uid, String pwd)
	;~ dbType = Type of Connection ("ADO", "mySQL", "SQLLite")
	;~ Server = Name of Server or location File
	;~ dbase = Name of Database to Use
	;~ uid = User Login
	;~ pwd = Password
}
openDatabase(dbType="ADO", Server="GLGWKYMS009",dBase="ramac",uid="autohotkey",pwd="autohotkey")
{
	try {
		result := DBA.DataBaseFactory.OpenDataBase(dbType, "Provider=sqloledb;Data Source=" . Server . ";Initial Catalog=" . dBase . ";User Id=" . uid . ";Password=" . pwd . ";")
	} catch e {
		; own message box
		MsgBox,16, Error, % "Failed to create connection. Check your Connection string and DB Settings!`n`n" ExceptionDetail(e)
	}
	return, result
}
		