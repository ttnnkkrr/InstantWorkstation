
SetTitleMatchMode, 2

FileDelete,times.csv
FileDelete,errors.csv
		FileAppend, , times.csv
		FileAppend, , errors.csv
;~ data := clipboard
IniRead,excelfile,inilog,lastrecord,excelfile,false
if !FileExist( excelfile )
	{
	FileSelectFile,excelfile
	IniWrite,%excelfile%,inilog,lastrecord,excelfile
	}
oExcl := ComObjCreate("Excel.Application")
oExcl.visible := true
oExcl.workbooks.open( excelfile ) 
sTitle := "Care360"
oIE := oIE_get( sTitle )
oIE_HWND := oIE.hwnd
try
	{
		IniRead,lastrecord,inilog,lastrecord,lastRow,1
	loop, 
		{
		timerstart := A_TickCount
		if a_index = 1
			continue
		if (a_index < lastrecord)
			continue
		name := oExcl.activeworkbook.activesheet.cells(a_index,1).value ", " oExcl.activeworkbook.activesheet.cells(a_index,2).value
		dob := oExcl.activeworkbook.activesheet.cells(a_index,4).value 
		acct := oExcl.activeworkbook.activesheet.cells(a_index,5).value 
		if !acct
			break
		oIE.document.getElementsByTagName("input")["1"].value := name
		oIE.document.getElementsByTagName("input")["5"].value := dob
		oIE.document.getElementsByTagName("input")["6"].value := acct
		oIE.document.getElementsByTagName("button")["1"].click()
		
		waiting(oIE)
		;;; now lets click the link for the patient 
		if ( document.getElementsByClassName("gwt-HTML")["2"].innerText != "1 of 1" ) ;1 of 1
		oIE.document.getElementsByTagName("span")["4"].click()
		waiting(oIE)
		waiting(oIE)
		
		oIE.document.getElementById( "gwt-uid-376" ).click()
		waiting(oIE)
		waiting(oIE)
		try
		oIE.document.getElementsByClassName( "gwt-Image pointer" )["3"].click()
		waiting(oIE)
		waiting(oIE)
		
		oIE.document.getElementById( "reasonForDisclosureListBox" ).value := "26"
		oIE.document.getElementById( "pwText" ).value := "ymm1ysm2"
		oIE.document.getElementById( "reEnterPwText" ).value := "ymm1ysm2"
		
	
		waiting(oIE)
		try
			{
			
			tabscol := oIE.document.getElementsByClassName( "noBullets leftfloatChildrenLI tabBar")
			}
		catch e
			{
			;~ whait a while
			sleep 8000
			waiting(oIE)
			tabscol := oIE.document.getElementsByClassName( "noBullets leftfloatChildrenLI tabBar")
			}
		try
		loop % tabscol["0"].childNodes.length
			{
			node := A_Index - 1
			tabscol["0"].getElementsByTagName( "li")[node].getElementsByClassName( "gwt-HTML" )["0"].click()

			checkcount = 1
			loop % oIE.document.getElementsByClassName( "gwt-CheckBox" ).length 
				{
				inputs := A_Index - 1
				if ( oIE.document.getElementsByClassName( "gwt-CheckBox" )[inputs].getElementsByTagName( "input")["0"].type == "checkbox" )
					{
					oIE.document.getElementsByClassName( "gwt-CheckBox" )[inputs].getElementsByTagName( "input")["0"].checked := true
					;sleep 20
					}
				}	
			}
		catch e
			{
			if ( Patient_Consent_Message( oIE )  )
				continue
			if ( SNOMED( oIE )  )
				continue
			}
			
		if ( SNOMED( oIE )  )
			continue
		;~ actually export the record
		indexofsave := oIE.document.getElementsByClassName( "gwt-Button" ).length -2
		oIE.document.getElementsByClassName( "gwt-Button" )[indexofsave].click()
		waiting(oIE)
		waiting(oIE)
		
		
		if ( SNOMED( oIE )  )
			continue
		if ( Patient_Consent_Message( oIE )  )
			continue
		findpatientclick( oIE )
		WinActivate, ahk_id %oIE_HWND%
		WinWaitActive, ahk_id %oIE_HWND%
		MouseClick,,1008, 714
		
		
		totaltime := A_TickCount - timerstart
		
		FileAppend, Time for %name% was `,%totaltime%`r`n, times.csv
		IniWrite,%A_Index%,inilog,lastrecord,lastRow
		;~ throw { Message: "Custom error", what: "Custom error", file: A_LineFile, line: A_LineNumber }
		}
	oExcl.quit()
	}
catch e
{
	if ( e.message = "Consent")
		FileAppend, % e.what "`r`n", errors.csv
	else 
		FileAppend, % "Exception thrown!`n`nwhat: " e.what "`nfile: " e.file
        . "`nline: " e.line "`nmessage: " e.message "`nextra: " e.extra "`r`n", errors.csv
		
run errors.csv
    return
}
run times.csv
run errors.csv
msgbox done
ExitApp

ExitApp
#include IE functions.ahk




SNOMED( oIE ) ; occurs after attempt to export
	{
		

;~ Problems Not Coded in SNOMED

 



 



;~ This patient has problems that are not coded in SNOMED format and/or problems with auto-assigned SNOMED coding that have not been reviewed.

;~ Problems must be exported in SNOMED format to comply with meaningful use. SNOMED problems can be added or marked as reviewed on the patient Problems list: Patient Visit > Problems 

;~ Clicking Continue will export patient problems as is, without SNOMED codes and/or without review of auto-assigned SNOMED coding.

;~ Do you wish to continue? 

	try
	if InStr( oIE.document.getElementsByClassName( "popupPanelHeader" )["1"].innertext, "Problems Not Coded in SNOMED")
		{
		cancelDialog( oIE , "SNOMED", 1)
		indexofsave := oIE.document.getElementsByClassName( "gwt-Button" ).length -1
		oIE.document.getElementsByClassName( "gwt-Button" )[indexofsave].click()
		return true
		}
	catch e
		return false
	
	}
cancelDialog( oIE , m, popupPanelHeader)
	{
	global name,dob,acct
	FileAppend, % m " :: " name ":" dob ":" acct "`r`n", errors.csv
	try
		oIE.document.getElementsByClassName( "finalizeBar" )[popupPanelHeader].getElementsByTagName( "button")["1"].click()
	catch e
		throw { Message: m, what: "Couldnt cancel " oIE.document.getElementsByClassName( "finalizeBar" )[popupPanelHeader].outerHTML, file: A_LineFile, line: A_LineNumber }
	waiting(oIE)
	waiting(oIE)
	findpatientclick( oIE )
}

Patient_Consent_Message( oIE ) ; occurs after attempt to export
	{
	try
	if InStr( oIE.document.getElementsByClassName( "popupPanelHeader" )["0"].innertext, "Patient Consent Message")
		{
		;~ click cancel
		;~ Patient Consent Message
		;~ The patient's consent to share clinical documentation has expired. You can update it on the patient demographics page. 
		cancelDialog( oIE , "Consent", "0")
		return true
		}
	catch e
		return false
	
	}


findpatientclick( oIE )
	{
	
	try
		oIE.document.getElementsByClassName( "gwt-Anchor inline paddingLeft5 paddingRight5" )["2"].click()
	catch e
		try
			oIE.document.getElementsByClassName( "gwt-Anchor inline paddingLeft5 paddingRight5" )["1"].click()
		catch e
			try
				oIE.document.getElementsByClassName( "gwt-Anchor inline paddingLeft5 paddingRight5" )["0"].click()
			catch e
				throw { Message: "could not load search form", what: "Custom error", file: A_LineFile, line: A_LineNumber }
	waiting(oIE)
	waiting(oIE)
	
	}
