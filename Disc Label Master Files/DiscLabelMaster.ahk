#include M:\libshare\Disc Label Master\Disc Label Master Files\DiscLabelMasterInitialSection.ahk
#ifWinActive , SirsiDynix Symphony WorkFlows
numpadenter::
enter::
	if(!winActive("Modify Title") && modifyOnly)
		return
		
	displayMessage("Checking the Call Number/Item tab in workflows, please wait til finished...")
	Send !i
	Sleep 200
	Send I
	Send !s
	Sleep 500
	IfWinExist , Lookup
		return
	beforeChange()
	trueClick("CallTab","Click the Call Number/Item Tab")
	waitChange()	
	trueClick("CallNumberField","Click the Call Number Field")
	Sleep 100
	copyAll()
	if(InStr(clipboard, "`n") || StrLen(clipboard) > 45)
	{
		trueClick("CallTab","Click the Call Number/Item Tab")
		sleep 100
		trueClick("CallNumberField","Click the Call Number Field")
		Sleep 100
		
		displayMessage("Click the Call Number Field")
		copyAll()
	}
	clearMessage()
	displayMessage("Checking the Call Number/Item tab in workflows, please wait til finished...")
	currentCallNumber := clipboard
	
	
	if(inStr(currentCallNumber, "pt.") && !inStr(currentCallNumber , "|z"))
		checkCallNumber = 1
	else
		checkCallNumber = 0
		
	StringReplace , currentCallNumber , currentCallNumber , |z , %A_Space% , 1
	if (setSchemeAndLibrary)
	{
		trueClick("ClassSchemeField","Click the Class Scheme Field")
		send LC
		Sleep 50
		Send {ENTER}
		trueClick("ClassLibraryField","Click the Call Library Field")
		if(MCMode)
			Send LEE-LRC
		else
			Send MUSIC
		Sleep 50
		Send {ENTER}
	}
	trueClick("NumberOfPiecesField","Click the Number of Pieces Field")
	copyAll()
	if (checkCallNumber)
		MsgBox , 0 , %appName% , Make sure the "|z" is in the call number as appropriate`, then click OK.
	StringLeft , currentNumberOfPieces , clipboard , 2
	goto itemSpecificationsGuiShow
	
recordCorrectionCheck:	
	trueClick("NumberOfPiecesField","Click the Number of Pieces Field")
	copyAll()
	StringLeft , currentNumberOfPieces , clipboard , 2
	goto itemSpecificationsGuiShow

afterPromptContinueCheck:
	displayMessage("Making disc labels, please wait til finished...")
	appendToExcel(currentCallNumber, hasSupplement)	
		
	discCount := discCount + currentDiscCount
	howManyInARow := howManyInARow + currentDiscCount
	
	GuiControl , main: , discCount , %discCount%
	saveVariables()
	
	maxWorkFlows()
	ControlSend , , !u , SirsiDynix Symphony WorkFlows
	clearMessage()

	displayMessage("Making disc labels, please wait til finished...")
	
	if (useLabeler)
	{	
		Critical , on
		partCount = 1
		loop , %currentDiscCount%
		{
		
			if(MCMode)
				editDiscLabel("HBLL MC`n`n"currentCallNumber)
			else
				editDiscLabel("HBLL MLL`n`n"currentCallNumber)
				
			
			while(winExist("Text Effect"))
			{
				ControlSend , MVMINI2, {tab}{enter}, Text Effect
				sleep 500
			}
			ControlSend , , ^s , ahk_exe stdl.exe
			
			if(howManyInARow > 8)
			{
				displayMessage("Refreshing SureThing Disc Labeler, please wait til finished...")
				howManyInARow = 0
				loadingBlankSureThing = 1
				winClose , ahk_exe stdl.exe
				winWaitClose , ahk_exe stdl.exe
				sleep 200
				gosub runLabeler
				loadingBlankSureThing = 0
				winActivate , SirsiDynix Symphony WorkFlows
			}
			
			if (discCount <= 26)
				newDisc()
		}	
		Critical , off
	}
	currentDiscCount = 1
	saveVariables()
	if (discCount > 26)
	{
		
		GuiControl , main: , discCount , %discCount%
		saveVariables()
		if (useLabeler)
			goto printButtonG
	}
	endOfProcess:
	clearMessage()
	return
#ifWinActive
#include M:\libshare\Disc Label Master\Disc Label Master Files\DiscLabelMasterFunctionAndLabels.ahk