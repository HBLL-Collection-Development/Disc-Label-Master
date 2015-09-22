;Initial Script
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
#SingleInstance force
CoordMode , Mouse , Relative
CoordMode , Pixel , Relative
FileCreateDir , BiblioScriptDVDProjectFiles
SetWorkingDir , BiblioScriptDVDProjectFiles
	if ( errorLevel = 1 )
		MsgBox Error Creating Workspace Directory
loadVariables()
Menu , Tray, Tip, BiblioScript DVD Project

Gui , 3:+AlwaysOnTop -MaximizeBox -MinimizeBox
Gui , 3:Add , Text , , Current Disc Number:
Gui , 3:Add , Edit , x+5 yp-3 center vdiscCount gdiscRange number range1-99 limit2 w30, %discCount%
Gui , 3:Add, Radio, xm y+15 group vMCMode checked%MCMode%, HBLL MC (DVDs)
Gui , 3:Add, Radio, xm vMLLMode checked%MLLMode%, HBLL MLL (Music discs)
Gui , 3:Add , Checkbox , ym vsetSchemeAndLibrary Checked%setSchemeAndLibrary%, &Set "LC" and "LEE-LRC"`n(Modify title only)
Gui , 3:Add , Checkbox , vuseExcel Checked%useExcel%, Handle &Excel spreadsheet
Gui , 3:Add , Checkbox , vuseLabeler Checked%useLabeler%, Handle &SureThing Labeler
Gui , 3:Add , Checkbox , vuseMousePositions Checked%useMousePositions%, &Use saved mouse positions
Gui , 3:Add , Button , w300 xm , &Clear Mouse Positions
Gui , 3:Add , Button , xm Default w300 , &OK
;gui , 3:show
;gosub properties

Gui , 4:+AlwaysOnTop -MaximizeBox -MinimizeBox
Gui , 4:Add , Text , xm , Number of Pieces:
Gui , 4:Add , Edit , x+4 center yp-3 vcurrentNumberOfPieces number limit1 range1-9 gitemDistribution w30, %currentNumberOfPieces%
Gui , 4:Add , Checkbox , x+10 yp+3 vhasSupplement Checked%hasSupplement% , Has &Supplement 
Gui , 4:Add , Button , xm Default w110, OK	
Gui , 4:Add , Button , x+10 w110, Cancel	


myBubble("Welcome to BiblioScript by S. Jacob Powell!" , "BiblioScript DVD Project is now running")
sleep 2000
;Opening working excel spreadsheet
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
if (useExcel)
	ifNotExist , BiblioScriptDVDProject.xlsx
	{
		myBubble("BiblioScript Excel Spreadsheet","Opening Excel...")
		run excel
		winWait Excel
		winActivate , Excel
		myBubble("BiblioScript Excel Spreadsheet","Creating new excel spreadsheet...")
		Send ^n
		Sleep 200
		myBubble("BiblioScript Excel Spreadsheet","Saving spreadsheet...")
		Send ^s
		Sleep 200
		Send {alt}
		Sleep 200
		myBubble("BiblioScript Excel Spreadsheet","Opening ""Save As""...")
		Send a
		Sleep 200
		myBubble("BiblioScript Excel Spreadsheet","Opening ""Browse""...")
		Send b
		WinWaitActive , Save As
		Send !d
		Sleep 200
		TrayTip , BiblioScript Excel Spreadsheet , Navigating to the directory %A_WorkingDir%, , 20
		Send %A_WorkingDir%{enter}
		Sleep 200
		Send !n
		Sleep 200
		myBubble("BiblioScript Excel Spreadsheet","Entering the name ""BiblioScriptDVDProject""...")
		Send BiblioScriptDVDProject
		Sleep 200
		myBubble("BiblioScript Excel Spreadsheet","Saving new spreadsheet...")
		Send !s
		Sleep 700
		myBubble("BiblioScript Excel Spreadsheet","Minimizing Excel Spreadsheet...")
		minExcel()
	}
	else
	{
		ifWinNotExist , BiblioScriptDVDProject - Excel
			ifWinNotExist , BiblioScriptDVDProject.xlsx - Excel
			{
				myBubble("BiblioScript Excel Spreadsheet","Opening Excel Spreadsheet...")
				run BiblioScriptDVDProject.xlsx
				WinWait , BiblioScriptDVDProject
				myBubble("BiblioScript Excel Spreadsheet","Clearing Spreadsheet...")
				gosub clearExcelSpreadsheet
				Sleep 700
				myBubble("BiblioScript Excel Spreadsheet","Minimizing Excel Spreadsheet...")
				minExcel() 
			}
	}
myBubbleOff()

inProcess = 0
notShowingShortcuts = 1
;Main loop
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
loop
{
	if (!inProcess)
	{
		IfWinExist , BiblioScriptDVDProject - Excel
		{	
			WinGet , isMin , MinMax, BiblioScriptDVDProject - Excel
			if (isMin != -1)
				winMaximize , BiblioScriptDVDProject - Excel
		}	
		
		IfWinExist , BiblioScriptDVDProject.xlsx - Excel
		{	
			WinGet , isMin , MinMax, BiblioScriptDVDProject.xlsx - Excel
			if (isMin != -1)
				winMaximize , BiblioScriptDVDProject.xlsx - Excel
		}
		
		IfWinExist , SirsiDynix Symphony WorkFlows
		{	
			myBubbleOff()
			WinGet , isMin , MinMax, SirsiDynix Symphony WorkFlows
			if (isMin != -1)
				winMaximize , SirsiDynix Symphony WorkFlows
		}
		else
		{
			myBubble("Attention","Open WorkFlows and the ""Modify Title"" wizard.`n`nMay also work with ""Item Search and Display""`nwizard, but it won't make corrections.")
			Sleep 2000
		}
		
		IfWinNotExist , SureThing Disc Labeler
		{
			myBubble("Attention","Open ""SureThing DVD Labelling"" program, and that the template is titled ""MC DVD Labels"".")
			Sleep 2000
		}
		else
		{
			myBubbleOff()
			ifWinNotActive , ahk_class MVDIALOG
			{
				WinGet , isMin , MinMax, SureThing Disc Labeler
				if (isMin != -1)
					winMaximize , SureThing Disc Labeler
			}
		}
	}
}

;Main Process with {Enter}
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
#ifWinActive , SirsiDynix Symphony WorkFlows
numpadenter::
enter::
	inProcess = 1
	Send !i
	Sleep 200
	Send I
	Send !s
	Sleep 2000
	IfWinExist , Lookup
		return
	beforeChange()
	trueClick("CallTab","Click the Call Number/Item Tab")
	waitChange()
	recordCheck:	
	trueClick("CallNumberField","Click the Call Number Field")
	Sleep 200
	copyAll()
	currentCallNumber := clipboard
	IfInString , currentCallNumber , pt.
	{
		IfNotInString , currentCallNumber , |z
			checkCallNumber = 1
		else
			checkCallNumber = 0	
	}
	else
		checkCallNumber = 0
	StringReplace , currentCallNumber , currentCallNumber , |z , %A_Space% , 1
	if (setSchemeAndLibrary)
	{
		trueClick("ClassSchemeField","Click the Call Scheme Field")
		if (MCMode)
			Send LC
		else
			Send MUSIC DESK
		Sleep 200
		Send {ENTER}
		trueClick("ClassLibraryField","Click the Call Library Field")
		Send LEE-LRC
		Sleep 200
		Send {ENTER}
	}
	trueClick("NumberOfPiecesField","Click the Number of Pieces Field")
	copyAll()
	if (checkCallNumber)
		MsgBox , Make sure the "|z" is in the call number as appropriate, then click OK.
	StringLeft , currentNumberOfPieces , clipboard , 2
	goto 4GuiShow
	
recordCorrectionCheck:	
	trueClick("NumberOfPiecesField","Click the Number of Pieces Field")
	copyAll()
	StringLeft , currentNumberOfPieces , clipboard , 2
	goto 4GuiShow

afterPromptContinueCheck:	
	if (useExcel)
	{	
		maxExcel()
		Send {Esc}
		sleep 200	
		Send %currentCallNumber%
		Send {enter}
		if (hasSupplement)
		{
			Send %currentCallNumber%
			Send {enter}
		}	
		sleep 700
		minExcel()
	}
	discCount := discCount + currentDiscCount
	refreshVariables()
	sepParts = 0
	volume = 0
	/*
	if (currentDiscCount > 1)
	{
		MsgBox , 4356 , Disc Labels, Do you want the disc parts labeled with separate parts?`n`nExample: "Pt. 1" for Disc 1 and "Pt. 2" for Disc 2
		ifMsgBox yes
			sepParts = 1
		
		MsgBox , 4356 , Disc Labels, Do you want the disc parts labeled as volumes?`n`nExample: "vol. 1"`nPress "Yes" if the word "Volume" is used on the case.
		ifMsgBox yes
			volume = 1
		
	}
	*/
	if (useLabeler)
	{	
		partCount = 1
		loop , %currentDiscCount%
		{
			maxLabeler()
			checkCDClick = 0
			if (cdLabelFromFile = "")
				checkCDClick = 1
			forSureThing = 1
			trueClick("cdLabelFromFile","Click the CD label")
			Sleep 200
			forSureThing = 0
			if (checkCDClick)
				WinWait , Text Effect , , 1
			ifWinExist , Insert Field
				WinClose , Insert Field
			ifWinNotExist , Text Effect
			{
				Send !o
				Sleep 200
				Send ee{enter}
				WinWait , Text Effect , , 3
				ifWinExist , Insert Field
					WinClose , Insert Field
				ifWinNotExist , Text Effect
					if (checkCDClick)
					{
						loop
						{
							MsgBox , 4100 , Verify , Did the "Text Effect" window open?
							ifMsgBox , no
							{
								cdLabelFromFile = 
								refreshVariables()
								forSureThing = 1
								trueClick("cdLabelFromFile","Click the CD label")
								Sleep 200
								forSureThing = 0
								if (checkCDClick)
									WinWait , Text Effect , , 1
									ifWinExist , Insert Field
										WinClose , Insert Field
								ifWinNotExist , Text Effect
								{
									Send !o
									Sleep 200
									Send ee{enter}
									WinWait , Text Effect , , 3
									ifWinExist , Insert Field
										WinClose , Insert Field
								}
							}
							ifMsgBox , yes
								break
						}
					}
			}
			Send {tab 9}
			Sleep 200
			if (volume)
				StringReplace , currentCallNumber , currentCallNumber , pt. , vol. , A
			else
				StringReplace , currentCallNumber , currentCallNumber , pt. , pt. , A
			if ( ErrorLevel )
				replacedPtOrVol = 0
			else
				replacedPtOrVol = 1
				
			if (MCMode)
				Send HBLL MC`n`n%currentCallNumber%
			else	
				Send HBLL MLL`n`n%currentCallNumber%	
			/*
			if (replacedPtOrVol)
			{
				if (sepParts)
				{
					Send {Backspace}%partCount%
					partCount := partCount + 1	
				}
			}
			else
			{
				if (currentDiscCount > 1)
				{
					if (volume)
						Send {Space}vol.
					else
						Send {Space}pt.
						
					if (sepParts)
					{
						Send %partCount%
						partCount := partCount + 1	
					}
					else
					{
						myBubble("Volume Set","Enter the volume number, then press enter")
						WinActivate , Text Effect
						WinWaitActive , Text Effect
						if (volumeNumber = "")
						{
							Input , volumeNumber , V , {Enter}
							Send {backSpace}
						}
						else
							Send %volumeNumber%
						myBubbleOff()	
					}
				}
			}
			*/
			MsgBox , 4100 , Disc Labels, Continue?
			ifMsgBox no
				return
			WinActivate , Text Effect
			WinWaitActive , Text Effect
			Send !o
			WinWaitClose , Text Effect
			if (discCount <= 26)
				newDisc()
		}
		minLabeler()	
	}
	volumeNumber =
	currentDiscCount = 1
	refreshVariables()
	maxWorkFlows()
	Send !u
	if (discCount > 26)
	{
		discCount = 0
		if (useLabeler)
			goto printSequence
	}
	inProcess = 1
	endOfProcess:
	return
#ifWinActive

;Labels	
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------

discCountExceedTimer: 
	IfWinNotExist , Disc Count Exceeds Labels
		return
	SetTimer , discCountExceedTimer, off 
	WinActivate 
	ControlSetText , Button1 , &Continue
	ControlSetText , Button2 , &Different case 
	ControlSetText , Button3 , C&ancel
	return
	
printSequence:
	if (useExcel)
	{
		myBubble("Print the spine labels","After printing the spine labels, minimize excel to continue.")
		maxExcel()
		Send {Esc}
		sleep 200
		Send ^p
		Sleep 1000
		Send {tab 2}{down}
		Sleep 200
		WinWaitNotActive , ahk_class Net UI Tool Window
		ifWinExist , BiblioScriptDVDProject - Excel
			WinWaitNotActive ,  BiblioScriptDVDProject - Excel
		else
			WinWaitNotActive ,  BiblioScriptDVDProject.xlsx - Excel
		myBubbleOff()
	}
	if (useLabeler)
	{
		myBubble("Print the disc labels","After printing the disc labels, minimize SureThing Disc Labeler to continue.")
		maxLabeler()
		Send ^p
		WinWait , Print
		WinWaitClose , Print
		WinWaitNotActive , SureThing Disc Labeler
		myBubbleOff()
	}
	if (useLabeler || useExcel)
	{
		msgbox , 4100 , Waiting Input , Do you want to start a new set of DVDs?`n(This clears all of the stored values)
		ifMsgBox Yes
		{
			if (useExcel)
				gosub clearExcelSpreadsheet
			if (useLabeler)
			{		
				maxLabeler()
				Send !sd
				WinWaitActive , Delete Designs
				Send 25
				Sleep 200
				Send !o
				winWaitclose , Delete Designs
				minLabeler()
			}
		}
	}
	return
	
discRange:
	GuiControlGet , discCount , 3: , Edit1 
	GuiControl , 3:Text , Edit1 , %discCount%
	if (discCount = 0)
	{
		discCount = 1
		GuiControl , 3:Text , Edit1 , 1
	}
	ifWinActive , Properties
		Send {End}
	return
	
clearExcelSpreadsheet:
	maxExcel()
	Send {Esc}
	sleep 200
	Send ^a
	Send {Delete}
	Send ^{up}{left}{up}{left}
	Sleep 700
	minExcel()
	return	
	
3ButtonOK:
	Gui , 3:Submit
	refreshVariables()
	return
	
3ButtonClearMousePositions:
	MsgBox , 4100 , Warning , Are you sure you want to clear all of the saved mouse positions?
	IfMsgBox Yes
	{
		CallTab =
		CallNumberField =
		ClassSchemeField =
		ClassLibraryField =
		NumberOfPiecesField =
		refreshVariables()
		myBubble("Attention" , "All mouse positions have been cleared." , 3)
	}
	return
	
4ButtonOK:
		Gui , 4:Submit
		myBubbleOff()
		if (hasSupplement)
			currentDiscCount := currentNumberOfPieces - 1
		else
			currentDiscCount := currentNumberOfPieces
		refreshVariables()
		maxWorkFlows()
		if (discCount + currentDiscCount > 27)
		{
			SetTimer , discCountExceedTimer , 50
			MsgBox , 4099 , Disc Count Exceeds Labels , There are too many discs in this case.`nIf you choose to continue, it will try to print too many CD labels.`nIt can be fixed manually afterwards, or you can choose to do another case.`n`nPlease choose what you would like to do:
			ifMsgBox yes
			{
				goto afterPromptContinueCheck
			}
			ifMsgBox no
			{
				gosub 4ButtonCancel
				Send !u
			}
			ifMsgBox cancel
			{
				gosub 4ButtonCancel
			}
		}
		else
			goto afterPromptContinueCheck
	return	
	
4Corrections:
	Gui , 4:Hide
	goto makeCorrection
	
makeCorrection:
	trueClick("NumberOfPiecesField","Click the Number of Pieces Field")
	copyAll()
	Send %currentNumberOfPieces%
	goto 4GuiShow
	
4ButtonCancel:
	goto 4GuiClose
	
4GuiClose:
	Gui , 4:Cancel
	myBubbleOff()
	goto endOfProcess
	
4GuiShow:
	hasSupplement = 0
	Guicontrol , 4: , hasSupplement , %hasSupplement%
	myBubble("Shortcuts","Use numpad ""+""/""-"" for forward/backward`nUse numpad ""."" to toggle checkbox`nUse ""C"" for Cancel`nUse ""Enter"" for OK")
	GuiControl , 4:Text , Edit1 , %currentNumberOfPieces%
	Gui , 4:show ,  w250 , Number of Pieces
	ControlFocus , Edit1 , Number of Pieces
	Send ^a
	return	

itemDistribution:
	GuiControlGet , tempNumberOfPieces , 4:, Edit1
	promptToChange = 0	
	if (tempNumberOfPieces = "" || tempNumberOfPieces < 1)
	{
		currentNumberOfPieces = 1
		GuiControl , 4:Text , Edit1 , %currentNumberOfPieces%	
	}
	else
	{
		if (currentNumberOfPieces != tempNumberOfPieces)
			promptToChange = 1
	}
	ifWinActive , Number of Pieces
	{	
		Send ^a
	}
	if (promptTochange)
	{
		msgbox , 4100 , Number of Pieces , Change Number of Pieces to %tempNumberOfPieces%?
		ifMsgBox yes
		{
			currentNumberOfPieces := tempNumberOfPieces
			goto 4Corrections
		}
		ifMsgBox no
			GuiControl , 4:Text , Edit1 , %currentNumberOfPieces%
	}
	return
	
;Functions	
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
newDisc()
{
	Sleep 200
	Send {PgDn}
	WinWait , Create a New Design?
	Send !d
	Send !y
	WinWaitClose , Create a New Design?
	Sleep 200
}
deleteDisc()
{
	maxLabeler()
	Send !sd
	WinWaitActive , Delete Designs
	Send !o
	WinWaitClose , Delete Designs
}
copyAll()
{
	clipboard =
	loop
	{
		Send ^a
		Send ^c
		if (clipboard != "")	
			break
	}
}	
minExcel()
{
	ifWinExist , BiblioScriptDVDProject - Excel
		WinMinimize , BiblioScriptDVDProject - Excel
	else
		WinMinimize , BiblioScriptDVDProject.xlsx - Excel
	WinWaitNotActive , BiblioScriptDVDProject.xlsx - Excel
}
maxExcel()
{
	ifWinExist , BiblioScriptDVDProject - Excel
		WinActivate , BiblioScriptDVDProject - Excel
	else
		WinActivate , BiblioScriptDVDProject.xlsx - Excel
	WinWaitActive , BiblioScriptDVDProject
	WinMaximize , BiblioScriptDVDProject
}
minLabeler()
{
	WinMinimize , SureThing Disc Labeler 
	WinWaitNotActive , SureThing Disc Labeler
}
maxLabeler()
{
	WinActivate , SureThing Disc Labeler
	WinWaitActive , SureThing Disc Labeler
	WinMaximize , SureThing Disc Labeler
}
maxWorkFlows()
{
	WinActivate , SirsiDynix Symphony WorkFlows
	WinWaitActive , SirsiDynix Symphony WorkFlows
	WinMaximize , SirsiDynix Symphony WorkFlows
}
beforeChange(startXValue = 700, startYValue = 500)
{
	global 
	startX := startXValue
	startY := startYValue
	index := 1
	loop , 50
	{
		PixelGetColor , pixelLine%index% , %startXValue% + %index% , %startYValue% + %index%
		index += 1
	}
}
waitChange()
{
	global
	loop
	{
		waitForChangeX := startX
		waitForChangeY := startY
		index := 1
		loop , 50
		{
			PixelGetColor , currentPixel , %waitForChangeX% + %index% , %waitForChangeY% + %index%
			if (currentPixel = pixelLine%index%)
				break
			else
				index += 1
		}
		if (index < 50)
			break
	}
	Sleep 1250
}	
myBubbleOff()
{
	SetTimer , RefreshTrayTip, Off
	SetTimer , RefreshTrayTipForTrueClick, Off
	TrayTip
}		
myBubble(title = "Title" , text = "Text", seconds = 0 , type = 20)
{
	global
	myBubbleOff()
	#Persistent
	backupTitlesAndText = %title%:`n%text%`n`n
	bubbleTitle := title
	bubbleText := text
	bubbleSeconds := seconds
	bubbleType := type
	refreshVariables()
	
	if (!seconds)
	{
		SetTimer , RefreshTrayTip, 1000
		Gosub , RefreshTrayTip 
		return

		RefreshTrayTip:
		TrayTip , %bubbleTitle%, %bubbleText%, , %bubbleType%
		compoundBubble = 0
		return
	}
	else
	{
		TrayTip , %bubbleTitle%, %bubbleText%, , %bubbleType%
		bubbleTime := bubbleSeconds * 1000
		SetTimer , RemoveTrayTip, %bubbleTime%
		return

		RemoveTrayTip:
		SetTimer , RemoveTrayTip, Off
		TrayTip
		bubbleTitle =
		bubbleText = 
		bubbleSeconds = 
		bubbleType =
		return
	}
}
trueClick(variableFromFile , inputStringPar)
{
	global
	variableName := variableFromFile
	variableFromFile := %variableName%
	if (useMousePositions)
	{
		if (variableFromFile = "")
		{
			#Persistent
			SetTimer , RefreshTrayTipForTrueClick, 1000
			Gosub , RefreshTrayTipForTrueClick 
			if (!forSureThing)
				maxWorkFlows()
			else
			{
				ifWinNotExist , Text Effect
					maxLabeler()
				else
					winActivate , Text Effect
			}
			KeyWait , LButton , D
			MouseGetPos , variablex , variabley
			variableFromFile = %variablex% , %variabley%
			loop
			{
				GetKeyState , LButtonState , LButton
				if (LButtonState = "U")
					break
			}
			myBubbleOff()
		}
		else
		{
			Click , %variableFromFile%
			loop
			{
				GetKeyState , LButtonState , LButton
				if (LButtonState = "U")
					break
			}
		}	
	}
	else
	{
		#Persistent
		SetTimer, RefreshTrayTipForTrueClick, 1000
		Gosub , RefreshTrayTipForTrueClick 
		if (!forSureThing)
				maxWorkFlows()
			else
			{
				ifWinNotExist , Text Effect
					maxLabeler()
				else
					winActivate , Text Effect
			}
		KeyWait , LButton , D
		MouseGetPos , variablex , variabley
		variableFromFile = %variablex% , %variabley%
		loop
		{
			GetKeyState , LButtonState , LButton
			if (LButtonState = "U")
				break
		}
		myBubbleOff()
	}	
	%variableName% := variableFromFile
	refreshVariables()
	return
	
	RefreshTrayTipForTrueClick:
	TrayTip , Waiting..., %inputStringPar%, , 20
	return	
}
refreshVariables()
{
	saveVariables()
	updateGuis()
}


updateGuis()
{
	global
	GuiControl , 3:Text , Edit1 , %discCount%
	GuiControl , 4:Text , Edit1 , %currentNumberOfPieces%
}

loadVariables()
{
	global
	Critical , On
	helpToggle = 0
	IniRead , useMousePositions , BiblioScriptData.ini , variables , useMousePositions , 1
	IniRead , currentCallNumber , BiblioScriptData.ini , variables , currentCallNumber , %A_Space%
	IniRead , currentNumberOfPieces , BiblioScriptData.ini , variables , currentNumberOfPieces , 1
	IniRead , setSchemeAndLibrary , BiblioScriptData.ini , variables , setSchemeAndLibrary , 1
	IniRead , discCount , BiblioScriptData.ini , variables , discCount , 1
	IniRead , currentDiscCount , BiblioScriptData.ini , variables , currentDiscCount , 1
	IniRead , hasSupplement , BiblioScriptData.ini , variables , hasSupplement , 0
	IniRead , MLLMode , BiblioScriptData.ini , variables , MLLMode , 0
	IniRead , MCMode , BiblioScriptData.ini , variables , MCMode , 1
	IniRead , useExcel , BiblioScriptData.ini , variables , useExcel , 1
	IniRead , useLabeler , BiblioScriptData.ini , variables , useLabeler , 1
	
	
	StringLeft , useMousePositions , useMousePositions , 1
	StringLeft , currentNumberOfPieces  , currentNumberOfPieces , 3
	StringLeft , setSchemeAndLibrary  , setSchemeAndLibrary , 1
	StringLeft , discCount  , discCount , 3
	StringLeft , currentDiscCount  , currentDiscCount , 3
	StringLeft , MLLMode  , MLLMode , 1
	StringLeft , MCMode  , MCMode , 1
	if (MCMode)
		MLLMode = 0
	else
		MCMode = 1
		
	IniRead , CallTab , BiblioScriptData.ini , variables , CallTab , %A_Space%
	IniRead , CallNumberField , BiblioScriptData.ini , variables , CallNumberField , %A_Space%
	IniRead , ClassSchemeField , BiblioScriptData.ini , variables , ClassSchemeField , %A_Space%
	IniRead , ClassLibraryField , BiblioScriptData.ini , variables , ClassLibraryField , %A_Space%
	IniRead , NumberOfPiecesField , BiblioScriptData.ini , variables , NumberOfPiecesField , %A_Space%
	IniRead , cdLabelFromFile , BiblioScriptData.ini , variables , cdLabelFromFile , %A_Space%
	Critical , Off
	saveVariables()
}	

saveVariables()
{
	global
	Critical , On
	IniWrite , %useMousePositions% , BiblioScriptData.ini , variables , useMousePositions
	IniWrite , %currentCallNumber% , BiblioScriptData.ini , variables , currentCallNumber
	IniWrite , %currentNumberOfPieces% , BiblioScriptData.ini , variables , currentNumberOfPieces
	IniWrite , %setSchemeAndLibrary% , BiblioScriptData.ini , variables , setSchemeAndLibrary
	IniWrite , %discCount% , BiblioScriptData.ini , variables , discCount
	IniWrite , %currentDiscCount% , BiblioScriptData.ini , variables , currentDiscCount
	IniWrite , %hasSupplement% , BiblioScriptData.ini , variables , hasSupplement
	IniWrite , %MLLMode% , BiblioScriptData.ini , variables , MLLMode
	IniWrite , %MCMode% , BiblioScriptData.ini , variables , MCMode
	IniWrite , %useExcel% , BiblioScriptData.ini , variables , useExcel 
	IniWrite , %useLabeler% , BiblioScriptData.ini , variables , useLabeler
	IniWrite , %CallTab% , BiblioScriptData.ini , variables , CallTab
	IniWrite , %CallNumberField% , BiblioScriptData.ini , variables , CallNumberField
	IniWrite , %ClassSchemeField% , BiblioScriptData.ini , variables , ClassSchemeField
	IniWrite , %ClassLibraryField% , BiblioScriptData.ini , variables , ClassLibraryField
	IniWrite , %NumberOfPiecesField% , BiblioScriptData.ini , variables , NumberOfPiecesField
	IniWrite , %cdLabelFromFile% , BiblioScriptData.ini , variables , cdLabelFromFile
	Critical , Off
}	

;Hotkeys
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
F12::
	WinGetTitle , currentWindow
	if (helpToggle != 1)
	{
		Progress, zh0 B2  w550 fs18 C00 CT000000 CWccccc0, ,
(	
BiblioScript DVD Project Help Page (F12 to close)
---------------------------------------------------------------------

Pause Button 			| Pause Script
Ctrl + Shift + Pause Button	| Suspend Script
Ctrl + Shift + R			| Reload Script
Ctrl + Shift + P			| Properties
cc				| Type Call Number
)
		helpToggle = 1
	}
	else	
	{
		helpToggle = 0
		Progress , OFF
	}	
	WinActivate , currentWindow	
	return
^+p::
properties:
	Gui , 3:show , , Properties
	ControlFocus , Edit1 , Properties
	Send ^a
	return	
^+r::reload
^+pause::suspend	
pause::pause
:*:cc::
	Send %currentCallNumber%
	Return
#ifWinActive , Item Distribution and Record Check
	NumPadAdd::tab
	NumPadSub::
	send +{tab}
	return
#ifWinActive
#ifWinActive , Number of Pieces
	NumPadAdd::tab
	NumPadSub::
	send +{tab}
	return
	NumPadDot::
		if (hasSupplement || currentNumberOfPieces = 1)
		{
			hasSupplement = 0
			Guicontrol , 4: , hasSupplement , %hasSupplement%
		}
		else
		{
			hasSupplement = 1
			Guicontrol , 4: , hasSupplement , %hasSupplement%
		}
		return
	c::
		goto 4GuiClose
#ifWinActive

#t::
return	