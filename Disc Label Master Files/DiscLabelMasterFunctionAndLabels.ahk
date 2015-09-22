;Labels	
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------

timers:
	If((modifyOnly && winExist("Modify Title")) || (!modifyOnly && winExist("SirsiDynix Symphony WorkFlows")))
	{	
		WinGet , isMin , MinMax, SirsiDynix Symphony WorkFlows
		if (isMin != -1)
			winMaximize , SirsiDynix Symphony WorkFlows
	}
	else
		displayMessage("Open ""Modify Title"" in WorkFlows. " appName " should also work with ""Item Search and Display"", but it won't make corrections.", 5)
	
	if(winExist("SureThing Disc Labeler ahk_exe WerFault.exe","Close program"))
		goto myReload
	
	if(useExcel)
	{
		Process, Exist , EXCEL.EXE
		if(!errorLevel)
			goto exit
	}
	
	if(useLabeler && !winExist("ahk_exe stdl.exe") && !loadingBlankSureThing)
		goto exit
		
	return

discCountExceedTimer: 
	if(loaded != "loaded")
		return
	IfWinNotExist , Disc Count Exceeds Labels
		return
	SetTimer , discCountExceedTimer, off 
	WinActivate 
	ControlSetText , Button1 , &Continue
	ControlSetText , Button2 , &Different case 
	ControlSetText , Button3 , C&ancel
	return
	
;MainGui

printButtonG:
	if(loaded != "loaded")
		return
	printDiscLabels()
	printSpineLabels()
		
	msgbox , 4100 , %appName% , Do you want to start a new set of DVDs?`n(This clears all of the stored values)
	ifMsgBox No
		return
		
	gosub blankTemplateButtonG
	return

openBlankDiscProjectFromTemplate()
{
	global
	if(!useLabeler || loaded != "loaded")
		return
	loadingBlankSureThing = 1
	ControlSend , , ^s , ahk_exe stdl.exe
	winClose , ahk_exe stdl.exe
	winWaitClose , ahk_exe stdl.exe
	
	FileDelete, %SureThingProjectdirectory%
	
	if(MCMode)
		FileCopy, %dvdLabelTemplateDirectory% , %SureThingProjectdirectory%
	else		
		FileCopy, %cdLabelTemplateDirectory% , %SureThingProjectdirectory%
	
	gosub runLabeler
	
	
	discCount := 1

	GuiControl , main: , discCount , %discCount%
	saveVariables()
	howManyInARow = 0
	loadingBlankSureThing = 0
}

MLLModeG:
	if(loaded != "loaded")
		return
	GuiControlGet , TestIfSame, main: , MLLMode
	if(!winExist(appName " Restart") && TestIfSame != MLLMode) 
	{
		gui , main:+Disabled
		MsgBox , 4100 , %appName% Restart , %appName% will restart to apply changes.`n`nDo you want to continue?
		IfMsgBox Yes
		{
			GuiControlGet , MCMode , main:
			GuiControlGet , MLLMode , main:
			discCount := 1
			saveVariables()
			goto myReload
		}
		ifMsgBox No
		{
			Guicontrol , main: , MCMode , %MCMode%
			Guicontrol , main: , MLLMode , %MLLMode%
		}
	}
	
	gui , main:-Disabled
	return
	
MCModeG:
	if(loaded != "loaded")
		return
	GuiControlGet , TestIfSame, main: , MCMode
	if(!winExist(appName " Restart") && TestIfSame != MCMode) 
	{
		gui , main:+Disabled
		MsgBox , 4100 , %appName% Restart , %appName% will restart to apply changes.`n`nDo you want to continue?
		IfMsgBox Yes
		{
			GuiControlGet , MCMode , main:
			GuiControlGet , MLLMode , main:
			discCount := 1
			saveVariables()
			goto myReload
		}
		ifMsgBox No
		{
			Guicontrol , main: , MCMode , %MCMode%
			Guicontrol , main: , MLLMode , %MLLMode%
		}
	}
	gui , main:-Disabled
	return
	
helpButtonG:
	Gui , main:Hide
	MsgBox , 0 , %appName% Help,
(	
General:
This software helps automate the repetitive tasks associated with labling DVDs and CDs in the Lettering Department. The basic usage is to open "Modify Title" in Workflows, and to scan the barcode of the item (you can also type it in and press "Enter" to get the same effect). It checks and corrects certain parts of the record and then prompts you on how many parts are included (and whether there is a supplement ). You can use the prompt to make corrections to the number of parts. Once entered, the appropriate tasks and information will be sent to Excel and SureThing Disc Labeler in order to print when finished.

Current Disc Number - This should match the disc number in SureThing Labeler so it can detect whether you are at 26 labels.

HBLL MC/MLL - These modes determine which Disc Labeling Templates to use.
---------------------------------------------------------------------------------------
Settings:
Set Scheme and Library - This will make sure that each disc will be set toe the correct Class Scheme and Library in WorkFlows.

Use Saved Mouse Positions - After the first use, mouse clicks will be saved and used. Turn this off to click manually each time. 

Use SureThing Labeler - Determines whether to automate tasks in SureThing Labeler for Disc Labels.

Use Excel Spreadsheet - Determines whether to automate tasks in Excel for Spine Labels.

Always On Top - This keeps the main window on top of all other windows. (You can still minimize it if needed).

Work Only in Modify Title - This means that the automation will execute only when a "Modify Title" window is detected. Otherwise, it will execute as long as WorkFlows is detected. (Pressing "Enter" or scanning a barcode will execute the automation).
---------------------------------------------------------------------------------------
Buttons:
Clear Mouse Positions - This will reset all of the saved mouse positions, consequently requiring you to click manually on the next execution of the automation in WorkFlows.

Help - Opens this window.

Show/Hide Excel - Excel is kept hidden for ease of use. You can show or hide Excel at anytime by pressing this button.

Blank SureThing Template - This simply deletes the current open SureThing Project (make sure to print before doing this) and opens a blank template.

Print Labels - This assists in printing both Excel and SureThing, and then gives the option to start everything fresh. 
---------------------------------------------------------------------------------------
Other:
Certain prompts and messages are displayed in the "Messages" box. If the saved mouse positions are cleared, it will tell you where to click in WorkFlows. It will also tell you when to print and other bits of information.

You can also press "cc" to type the current call number.
---------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------
For support or questions feel free to contact me:

		S. Jacob Powell
		s.jacob.powell@gmail.com
		469-274-7045
)
	gui , main:show , h%guiHeight% w%guiWidth% x%guiX% y%guiY% , %appName%
	return
	
blankTemplateButtonG:
	if(loaded != "loaded")
		return
	MsgBox , 4100 , %appName% , This will delete your current SureThing disc Labeler project and the Spine Labels in Excel (columns A and B) and load blank windows.`n`nDo you want to continue?
	ifMsgBox Yes
	{
		openBlankDiscProjectFromTemplate()
		clearExcelSpreadsheet()
	}
	return

discRange:
	if(loaded != "loaded")
		return
	GuiControlGet , discCount , main:
	saveVariables()
	return
	
MainGuiClose:
	MsgBox , 4100 , %appName% , Are you sure you want to exit %appName%?
	IfMsgBox Yes
		goto exit
	return	
	
setSchemeAndLibraryG:
	if(loaded != "loaded")
		return	
	GuiControlGet , setSchemeAndLibrary , main:
	saveVariables()
	return

useMousePositionsG:
	if(loaded != "loaded")
		return
	GuiControlGet , useMousePositions , main:
	saveVariables()
	return

useLabelerG:
	if(!winExist(appName " Restart") && loaded = "loaded") 
	{
		gui , main:+Disabled
		MsgBox , 4100 , %appName% Restart ,%appName% will restart to apply changes.`n`nDo you want to continue?
		IfMsgBox Yes
		{
			GuiControlGet , useLabeler , main:
			saveVariables()
			goto myReload
		}
		ifMsgBox No
			Guicontrol , main: , useLabeler , %useLabeler%
	}
	gui , main:-Disabled
	return

useExcelG:
	if(!winExist(appName " Restart") && loaded = "loaded") 
	{
		gui , main:+Disabled
		MsgBox , 4100 , %appName% Restart , %appName% will restart to apply changes.`n`nDo you want to continue?
		IfMsgBox Yes
		{
			GuiControlGet , useExcel , main:
			saveVariables()
			goto myReload
		}
		ifMsgBox No
			Guicontrol , main: , useExcel , %useExcel%
	}
		gui , main:-Disabled
	return
modifyOnlyG:
	if(loaded != "loaded")
		return
	GuiControlGet , modifyOnly , main:
	saveVariables()
	return
	
onTopG:
	if(loaded != "loaded")
		return
	GuiControlGet , onTop , main:
	if(onTop)
		gui , main:+AlwaysOnTop
	else
		gui , main:-AlwaysOnTop
	return

ToggleExcelButtonG:
	if(loaded != "loaded")
		return
	oExcel.Visible := !oExcel.Visible
	if(oExcel.Visible)
		winMaximize , ahk_exe EXCEL.exe
	return
	
clearButtonG:
	if(loaded != "loaded")
		return
	MsgBox , 4100 , %appName% , Are you sure you want to clear all of the saved mouse positions?
	IfMsgBox Yes
	{
		CallTab =
		CallNumberField =
		ClassSchemeField =
		ClassLibraryField =
		NumberOfPiecesField =
		saveVariables()
		displayMessage("All mouse positions have been cleared." , 3)
	}
	return

;itemSpecificationsGui	
itemSpecificationsButtonOK:
	if(loaded != "loaded")
		return
	Gui , itemSpecifications:Submit
	gui , main:-Disabled
	if (hasSupplement)
		currentDiscCount := currentNumberOfPieces - 1
	else
		currentDiscCount := currentNumberOfPieces
	saveVariables()
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
			gosub itemSpecificationsButtonCancel
			ControlSend , , !u , SirsiDynix Symphony WorkFlows
		}
		ifMsgBox cancel
		{
			gosub itemSpecificationsButtonCancel
		}
	}
	else
		goto afterPromptContinueCheck
	clearMessage()
	return	
	
itemSpecificationsCorrections:
	if(loaded != "loaded")
		return
	Gui , itemSpecifications:Hide
	goto makeCorrection
	
makeCorrection:
	if(loaded != "loaded")
		return
	trueClick("NumberOfPiecesField","Click the Number of Pieces Field")
	copyAll()
	Send %currentNumberOfPieces%
	goto itemSpecificationsGuiShow
	
itemSpecificationsButtonCancel:
	if(loaded != "loaded")
		return
	gui , main:-Disabled
	clearMessage()
	goto itemSpecificationsGuiClose
	
itemSpecificationsGuiClose:
	if(loaded != "loaded")
		return
	Gui , itemSpecifications:Cancel
	gui , main:-Disabled
	clearMessage()
	goto endOfProcess
	
itemSpecificationsGuiShow:
	if(loaded != "loaded")
		return
	gui , main:+Disabled
	hasSupplement = 0
	Guicontrol , itemSpecifications: , hasSupplement , %hasSupplement%
	displayMessage("Use Up and Down to change number`nUse numpad ""."" to toggle checkbox`nUse ""C"" for Cancel`nUse ""Enter"" for OK")
	GuiControl , itemSpecifications: , currentNumberOfPieces , %currentNumberOfPieces%
	Gui , itemSpecifications:show ,  w250 , Number of Pieces
    GuiControl, itemSpecifications:+default, OKButtonItemSpecifications
    GuiControl, itemSpecifications:-default, CancelButtonItemSpecifications
	ControlFocus , itemSpecifications:currentNumberOfPieces , Number of Pieces
	Send ^a
	return	

itemDistribution:
	if(loaded != "loaded")
		return
	GuiControlGet , tempNumberOfPieces , itemSpecifications:, currentNumberOfPieces
	if (currentNumberOfPieces != tempNumberOfPieces)
	{
		msgbox , 4100 , Number of Pieces , Change Number of Pieces to %tempNumberOfPieces%?
		ifMsgBox yes
		{
			currentNumberOfPieces := tempNumberOfPieces
			goto itemSpecificationsCorrections
		}
		ifMsgBox no
			GuiControl , itemSpecifications: , currentNumberOfPieces , %currentNumberOfPieces%
	}
	return

hasSupplementG:
	if(loaded != "loaded")
		return
	GuiControlGet , currentNumberOfPieces , itemSpecifications:
	if(currentNumberOfPieces = 1)
		GuiControl , itemSpecifications: , hasSupplement , 0
	return
	
myReload:
	gosub myCloseStuff
	reload
	
myCloseStuff:
	SetTimer , timers , off
	gui , main:Hide
	winActivate , ahk_exe stdl.exe
	winClose , ahk_exe stdl.exe
	winWaitClose , ahk_exe stdl.exe
	
	try
	{
		oSheet.save()
		oExcel.Quit()
	}
	Process, Close , EXCEL.EXE
	sleep 500
	return
	
exit:
	gosub myCloseStuff
	exitApp		
	
;Hotkeys
:*:cc::
	Send %currentCallNumber%
	return

#ifWinActive , Item Distribution and Record Check
	NumPadAdd::tab
	NumPadSub::
	send +{tab}
	return
#ifWinActive
#ifWinActive , Number of Pieces
	NumPadDot::
		if (hasSupplement || currentNumberOfPieces = 1)
		{
			hasSupplement = 0
			Guicontrol , itemSpecifications: , hasSupplement , %hasSupplement%
		}
		else
		{
			hasSupplement = 1
			Guicontrol , itemSpecifications: , hasSupplement , %hasSupplement%
		}
		return
	c::
		goto itemSpecificationsGuiClose
#ifWinActive
	
;Functions	
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------
;-------------------------------------------------------------------------------------------------

printSpineLabels()
{
	global oExcel
	global useExcel
	if(!useExcel)
		return
	displayMessage("Print the desired spine labels in Excel, then click ""Show/Hide Excel"" to continue.")
	while(!oExcel.Visible)
	{
		gosub ToggleExcelButtonG
		sleep 1000
	}
	
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea := "$A:$A"
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.LeftMargin := 0
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.RightMargin := 0
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.TopMargin := 0
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.BottomMargin := 0
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.HeaderMargin := 0
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.FooterMargin := 0
	
	controlSend , , ^p , ahk_exe EXCEL.exe
	while(oExcel.Visible)
		sleep 200
	clearMessage()
}

printDiscLabels()
{
	global useLabeler
	if(!useLabeler)
		return
	ControlSend , , ^s , ahk_exe stdl.exe
	displayMessage("Print the desired disc labels in SureThing Disc Labeler.")
	while(!winExist("Print"))
	{
		PostMessage, 0x111, 448, 0, , ahk_exe stdl.exe
		sleep 2000
	}
	WinWaitClose , Print
	clearMessage()
}
newDisc()
{

	while(!WinExist("Add Designs"))
	{
		PostMessage, 0x111, 167, 0, , ahk_exe stdl.exe
		sleep 2000
	}
	while(WinExist("Add Designs"))
	{
		ControlSend , MVMINI1 , 1{enter}, Add Designs
		sleep 1500
	}
	ControlSend , , ^s , ahk_exe stdl.exe
}
deleteDisc(howMany = 1)
{
	
	while(!WinExist("Delete Designs"))
	{
		PostMessage, 0x111, 166, 0, , ahk_exe stdl.exe
		sleep 2000
	}
	while(WinExist("Delete Designs"))
	{
		ControlSend , MVMINI2 , %howMany%{enter}, Delete Designs
		sleep 1500
	}
	ControlSend , , ^s , ahk_exe stdl.exe
}
editDiscLabel(inputCallNumber)
{
	while(!WinExist("Text Effect"))
	{
		ControlSend , MVDECHILD1, ^a, ahk_exe stdl.exe
		PostMessage, 0x111, 548, 0, , ahk_exe stdl.exe
		sleep 2000
	}
	ControlSend , MVMINI2 , {tab 9}%inputCallNumber% , Text Effect
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

maxWorkFlows()
{
	WinActivate , SirsiDynix Symphony WorkFlows
	WinWaitActive , SirsiDynix Symphony WorkFlows
	WinMaximize , SirsiDynix Symphony WorkFlows
}

clearMessage()
{
	GuiControl , main: , messages , Ready.`n`nScan a barcode or press enter in WorkFlows (Modify Title)
}		

displayMessage(text = "Message", seconds = 0)
{
	GuiControlGet , messages , main:
	GuiControl , main: , messages , %text%
	if(messages != text)
		SetTimer , flashMainGui , -100
	if(seconds)
		Settimer , clearMessage , % -(seconds * 1000)
}
flashMainGui()
{
	loop, 2
	{
		Gui , main:Flash
		Sleep 300
	}
}

;Excel
clearExcelSpreadsheet()
{
	global
	if(!useExcel)
		return
	try
	{
		oExcel.ActiveWorkbook.ActiveSheet.Columns("A").Value := ""
		oExcel.ActiveWorkbook.ActiveSheet.Columns("B").Value := ""
	}
}
	
appendToExcel(callNumber, hasSupplement)
{
	global useExcel
	if(!useExcel)
		return
	global oExcel
	FormatTime, timeStamp , , MM/dd/yy h:mm tt
	
	curentEmpty := findNextEmptyCell()
	oExcel.Worksheets(1).Cells(curentEmpty, 1).Value := callNumber
	oExcel.Worksheets(1).Cells(curentEmpty, 2).Value := timeStamp
	if(hasSupplement)
	{
		curentEmpty := findNextEmptyCell()
		oExcel.Worksheets(1).Cells(curentEmpty, 1).Value := callNumber
		oExcel.Worksheets(1).Cells(curentEmpty, 2).Value := "Supplement Label (Same as above)"
	}
}

findNextEmptyCell()
{
	global useExcel
	if(!useExcel)
		return
	global oExcel
	i = 1
	while (oExcel.Worksheets(1).Cells(i, 1).Value)        
	{
		i++                                                
	}
	return i
}
	
;variables	
cleanIniFile()
{
	global 
	fileDelete , %iniFile%
	saveVariables()
}
loadVariables()
{
	global
	Critical , On
	IniRead , useMousePositions , %iniFile% , variables , useMousePositions , 1
	IniRead , currentCallNumber , %iniFile% , variables , currentCallNumber , %A_Space%
	IniRead , currentNumberOfPieces , %iniFile% , variables , currentNumberOfPieces , 1
	IniRead , setSchemeAndLibrary , %iniFile% , variables , setSchemeAndLibrary , 1
	IniRead , discCount , %iniFile% , variables , discCount , 1
	IniRead , currentDiscCount , %iniFile% , variables , currentDiscCount , 1
	IniRead , hasSupplement , %iniFile% , variables , hasSupplement , 0
	IniRead , MLLMode , %iniFile% , variables , MLLMode , 0
	IniRead , MCMode , %iniFile% , variables , MCMode , 1
	IniRead , useLabeler , %iniFile% , variables , useLabeler , 1
	IniRead , useExcel , %iniFile% , variables , useExcel , 1
	IniRead , onTop , %iniFile% , variables , onTop , 1
	IniRead , modifyOnly , %iniFile% , variables , modifyOnly , 1
	
	
	StringLeft , useMousePositions , useMousePositions , 1
	StringLeft , currentNumberOfPieces  , currentNumberOfPieces , 3
	StringLeft , setSchemeAndLibrary  , setSchemeAndLibrary , 1
	StringLeft , discCount  , discCount , 3
	StringLeft , currentDiscCount  , currentDiscCount , 3
	StringLeft , MLLMode  , MLLMode , 1
	StringLeft , MCMode  , MCMode , 1
	if (MCMode)
	{
		MCMode = 1
		MLLMode = 0
	}
	else
	{
		MCMode = 0
		MLLMode = 1
	}
		
	IniRead , CallTab , %iniFile% , variables , CallTab , %A_Space%
	IniRead , CallNumberField , %iniFile% , variables , CallNumberField , %A_Space%
	IniRead , ClassSchemeField , %iniFile% , variables , ClassSchemeField , %A_Space%
	IniRead , ClassLibraryField , %iniFile% , variables , ClassLibraryField , %A_Space%
	IniRead , NumberOfPiecesField , %iniFile% , variables , NumberOfPiecesField , %A_Space%
	Critical , Off
	saveVariables()
}	

saveVariables()
{
	global
	Critical , On
	IniWrite , %useMousePositions% , %iniFile% , variables , useMousePositions
	IniWrite , %currentCallNumber% , %iniFile% , variables , currentCallNumber
	IniWrite , %currentNumberOfPieces% , %iniFile% , variables , currentNumberOfPieces
	IniWrite , %setSchemeAndLibrary% , %iniFile% , variables , setSchemeAndLibrary
	IniWrite , %discCount% , %iniFile% , variables , discCount
	IniWrite , %currentDiscCount% , %iniFile% , variables , currentDiscCount
	IniWrite , %hasSupplement% , %iniFile% , variables , hasSupplement
	
	if (MCMode)
	{
		MCMode = 1
		MLLMode = 0
	}
	else
	{
		MCMode = 0
		MLLMode = 1
	}
	IniWrite , %MLLMode% , %iniFile% , variables , MLLMode
	IniWrite , %MCMode% , %iniFile% , variables , MCMode 
	IniWrite , %useLabeler% , %iniFile% , variables , useLabeler
	IniWrite , %useExcel% , %iniFile% , variables , useExcel
	IniWrite , %onTop% , %iniFile% , variables , onTop
	IniWrite , %modifyOnly% , %iniFile% , variables , modifyOnly
	IniWrite , %CallTab% , %iniFile% , variables , CallTab
	IniWrite , %CallNumberField% , %iniFile% , variables , CallNumberField
	IniWrite , %ClassSchemeField% , %iniFile% , variables , ClassSchemeField
	IniWrite , %ClassLibraryField% , %iniFile% , variables , ClassLibraryField
	IniWrite , %NumberOfPiecesField% , %iniFile% , variables , NumberOfPiecesField
	Critical , Off
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
	Sleep 300
}	

trueClick(variableFromFile , inputStringPar)
{
	global
	variableName := variableFromFile
	variableFromFile := %variableName%
	
	
	if (useMousePositions)
	{
		maxWorkFlows()
		if (variableFromFile = "")
		{
			displayMessage(inputStringPar)
			KeyWait , LButton , D
			MouseGetPos , variablex , variabley
			variableFromFile = %variablex% , %variabley%
			loop
			{
				GetKeyState , LButtonState , LButton
				if (LButtonState = "U")
					break
			}
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
		displayMessage(inputStringPar)
		maxWorkFlows()
		KeyWait , LButton , D
		MouseGetPos , variablex , variabley
		variableFromFile = %variablex% , %variabley%
		loop
		{
			GetKeyState , LButtonState , LButton
			if (LButtonState = "U")
				break
		}
	}	
	clearMessage()
	%variableName% := variableFromFile
	saveVariables()
	return
}
