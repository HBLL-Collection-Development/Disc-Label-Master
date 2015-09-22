#singleInstance ignore
#persistent
#noTrayIcon
CoordMode , Mouse , Relative
CoordMode , Pixel , Relative
SetTitleMatchMode, 2
SetKeyDelay , 20, 20

appName = Disc Label Master
workDirectory = %A_AppData%\%appName%
spreadSheetDirectory = %workDirectory%\%appName% Spine Labels.xlsx
iniFile = %workDirectory%\%appName%.ini
dvdLabelTemplateDirectory = %workDirectory%\MC DVD Labels.std
cdLabelTemplateDirectory = %workDirectory%\MLL Music Labels.std
dvdLabelProjectDirectory = %workDirectory%\%appName% SureThing Project MC.std
cdLabelProjectDirectory = %workDirectory%\%appName% SureThing Project MLL.std
splashDirectory = %workDirectory%\splash.png

columnWidth := 30.71
notShowingShortcuts = 1
loadingBlankSureThing = 0
howManyInARow = 0

FileCreateDir , %workDirectory%
SetWorkingDir , %workDirectory%
	if ( errorLevel = 1 )
		MsgBox , 0 , %appName% , Error Creating Workspace Directory

FileInstall, MC DVD Labels.std, %dvdLabelTemplateDirectory%
FileInstall, MLL Music Labels.std, %cdLabelTemplateDirectory%
FileInstall, MC DVD Labels.std, %dvdLabelProjectDirectory%
FileInstall, MLL Music Labels.std, %cdLabelProjectDirectory%
FileInstall, splash.png , %splashDirectory%

SplashImage , %splashDirectory% , B M

loopCount = 0
loop 500
{
	ifExist , %dvdLabelTemplateDirectory%
		ifExist , %cdLabelTemplateDirectory%
			break
	loopCount := A_Index		
}
if(loopCount > 495)
{
	MsgBox , 0 , %appName% , There was an error creating the SureThing projects files.`n`nPlease wait a few moments and try running %appName% again.
	SplashImage , Off
	goto exit
}
	

Menu , Tray, Tip, %appName%		
		
loadVariables()
cleanIniFile()

SysGet, mon, MonitorWorkArea , 1
guiHeight = 260
guiWidth = 350
guiY := monBottom - guiHeight - 45
guiX := monRight - guiWidth - 25

buttonWidth := guiWidth / 2 - 15 
messagesWidth := guiWidth - 15

Gui , main:+AlwaysOnTop -MaximizeBox
Gui , main:Add , Text , , Current Disc Number:
Gui , main:Add , Edit , x+5 yp-3 center vdiscCount gdiscRange number range1-99 limit2 w50
Gui , main:Add , UpDown ,  range1-99 , %discCount%
Gui , main:Add, Radio , xp+60 group vMCMode gMCModeG checked%MCMode%, HBLL MC (DVDs)
Gui , main:Add, Radio, vMLLMode checked%MLLMode% gMLLModeG , HBLL MLL (Music discs)
Gui , main:Add, GroupBox, xm w165 h135 section , Settings
Gui , main:Add , Checkbox , xp+5 yp+15  vsetSchemeAndLibrary gsetSchemeAndLibraryG Checked%setSchemeAndLibrary%, Set Scheme and &Library
Gui , main:Add , Checkbox , vuseMousePositions guseMousePositionsG Checked%useMousePositions%, Use Saved &Mouse Positions
Gui , main:Add , Checkbox , vuseLabeler guseLabelerG Checked%useLabeler%, Use &SureThing Labeler
Gui , main:Add , Checkbox , vuseExcel guseExcelG Checked%useExcel%, Use E&xcel Spreadsheet
Gui , main:Add , Checkbox , vonTop gonTopG Checked%onTop%, Always On &Top
Gui , main:Add , Checkbox , vmodifyOnly gmodifyOnlyG Checked%modifyOnly%, Work Only in Modify &Title

Gui , main:Add , Button , w%buttonWidth% ys+7 section gclearButtonG , &Clear Mouse Positions
Gui , main:Add , Button , xs w%buttonWidth% ghelpButtonG , &Help
Gui , main:Add , Button , xs  w%buttonWidth% gToggleExcelButtonG vToggleExcel, Show/Hide &Excel
Gui , main:Add , Button , xs w%buttonWidth% gblankTemplateButtonG vblankTemplateButton , &Reset SureThing and Excel
Gui , main:Add , Button , xs w%buttonWidth% gprintButtonG , &Print Labels

Gui, main:Add , Text, xm yp+12, Messages
Gui , main:Add , Edit , w%messagesWidth% r4 xm yp+15 vMessages ReadOnly -VScroll limit150

if(!useExcel)
	GuiControl, main:disable, ToggleExcel
	
if(!useLabeler)
	GuiControl, main:disable, blankTemplateButton

gui , main:show , h%guiHeight% w%guiWidth% x%guiX% y%guiY% , %appName%
; gosub properties

Gui , itemSpecifications:+AlwaysOnTop -MaximizeBox -MinimizeBox +ToolWindow
Gui , itemSpecifications:Add , Text , xm , Number of Pieces:
Gui , itemSpecifications:Add , Edit , x+4  yp-3  w40 center vcurrentNumberOfPieces gitemDistribution
Gui , itemSpecifications:Add , Updown ,  range1-9  , %currentNumberOfPieces%
Gui , itemSpecifications:Add , Checkbox , x+10 yp+3 vhasSupplement Checked%hasSupplement% ghasSupplementG , Has &Supplement 
Gui , itemSpecifications:Add , Button , xm Default vOKButtonItemSpecifications w110, OK	
Gui , itemSpecifications:Add , Button , x+10 vCancelButtonItemSpecifications w110, Cancel	
; gui , itemSpecifications:show

colors = 0x99CCFF

gui , main:color , %colors% , %colors%
gui , itemSpecifications:color , %colors% , %colors%


if(MCMode)
	SureThingProjectdirectory = %dvdLabelProjectDirectory%
else
	SureThingProjectdirectory = %cdLabelProjectDirectory%

if(useLabeler)
{
	if(WinExist("ahk_exe stdl.exe"))
	{
		MsgBox , 0 , %appName% , Please close any open SureThing Disc Labeler Windows to use %appName%.`n`nThe "Use SureThing Labeler" setting will be set to off.
		useLabeler = 0
		saveVariables()
		goto startExcel
	}

	gosub runLabeler
}

startExcel:
if(useExcel)
{
	if(WinExist("ahk_exe EXCEL.EXE"))
	{
		MsgBox , 0 , %appName% , Please close any open Excel Windows to use %appName%.`n`nThe "Use Excel Spreadsheet" setting will be set to off.
		useExcel = 0
		saveVariables()
		goto finishExcel
	}

	oExcel := ComObjCreate("Excel.Application")

	ifExist , %spreadSheetDirectory%
	{
		oSheet := oExcel.Workbooks.Open(spreadSheetDirectory)
	}
	else
	{
		oSheet := oExcel.Workbooks.Add
		oSheet.saveAs(spreadSheetDirectory)
	}
	oExcel.ActiveWorkbook.ActiveSheet.Columns("A").ColumnWidth := columnWidth
	oExcel.ActiveWorkbook.ActiveSheet.Columns("B").ColumnWidth := columnWidth
	oExcel.ActiveWorkbook.ActiveSheet.Columns("A").font.bold := "true"
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea := "$A:$A"
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.LeftMargin := 0
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.RightMargin := 0
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.TopMargin := 0
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.BottomMargin := 0
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.HeaderMargin := 0
	oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.FooterMargin := 0
	; oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea.FooterMargin := Application.InchesToPoints(0)

	oSheet.save()

	; oExcel.visible := 1
}

finishExcel:
	
SetTimer , timers , 200

SplashImage , Off

displayMessage("Welcome to " appName " by S. Jacob Powell!", 5)
loaded = loaded
return

runLabeler:
	try
		run %SureThingProjectdirectory%
	catch
		goto SureThingError
		
	winWait , ahk_exe stdl.exe , , 10
	
	if(errorLevel)
		goto SureThingError
		
	winMaximize , ahk_exe stdl.exe
	return

SureThingError:
	MsgBox , 0 , %appName% , The file extensions ".std" is not associated with SureThing Labeler`,`n%appName% requires SureThing Labeler to be installed correctly for use.`n`nIf this is unusual`, try restarting the program after a few moments.
	goto exit
	