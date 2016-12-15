; Jargon Counter - This script reads calculates the number & density of corporate jargons in a selected text
;                  To start the analysis, select any text & press CTRL+SHIFT+W
; Contact/Help  -  See Help menu in the application. 


#NoEnv                                                                             ; For performance & compatibility with future releases.
SetWorkingDir %A_ScriptDir%                                                        ; Ensures a consistent starting directory.
#SingleInstance force                                                              ; Make sure only one instance is running

; Read the list of Jargons from file
JargonMasterList = 
FileRead, Contents, MasterJargonList.txt
if not ErrorLevel                                                                   ; Successfully loaded.
{
   SplashTextOn,,, Starting Jargon Counter...
      
   Sort, Contents, U                                                                ; Sort & remove duplicates
   JargonMasterList = %Contents%
   
   StringReplace, JargonMasterList, JargonMasterList , `r`n,`, , UseErrorLevel       ; Replace NL with ,   
   totalJargonsDefined := ErrorLevel + 1                                             ; UseErrorLevel in previous command puts replace-counter in ErrorLevel
   StringLower,   JargonMasterList, JargonMasterList                                 ; Make everything in lower case   
   StringReplace, JargonMasterList, JargonMasterList , `-,`. , All                   ; Replace "-" with . so that we can work with regular expressions

   Contents =                                                                        ; Free the memory.
   Sleep, 800
   SplashTextOff
}
else
{
  MsgBox, 0x10,Jargon Counter, Aborting ! The Jargon File Master list file is missing.
  ExitApp  
}


TRAYMENU:
Menu,TRAY,NoStandard 
Menu,TRAY,DeleteAll 
Menu,TRAY,Add,&Settings,Settings
Menu,TRAY,Add
Menu,TRAY,Add,E&xit,EXIT
Menu,TRAY,Default,&Settings
Menu,Tray,Tip,Jargon Counter
Menu, Tray, Icon, resources/MisnStmtLogo.ico, , 1
return

EXIT:
ExitApp
return

ViewJargonFile:
Run notepad.exe %A_ScriptDir%\MasterJargonList.txt
return


SETTINGS:
Gui,1: Destroy
Gui,1: Add, Tab, x16 y10 w330 h160 , Instructions|Jargon List|About
Gui,1: Add, Button, x246 y180 w100 h30 gOverlay Default, OK
Gui,1: Tab, Instructions
Gui,1: Add, Text, x36 y50 w290 h45 , Select section of mail/document which you want to analyze and press CTRL+SHIFT+W
Gui,1: Add, Text, x36 y110 w290 h30, The tool will count the jargons and give detailed report of their (mis-)use
Gui,1: Add, Button, x246 y180 w100 h30 gOverlay Default, OK
Gui,1: Tab, Jargon List
;Gui,1: Add, Text,  x36 y50 w90 h15, Click to
Gui,1: Font, underline
Gui,1: Add, Text, cBlue gViewJargonFile, Click to view/edit the Jargon List
Gui,1: Font, norm
Gui,1: Add, Text,  x36 y60 w300 h50,`nYou can add more jargons to the file. Insert them 1 per line, just replace space/hypen with ".", save file and press reload button.
Gui,1: Add, Button, x36 y135 w250 h30 gReloadIt, Reload new Jargons
Gui,1: Add, Button, x246 y180 w100 h30 gOverlay Default, OK
Gui,1: Tab, About
Gui,1: Add, Text, x36 y50 w300 h30, Jargon Counter v1.8 by Arun Sivanandan 
Gui,1: Show, h218 w354, Jargon Counter
Gui,1: Add, Button, x246 y180 w100 h30 gOverlay Default, OK
Return

GuiClose:
Gui,1:Submit
Gui,1:Destroy
return

RELOADIT:
reload
return

OVERLAY:
Gui,1:Submit
Gui,1:Destroy
return

#MaxMem 256                                                           ; Allow 256 MB per variable.

^+w::
{  
  oCB := ClipboardAll                                                 ; Backup existing clipboard data
  Send, ^c   
  ClipWait 1
  
  strMesg = %clipboard%    
  resultMesg = 
  resultTotalJargonCount = 0  
  resultTotalJargonWordCount = 0
 
  SplashTextOn,,, Searching %totalJargonsDefined% Jargons 
   
  Loop, parse, JargonMasterList, `,
  {
      RegExReplace(strMesg, "i)"A_LoopField, "", thisJargonCount)                     ; i) makes it case insensitive, "" is replaced in string ?
      resultTotalJargonCount := resultTotalJargonCount + thisJargonCount
      
      if thisJargonCount > 0 
      {
           StringReplace, thisJargonFormatted, A_LoopField, . ,%A_Space%, UseErrorLevel
           thisJargonWordCount :=  ErrorLevel + 1
           thisJargonWordUsed :=   thisJargonWordCount*thisJargonCount                
           resultTotalJargonWordCount := resultTotalJargonWordCount + thisJargonWordCount*thisJargonCount
           ;resultMesg = %resultMesg% `n "%thisJargonFormatted%"(%thisJargonCount%x%thisJargonWordCount%=%thisJargonWordUsed%)                       
           resultMesg = %resultMesg% `n %thisJargonFormatted% (%thisJargonCount%)
      }     
  }
 
  SetFormat, Float, 3.2
  resultJargonDensity = 0
  RegExReplace( strMesg, "\w+", "", wordCount )   
  if ( wordCount > 0)
    resultJargonDensity :=  resultTotalJargonWordCount / wordCount * 100
  else
    resultJargonDensity := 0

  Sleep 1000
  SplashTextOff  

  MsgBox,0x30, Jargon Counter, %resultTotalJargonWordCount% words in the %wordCount% word text were used for jargons, that is ( %resultJargonDensity%`% ) ! `n `n %resultMesg%
    
  ;Restore old Clipboard data
  ClipBoard := oCB ; restore ClipBoard
  Return                     
}                            

; Test str: boil the ocean burning platform  burning platform buy in collaboration competitive  bandwidth band ; width