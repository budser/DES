; Script assembled by 
; Aron Talenfeld
; p 201.207.0753 | e budser@gmail.com

; 1)Download a portable copy (no installation required) of AHK from https://autohotkey.com/download/
;   Follow the link "Download AutoHotkey .zip" and extract "AutoHotkeyU64.exe" to a safe place on your hard drive.
; 2)Save this file (the one you're reading) in .ahk format to a safe place on your hard drive. Edit it in NotePad
;   or the text editor of your choice. For example, you can replace "[DEFAULT]" in the script below with your 
;   location specific information.
; 3)Select Autohotkey.exe as the default program to open this .ahk file
; 4)Create a shortcut to this file in the startup folder 
;   (i.e. C:\Users\[yourAlias]\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup)

; ################ Summary of hotkeys/hotstrings at bottom of this file #################

!F5::Reload ; Hit Alt + F5 to reload this AHK script (will also stop a running hotkey)

^#v::                            ; Text–only paste from ClipBoard!!!
   Clip0 = %ClipBoardAll%
   ClipBoard = %ClipBoard%       ; Convert to text
   Send ^v                       ; For best compatibility: SendPlay
   Sleep 50                      ; Don't change clipboard while it is pasted! (Sleep > 0)
   ClipBoard = %Clip0%           ; Restore original ClipBoard
   VarSetCapacity(Clip0, 0)      ; Free memory
Return

::-es::		; Edit the AutoHotkey script
Run C:\Users\atalenfeld\Documents\My Program Files\NotePad++\notepad++.exe C:\Users\atalenfeld\Documents\My Program Files\AutoHotkey.ahk
Return

::-ss::		; save and reload script but only if Notepad or Notepad++ is open and has focus
SetTitleMatchMode, 2
IfWinActive, Notepad ; any window with Notepad in title
{
	SendInput ^s
;	Sleep, 250
;	WinClose, Notepad
	Sleep, 250
	Reload, C:\Users\atalenfeld\Documents\My Program Files\AutoHotkey.ahk
}
Return

::-sleep::
Send Sleep, `%SleepTime`%{Enter}{Enter} 
Return 


::-notes::		; Automatically fills the delivery notes section with the most important checklist items
SendInput ^f 
Sleep, 500
SendInput Orientation Notes
Sleep, 500
SendInput {Escape}
Sleep, 500
SendInput {Shift Down}{Tab}{Shift Up}
Sleep, 500
SendInput {Enter}
Sleep, 500
SendInput {Ctrl Down}{Home}{Ctrl Up} 
Sleep, 500
SendInput --------------------------------------------{Enter}Delidate: {Enter}Payment: {Enter}Registration: {Enter}Home Charging: {Enter}Trade-In: {Enter}DL: {Enter}Insurance: {Enter}Plates:  {Enter}--------------------------------------------{Enter}{Enter} 
Return


::-price::		; scrapes price information to a notepad file for a pre-pre-MVPA price breakdown to help customers determine the amount they need to finance 

SleepTime := 150 ; set time to sleep between steps using browser

Price := ""
Discount := ""
TradeValue := ""
NoTrade := "0.00"
Referral := ""
ShowroomDiscount := ""
HasReferral := "Referral Credit" ; if this string is found in "Other Item 2," the customer has a referral
Paid := ""
ClipBoard = "" ; empty clipboard

Send ^f		; get vehicle price
Sleep, %SleepTime%
Send Vehicle Price for MVPA	
Sleep, %SleepTime%
Send {Escape}
Sleep, %SleepTime%
Send {Shift Down}{Ctrl Down}{Right}{Right}{Ctrl Up}{Left}{Shift Up} 
Sleep, %SleepTime%
Send ^c 
Sleep, %SleepTime%
Price := SubStr(ClipBoard, 29) ; trim search term off the clipboard to leave Vehicle Price only 
Sleep, %SleepTime% 

Send ^f		; get trade-in equity
Sleep, %SleepTime%
Send Other Downpayment
Sleep, %SleepTime%
Send {Escape}
Sleep, %SleepTime%
Send {Shift Down}{Up}{Right}{Right}{Right}{Right}{Shift Up}
Sleep, %SleepTime%
Send, ^c
Sleep, %SleepTime%
StringTrimRight, TradeValue, ClipBoard, 2

Send ^f		; see if there is a referral credit
Sleep, %SleepTime%
Send  Test Drive Scheduled Notes
Sleep, %SleepTime%
Send {Escape}
Sleep, %SleepTime%
Send {Shift Down}{Up}{Shift Up}
Sleep, %SleepTime%
Send, ^c
Sleep, %SleepTime%

IfInString, ClipBoard, %HasReferral% ; if "Referral Credit" found in section "Other Item 2," customer will be owed referral credit
{
	Send ^f		
	Sleep, %SleepTime%
	Send  referred by
	Sleep, %SleepTime%
	Send {Escape}
	Sleep, %SleepTime%
	Send {Shift Down}{Up}{Up}{Up}{Shift Up}
	Sleep, %SleepTime%
	Send, ^c
	StringTrimRight, Referral, ClipBoard, 6
	StringTrimLeft, Referral, Referral, 4 
	HasReferral := "Yes"
}

Send ^f		; see if there is a showroom discount
Sleep, %SleepTime%
Send  Vehicle Price for MVPA
Sleep, %SleepTime%
Send {Escape}
Sleep, %SleepTime%
Send {Shift Down}{Up}{Up}{Shift Up}
Sleep, %SleepTime%
Send, ^c
Sleep, %SleepTime%

IfInString, ClipBoard, Showroom Discount ; if "Showroom Discount" found in section "Other Item 1," customer will be owed discount
{
	Send ^f		
	Sleep, %SleepTime%
	Send  Vehicle Price for Configuration
	Sleep, %SleepTime%
	Send {Escape}
	Sleep, %SleepTime%
	Send {Shift Down}{Up}{Up}{Shift Up}
	Sleep, %SleepTime%
	Send, ^c
	StringTrimRight, ShowroomDiscount, ClipBoard, 2
	StringTrimLeft, ShowroomDiscount, ShowroomDiscount, 7 
	HasShowroomDiscount := "Yes"
}

Send ^f		; Get amount already paid
Sleep, %SleepTime%
Send  Sum of Payments	
Sleep, %SleepTime%
Send {Escape}
Sleep, %SleepTime%
Send {Shift Down}{Down}{Down}{Left}{Left}{Shift Up}
Sleep, %SleepTime%
Send, ^c
Sleep, %SleepTime%
StringTrimLeft, Paid, ClipBoard, 22 

Run Notepad
Sleep, 1500

SetTitleMatchMode, 2
IfWinExist, Notepad
{
	WinActivate
	Sleep, 500	
}
Else
{
MsgBox, Please open Notepad or a plain text editor to continue
}

SendInput %Price% (sticker price){Enter} 

IfNotInString, TradeValue, %NoTrade% ; if "0.00" is not in the trade value, type trade value
{
	IfInString, TradeValue, - ; If negative equity on trade, this value needs to be added to the cost of the car
	{
		StringTrimLeft, TradeValue, TradeValue, 1 
		SendRaw +
		SendInput %TradeValue% (Negative Trade-in equity){Enter}
	}
	Else
	{
		SendInput -%TradeValue% (Trade-in equity){Enter} ; Positive equity on trade needs to be subtracted from purchase price
	}
} 

Sleep, %SleepTime%
SendRaw +
SendInput 451 [title, registration, and processing fee for financed vehicle]
Sleep, %SleepTime%
SendInput {Enter}
Sleep, %SleepTime%

If(HasReferral = "Yes")
{
	SendInput %Referral% [Referral Credit]{Enter}
}

Sleep, %SleepTime%

If (HasShowroomDiscount = "Yes") 
{
	SendInput %ShowroomDiscount% [Showroom Discount]{Enter}
}

SendInput -%Paid% [Already paid toward purchase]{Enter}-[financed amount]{Enter}=[balance due]

Return

::-ins::		; scrapes insurance information from a customer's delivery object into a notepad file 

SleepTime := 150 ; set time to sleep between steps using browser

VIN := 0
Model := ""
Price := 0
Discount := ""
DeliDate := ""
ModelYear := ""
LienHolder := ""
LHSL := 0		; LienHolder String Length
OrderType := ""

Send ^f
Sleep, %SleepTime%
Send Lien-holder Address1
Sleep, %SleepTime%
Send {Escape}
Sleep, %SleepTime%
Send {Shift Down}{Up}{Up}{Shift Up} 
Sleep, %SleepTime%
Send ^c
Sleep, %SleepTime%
StringTrimRight, OrderType, ClipBoard, 2
StringTrimLeft, OrderType, OrderType, 3
;MsgBox, OrderType is _%OrderType%_

SendInput ^f
Sleep, %SleepTime%
SendInput Final Paperwork Planning ; Unique term near Lien-Holder Name1 field
Sleep, %SleepTime% 
SendInput ^f
Sleep, %SleepTime%
SendInput Order Type ; Unique term near Lien-Holder Name1 field
Sleep, %SleepTime% 
SendInput {Escape} 
SendInput {Shift Down}{Up}{Shift Up}
Sleep, %SleepTime%
Send ^c
Sleep, %SleepTime% 
StringTrimRight, LienHolder, Clipboard, 2   
Sleep, %SleepTime%
;MsgBox, LienHolder is _%LienHolder%_ 

StringLen, LHSL, LienHolder	; Set LHSL to the length of LienHolder

if (OrderType = "Cash")
{
	if(LHSL > 6)
	{
		MsgBox, Error: Order Type is "Cash" but there appears to be a lien-holder! Hit Alt + F5 to terminate hotkey or "OK" to continue.
	} 
	LienHolder := "Purchased Outright (No Lien-Holder)" ; No Lien-holder
} 
else if(OrderType = "Not Confirmed")
{ 
	LienHolder := "" ; Lien-holder probably not known yet
}
else if(LHSL < 6)	; LienHolder name is probably not correct since it's less than 6 characters
{
	MsgBox, Error: Lien-Holder name is a bit short.  Alt + F5 to terminate hotkey or "OK" to continue.
}
else if(LienHolder = "Tesla Lease Trust")
{
	LienHolder := "Lien-holder address for Tesla Leasing: Tesla Motors, 3500 Deer Creek Road, Palo Alto, CA 94304"
}
else if((LienHolder = "Alliant Credit Union") or (LienHolder = "Alliant C.U.") or (LienHolder = "Alliant CU")) 
{
	LienHolder := "Lien-holder address for financing: Alliant Credit Union, PO BOX 924350, Fort Worth, TX 76124"
}
else if(LienHolder = "Pentagon Federal Credit Union") 
{
	LienHolder := "Lien-holder address for financing: Pentagon Federal Credit Union, PO Box 255483, Sacramento CA 95865"
}
else if((LienHolder = "TECH CU") or (LienHolder = "Technology Credit Union") or (LienHolder = "Tech C.U."))
{ 
	LienHolder := "Lien-holder address for financing: TECH CU, PO BX 1409, SAN JOSE, CA 95109"
}
else if((LienHolder = "TD Auto Finance") or (LienHolder = "TD Bank") or (LienHolder = "T.D. Bank") or (LienHolder = "TD Auto Finance LLC")or (LienHolder = "TD Auto Finance L.L.C.")) 
{
	LienHolder := "Lien-holder address for financing: TD Auto Finance P.O. Box 9224 Farmington Hills, MI 48333-9224"
}
else if((LienHolder = "Chase Auto Finance") or (LienHolder = "Chase Bank") or (LienHolder ="JP Morgan Chase Bank")) 
{ 
	LienHolder := "Lien-holder address for financing: Chase Auto Finance, PO Box 80015, Atlanta, GA 30366-0015"
}
else if(LienHolder = "USB Leasing LT")
{
	LienHolder := "Lien-holder address for Tesla Leasing with US Bank: USB Leasing LT, P.O. Box 200029-9246, Kennesaw GA, 30156-9246"
}
else if(OrderType = "Third Party Lending")
{
	LienHolder := "Lien-Holder for Third Party Financing: [TBD]" ; DES can insert 3rd party finance address if known (FYI: Alliant, TCU, and PenFed are 3rd party but this else clause will not be reached if they're selected) 
}
else 
{
	MsgBox, Error: Lien-holder not recognized 
	Return
}

SendInput ^f
Sleep, %SleepTime%
SendInput Vehicle Information ; Unique term near VIN field 
Sleep, %SleepTime% 
SendInput {Escape} 
Sleep, %SleepTime% 
SendInput {Shift Down}{Down}{Down}{Down}{Left}{Shift Up}
Sleep, %SleepTime%
Send ^c
Sleep, %SleepTime% 
StringTrimLeft, VIN, Clipboard, 27   
;MsgBox, VIN is _%VIN%_

Send ^f
Sleep, %SleepTime%
Send WD: Containment Hold ; Unique term near Scheduled Delivery Date field 
Sleep, %SleepTime% 
Send {Escape} 
Sleep, %SleepTime% 
Send {Shift Down}{Up}{Shift Up} 
Sleep, %SleepTime% 
Send ^c 
Sleep, %SleepTime% 
StringTrimRight, DeliDate, Clipboard, 2
;MsgBox, Delidate is _%DeliDate%_

Send ^f
Sleep, %SleepTime%
Send Country of Use ; Unique term near Model Year 
Sleep, %SleepTime% 
Send {Escape} 
Sleep, %SleepTime% 
Send {Shift Down}{Up}{Shift Up} 
Sleep, %SleepTime% 
Send ^c 
Sleep, %SleepTime% 
StringTrimRight, ModelYear, Clipboard, 2
;MsgBox, Model Year is _%ModelYear%_
Send ^f
Sleep, %SleepTime%
Send Model-Series ; Unique term near Model Trim
Sleep, %SleepTime% 
Send {Escape}
Sleep, %SleepTime%
Send {Shift Down}{Up}{Up}{Shift Up} 
Sleep, %SleepTime% 
Send ^c 
Sleep, %SleepTime% 
StringTrimRight, Model, Clipboard, 2 ; trim extra characters from right 
Sleep, %SleepTime%
StringTrimLeft, Model, Model, 3 ; trim extra characters from left
Sleep, %SleepTime%
;MsgBox, Model is _%Model%_ 

Send ^f
Sleep, %SleepTime%
Send Vehicle Price for MVPA	
Sleep, %SleepTime%
Send {Escape}
Sleep, %SleepTime%
Send {Shift Down}{Ctrl Down}{Right}{Right}{Ctrl Up}{Left}{Shift Up} 
Sleep, %SleepTime%
Send ^c 
Sleep, %SleepTime%
Price := SubStr(ClipBoard, 29) ; trim search term off the clipboard to leave Vehicle Price only 
Sleep, %SleepTime% 
;MsgBox, Price is _%Price%_

; see if there is a discount
Send ^f
Sleep, %SleepTime%
Send Showroom Discount
Sleep, %SleepTime%
Send {Escape}
Sleep, %SleepTime%
Send ^c
Sleep, %SleepTime%
If (Clipboard = "Showroom Discount") ; There is a discount
{
	Send ^f
	Sleep, %SleepTime%
	Send Other Item 1 Amount
	Sleep, %SleepTime%
	Send {Escape}
	Sleep, %SleepTime%
	Send {Shift Down}{Ctrl Down}{Right}{Right}{Ctrl Up}{Left}{Shift Up} 
	Sleep, %SleepTime% 
	Send ^c 
	Sleep, %SleepTime% 
	Discount := SubStr(ClipBoard, 28) ; trim search term off the clipboard to leave "Showroom Discount" only 
	Sleep, %SleepTime%
}
;MsgBox, Discount is _%Discount%_
;---------------[Insert data into a blank notepad file]-------------------

Run Notepad
Sleep, 1500

SetTitleMatchMode, 2
IfWinExist, Notepad
{
	WinActivate
	Sleep, 500	
}
Else
{
MsgBox, Please open Notepad or a plain text editor to continue
}
Sleep, %SleepTime%

Send Make, Model, and Trim: Tesla %Model%{Enter}Year: %ModelYear%{Enter}VIN: %VIN%{Enter}Effective Date: %DeliDate%{Enter}Vehicle Price: %Price%
Sleep, %SleepTime%

If (Discount != "") 
{
	Send {Space}(Excluding the showroom discount of $%Discount%)
} 
Send {Enter}

Sleep, %SleepTime%
Send %LienHolder%{Enter}

Return

::-lha::	; Type the most common lien-holder addresses for insurance agents
SendInput Lien-holder address for financing: Alliant Credit Union, PO BOX 924350, Fort Worth, TX 76124{Enter}Lien-holder address for financing: TECH CU, PO BX 1409, SAN JOSE, CA 95109{Enter}Lien-holder address for financing: TD Auto Finance P.O. Box 9224 Farmington Hills, MI 48333-9224{Enter}Lien-holder address for financing: Chase Auto Finance, PO Box 80015, Atlanta, GA 30366-0015{Enter}Lien-holder address for Tesla Leasing: Tesla Motors, 3500 Deer Creek Road, Palo Alto, CA 94304{Enter}Lien-holder address for Tesla Leasing with US Bank: USB Leasing LT, P.O. Box 200029-9246, Kennesaw GA, 30156-9246
Return

::-tss::	; Open the printable TSS calendar 
Run https://tss.teslamotors.com/en-US/Print
Return 

::-psc:: 	 ; Type main contacts for paramus service center emails
SendInput Jose Garcia <jgarcia2@tesla.com>; Attalah Hadad <ahadad@tesla.com>; Jeny Vargas <jenvargas@tesla.com>; Karla Zaldivar <KZaldivar@tesla.com>; Jordan Reid <joreid@tesla.com>; Maria Recupero <MRecupero@tesla.com>; Leart Krasniqi <lkrasniqi@tesla.com>; Anthony Roy de Leon <aroydeleon@tesla.com>; John Raney <jraney@tesla.com>; Eric Brimat <ebrimat@tesla.com>; Christopher Dyne <cdyne@tesla.com>; Roberto Santoro <rsantoro@tesla.com>; David Han <dhan@tesla.com>
Return 

::-ws:: 	; Download a vehicle's window sticker from the delivery object page in Salesforce.com

SleepTime := 250 ; set time to sleep between steps 
SleepTime2 := 4000 ; set time to sleep between steps

SendInput ^f
Sleep, %SleepTime%
SendInput Delivery Record Overview ; Unique term near Vehicle object link
Sleep, %SleepTime%
SendInput {Escape}
Sleep, %SleepTime%
SendInput {Tab}{Tab}{Enter}

Sleep, %SleepTime2%

SendInput ^f
Sleep, %SleepTime%
SendInput Monroney_
Sleep, %SleepTime%
SendInput {Escape}
Sleep, %SleepTime%
SendInput {Enter}
Sleep, 2000
SendInput ^f
Sleep, %SleepTime%
SendInput view file
Sleep, %SleepTime%
SendInput {Escape}
Sleep, %SleepTime%
SendInput {Enter}
Return 



::-dpfdc::	;Set delivery probability in the delivery object to "High" and indicate that the FDC email has been sent

SleepTime := 500 ; set time to sleep between steps using browser

SendInput ^f 
Sleep, %SleepTime%
SendInput Delivery Planning{Enter} 
Sleep, %SleepTime%
SendInput {Escape}
Sleep, %SleepTime%
SendInput {Tab}{Enter}

Sleep, %SleepTime%
Sleep, %SleepTime%
Sleep, %SleepTime%

SendInput ^f 
Sleep, %SleepTime%
SendInput Dependent Fields
Sleep, %SleepTime%

SendInput {Enter} 
Sleep, %SleepTime%
SendInput {Escape}
Sleep, %SleepTime%
SendInput {Tab}{Up}{Up}{Up}{Up}{Up}
Sleep, %SleepTime%
SendInput      {Up}{Up}{Up}{Up}{Up}
Sleep, %SleepTime%
SendInput {Down} 
Sleep, %SleepTime%
SendInput {Tab}{Up}{Up}{Up}{Up}{Up}{Up}
Sleep, %SleepTime%
SendInput      {Up}{Up}{Up}{Up}{Up}{Up}{Down}{Down} 
Sleep, %SleepTime%
SendInput {Tab}{Enter}
Sleep, %SleepTime%

save() ; Click the Save button at the top of the Delivery Object

Return

::-dpi::		;Set delivery probability in the delivery object to "High" and indicate that the intro email has been sent 

SleepTime := 500 ; set time to sleep between steps using browser

SendInput ^f 
Sleep, %SleepTime%
SendInput Delivery Planning{Enter} 
Sleep, %SleepTime%
SendInput {Escape}
Sleep, %SleepTime%
SendInput {Tab}{Enter}
Sleep, %SleepTime%
Sleep, %SleepTime%
Sleep, %SleepTime% 

SendInput ^f 
Sleep, %SleepTime%
SendInput Dependent Fields{Enter}
Sleep, %SleepTime%
SendInput {Escape}
Sleep, %SleepTime%
SendInput {Tab}{Up}{Up}{Up}{Up}{Up}
Sleep, %SleepTime%
SendInput      {Up}{Up}{Up}{Up}{Up}
Sleep, %SleepTime%
SendInput {Down} 
Sleep, %SleepTime%
SendInput {Tab}{Up}{Up}{Up}{Up}{Up}{Up}
Sleep, %SleepTime%
SendInput      {Up}{Up}{Up}{Up}{Up}{Up}{Down}
Sleep, %SleepTime%
SendInput {Tab}{Enter}
Sleep, %SleepTime%

save() ; Click the Save button at the top of the Delivery Object

Return 
	
::-trade::		; Type self-inspection WuFoo link
Send https://teslafactory.wufoo.com/forms/ne-ny-metro-self-tradein-inspection-form/
Return
	
::-range::		; Type latest range figures (as of 6/2016)
Send EPA Rated Range as of June 2016:{Enter}{Enter} Model S: 60kWh - 210 mi, 60D - 218 mi, 75kWh - 249 mi, 75D - 259 mi, 90D 294 mi{Enter}{Enter}Model X: 75D - 237 mi, 90D - 257 mi, P90D - 250 mi
Return 

::-c:: ; Type "-c" in Chrome to go to the the cached version of almost any website.  This allows you to bypass advertisements, adblocker warnings, paywalls, and any number of other extraneous stuff (like javascript) that websites throw at you.
SleepTime := 250 ; Set the time to sleep between steps to .25 seconds
SendInput {Esc} ; In case "-c" hotstring typed in url bar, this restores the current URL in chrome
Sleep, %SleepTime%
SendInput ^l ; put focus on URL bar
Sleep, %SleepTime%
SendInput {Home} ; move to left of URL
Sleep, %SleepTime%
SendInput cache://{Enter} ; tell Chrome to show the cached version of page 
Return 


::-d:: ; opens dictionary.com word suggestions or definition for the last word typed or any text that might be highlighted

SleepTime := 250 ; Set the time to sleep between steps to .25 seconds

ClipBoard := ""

SendInput ^c
Sleep, %SleepTime%

if(ClipBoard <> "")
{
	ClipBoard := RegExReplace(ClipBoard, "[\W_]+")	; Remove non-letter, non-number, non-underscore
	Run http://www.dictionary.com/browse/%ClipBoard%
}
else 
{
	SendInput {Shift Down}{Ctrl Down}{Left}{Ctrl Up}{Shift Up}
	Sleep, %SleepTime%

	SendInput ^c
	Sleep, %SleepTime%

	ClipBoard := RegExReplace(ClipBoard, "[\W_]+")	; Remove non-letter, non-number, non-underscore
	Sleep, %SleepTime%
	Run http://www.dictionary.com/browse/%ClipBoard%

}
Return 


::-pdfc::		; Set Delivery Contact Status to Post Deliver Follow-up Complete

SleepTime := 500 ; set time to sleep between steps 

SendInput ^f
Sleep, %SleepTime%
SendRaw Delivery Contact Status
Sleep, %SleepTime%
SendInput {Escape}
Sleep, %SleepTime%
SendInput {Tab}{Enter}

Send ^f
Sleep, %SleepTime%
SendRaw Dependent Fields
Sleep, %SleepTime%
SendInput {Enter}
Sleep, %SleepTime%
SendInput {Enter} 
Sleep, %SleepTime%
Send {Escape}
Sleep, %SleepTime%
SendInput {Tab}{Tab}
SendInput {Down}{Down}{Down}{Down}
SendInput {Down}{Down}{Down}{Down}
SendInput {Down}{Down}{Down}{Down}
SendInput {Up}{Up}{Up}
Sleep, %SleepTime%
SendInput {Tab}{Enter}

save() ; save changes made to the delivery object

Return 


::-p::		;From the delivery object page, click the garage button and poke a VIN (Tested in Chrome Only)
::-poke::	;From the delivery object page, click the garage button and poke a VIN (Tested in Chrome Only)

poke()

Return

poke()		;From the delivery object page, click the garage button and poke a VIN (Tested in Chrome Only)
{
	Send ^f ;(Ctrl+f for find)
	Sleep, 250
	Send WarpDelivery ;element just right of Garage button ("Garage" is too generic a word to search for)
	Sleep, 250
	Send {Escape}
	Sleep, 250
	Send +{Tab} ;Send shift + Tab to move to adjacent "Garage" button
	Sleep, 250
	Send {Enter} 
	Sleep, 7000

	Clipboard = ;clear clipboard
		LoopCount := 0
		while ((Clipboard <> "VITALS" or Clipboard <> "Vitals") and LoopCount < 11) ; wait until page loads as indicated by finding the text "vitals" or 10 attempts
		 {
		   Send {Ctrl Down}f{Ctrl Up}
		   Sleep, 150
		   Send Vitals{Enter}
		   Sleep, 150	
		   Send {Escape}
		   Sleep, 150
		   Send {Ctrl Down}c{Ctrl Up} ; set Clibpoard = "Vitals" if this page has loaded
		   LoopCount++ 
		 }

		if (LoopCount = 10) ; Contents of garage page probably failed to load.
		{
			GarageError++
			SetTitleMatchMode, 2
			IfWinExist, Garage
			{
				WinClose, Garage ;close garage window, restoring focus to delivery object
			}
			Return ; Exit this hotkey
		}

	Send ^f ;(Ctrl+f for find)
	Send Poke
	Send {Escape} 	; focus on poke button by closing find bar
	Send {Enter}  	; poke
	Sleep, 3000 
	SetTitleMatchMode, 2
	WinClose, Garage  	; Close poke window
	Return

}

::-kronos::		; Type URL for Kronos
SendInput http://kronos7.teslamotors.com/wfc/navigator/logon
Return

::-matrix::		; Open the NJ Delivery Paperwork Matrix
Run, "\\teslamotors.com\us\Sales\Delivery\NewJersey\Paperwork Example Matrix\NJ Paperwork Example Matrix.html"
Return

::-contracts::	; Type phone number for the Contracts Team
SendInput 650.681.6872
Return

::-docs::		; Display reminder of document naming convention with always-on-top window
MsgBox, 4096,,Executed_PMT$xx.xx_DD_MVPA_Lease/RIC_DMV_RN_ShortVIN_LastName_FED-EX
return


::-rental::		; Type rental car instructions for customers
::-arms::		; Type rental car instructions for customers
SendInput We have created an authorization for a rental car that works at participating enterprise locations (please note that some locations, such as airports, do not allow authorization codes).  You can return your rental to our service center on your delivery date.  We’re hopeful we’ll be able to meet you sooner, but out of an abundance of caution, we’ve authorized the rental through ___________________________.{Enter}{Enter}Please call the Enterprise location closest to you (locator link: https://goo.gl/eWguh3) and provide them with “reservation number ___________________________."  This is really just an authorization code that can be transferred to the most convenient location--please note that you'll need to confirm the actual reservation details with the local Enterprise location.  Make sure that the location you plan to visit will have the type of vehicle you need available and let them know if you’ll need a lift (the closest location to you can pick you up).
return

::-app::		; Type URL for submitting credit application
::-credit::		; Type URL for submitting credit application
SendInput https://www.teslamotors.com/user/login{Enter}OR (if you have not yet made your order payment){Enter}http://teslamotors.com/creditapp{Enter}
return

::-wire::		; Type Wire Transfer instructions for a customer
SendInput Wire Transfer{enter}{enter} Please include your name and your order number - RN _____ when paying by wire transfer. {enter}{enter} Bank Name: Wells Fargo Bank, N.A.{enter}{enter} Bank Address: 420 Montgomery San Francisco, CA 94104{enter}{enter} Account Name: Tesla Motors Inc.{enter}{enter}Account number: 4000118323 {enter}{enter}ABA/Routing Number: 121000248{enter}{enter} Note: Your name, RN _____ {enter}{enter}
Return


::-g::		;Click the generate delivery configuration button 
Send ^f ;(Ctrl+f for find)
Sleep, 50
Send Delivery History[ ; Unique search phrase just above buttons
Sleep, 50
Send {Escape}
Sleep, 50
Send {Tab}
Sleep, 50
Send {Tab}
Sleep, 50
Send {Tab}
Sleep, 50
Send {Tab}
Sleep, 50
Send {Tab}
Sleep, 50
Send {Tab} 
Sleep, 50
Send {Enter}
Return

::-save::		;Click the Save button at the top of a Delivery Object	[[TESTED IN CHROME]]
save()

Return 

save() ; Click the Save button at the top of a Delivery Object
{
	SendInput ^f ;(Ctrl+f for find)
	Sleep, 50 ;(wait .05 second)
	SendInput Delivery Record Overview ;Unique phrase just below "save" button
	Sleep, 50
	SendInput {Escape}
	Sleep, 50
	SendInput {Shift Down}{Tab}{Tab}{Shift Up}
	Sleep, 50
	SendInput {Enter}
	
	Return 
}

::-plugin::		; Type plugshare.com and Tesla Supercharger links
SendInput http://www.plugshare.com/{Enter}{Enter}http://www.supercharge.info{Enter}{Enter}http://www.teslamotors.com/supercharger{Enter}
Return

::-newowner::		; Type instructions for transfer of ownership
SendInput To transfer ownership of your vehicle and allow the new owner to access the car from the app, please send us the following: {Enter}{Enter}{Tab}1) Color copy or photo of new owner's current driver’s license or state ID{Enter}{Enter}{Tab}2) Copy or photo of the title OR registration to indicate transfer of ownership{Enter}{Enter}{Tab}3) New owner's email address to link to the account{Enter}{Enter}{Tab}4) Home or mailing address of new owner{Enter}{Enter}{Tab}5) Phone number of new owner
Return

::-carwash::		; Type warning about automatic car washes and instructions in case they want to completely disregard it :)
SendInput We don’t recommend the automatic car wash because the rollers can accumulate dirt and scratch the paint.  If you need to take it to an automatic wash, you may need to put it in “Tow Mode.”{Enter}{Enter}If you or an attendant can stay seated in driver's seat, all you'll need to do is shift into neutral and turn off the windshield wipers.  {Enter}{Enter}Otherwise, you'll need to turn off the windshield wipers, drive the car onto the dollies that pull it through, and tap the following center screen buttons as indicated on the attached image to lock the car in neutral before you get out of the driver’s seat:  {Enter}{Enter}1)	Controls --> 2) Settings Tab --> 3) Tow Mode{Enter}{Enter}{Enter}With your foot on the brake you can move the slider on the center screen to "Tow Mode."  When the car comes out on the other end, you'll need to put your foot on the brake, press the metal button on the shifter for park, and then push and hold the red button on the center screen to exit Tow Mode.{Enter}{Enter}Feel free to give our service team a call at 973-921-0925 EXT 2 for assistance.  If it's after hours, you can also call 877-798-3752. 
Return

::-track::		; Type FedEx note for Delivery Notes field
::-tracking::	; Type FedEx note for Delivery Notes field
::-fedex::		; Type FedEx note for Delivery Notes field
SendInput DMV docs tracking: {Enter}Finance Docs Tracking:  
Return


::-usbank::		; Type US Bank's contact phone number
Send 800-872-2657
Return

::-wellsfargo::		; Type Wells Fargo's contact phone number
::-wells::			; Type Wells Fargo's contact phone number
::-wfbank::			; Type Wells Fargo's contact phone number
::-wf::				; Type Wells Fargo's contact phone number
Send 877-246-1015  
Return

::-chase::			; Type Chase's contact phone number
Send 800-336-6675 
Return

::-td::				; Type TD Bank's contact phone number
Send 800-556-8172
Return

::-rtfm::			; Read the F-ing manual...
SendInput You can follow the links in my signature below to our service page to review videos of your vehicle's functionality.  You can also see a brief video about summon here: https://www.tesla.com/blog/summon-your-tesla-your-phone and I’ve attached your manual – hit Control + F and type Summon to find the relevant page (page 90).  If you have any questions, feel free to give me a call. You can access your vehicle's manual at http://www.tesla.com/mytesla
Return 

::-mvpa::	; Type canned response about MVPA
SendInput Please see your purchase agreement attached indicating $_______ financed and a final balance of $_________ which can be paid online at http://www.tesla.com/mytesla or at the time of your delivery by personal check. Please let me know if you have any questions at all.{Enter}{Enter}
Return

::-lojack::		; Tracking system info for insurance
SendInput Hi ____,{Enter}{Enter}Please let your insurance agent know that your Model S comes with a tracking system that you can use if the vehicle were stolen.  You can learn more about it here:{Enter}{Enter}https://itunes.apple.com/us/app/tesla-model-s/id582007913?mt=8{Enter}{Enter}Please also see page 80 of the owner’s manual, which I’ve attached.{Enter}{Enter}Best,{Enter}Aron
Return

::-foff::		; Drop customer a hint
Send Since the Delivery Team is often in walk-throughs and away from our desks, please reach out to the service team directly at [DEFAULT] EXT 2 or email [DEFAULT] with any questions you might have in the future.  This will allow us to provide you with the fastest service going forward.
Return


::-plans::		; Type service plans page URL
Send https://www.tesla.com/support/service-plans
Return

::-pc::		; Type a post-call note
Send Hi ____,{Enter}{Enter}It was a pleasure speaking with you this morning.  I’ve tentatively scheduled your delivery for ____ but will keep in mind your preference to receive it sooner and definitely let you know if this becomes possible.{Enter}{Enter}I’ve highlighted your preferences below but feel free to give me a call any time with questions.  We can’t wait to meet you and introduce you to your new wheels.{Enter}{Enter}Best,{Enter}Aron
Return


::-hq::		; Type HQ's address
Send {Enter}{Tab}ATTN: Sales Administration{Enter}{Tab}Tesla Motors{Enter}{Tab}3500 Deer Creek{Enter}{Tab}Palo Alto, CA 94304{Enter}
return

::-fu::		; Fun understatement
SendRaw We can’t wait to meet you and introduce you to your new wheels!
return

::-incentives::		; Type link to incentives page
Send You can learn more about local and national incentives for your Model S at:{Enter}{Enter}{Tab}http://www.teslamotors.com/incentives/US/New`%2520Jersey{Enter}{Enter}http://www.irs.gov/Businesses/30D-New-Qualified-Plug-in-Electric-Drive-Motor-Vehicles--Tesla-Motors-Inc{Enter}{Enter}
return

::-d::		; Type timestamp
Send %A_DD%%A_MMM%%A_YYYY% [DEFAULT]:{Space}
return

::-t::		; Type the current hour
FormatTime, Time, ,htt:
SendInput, %Time%
Return

::-alliant::	; Type Alliant application info
SendInput Alliant Credit Union:{Enter}{Enter}Apply online: https://www.alliantcreditunion.com/Applications/LOS/TeslaHome{Enter{Enter}Phone: 800-328-1935{Enter}{Enter}tesla@alliantcreditunion.com 
Return


::-techcu::		; Type TCU contact info
::-tcu::		; Type TCU contact info
SendInput Application link: https://www.techcu.com/tesla/{Enter}{Enter}Contact #: 1-855-845-7060{Enter}{Enter}tesla@techcu.com{Enter}
Return

::-cust::	; Type month and "Tesla Customer" for saving a contact name
SendInput %A_MMM%%A_YYYY% Tesla Customer
return

::-store::		; Type Tesla store URL
::-shop::		; Type Tesla store URL
SendInput http://shop.tesla.com
return

::-charging::	; Type charging contact info
SendInput  650.681.6133 or ChargingInstallation@teslamotors.com
return

::-tax::		; Type info on federal tax credit
SendInput You can download the appropriate form for the $7,500 federal tax credit here, at the bottom of the page, under “Claiming the Credit”:{Enter}{Enter}http://www.fueleconomy.gov/feg/taxevb.shtml{Enter}{Enter}Your accountant or tax professional can advise on the appropriate form to use and how to complete it. Please let me know if you have any questions.{Enter}{Enter}
Return

::-mytesla::	; Type link to My Tesla
Send https://www.tesla.com/mytesla/{Space}
return 

::-walk::			; Type links to walk-through videos
::-walk-through::	; Type links to walk-through videos
::-walkthrough::	; Type links to walk-through videos
SendInput https://www.teslamotors.com/support/model-s-walkthrough{Enter}https://www.teslamotors.com/support/model-x-walkthrough{Enter}https://www.teslamotors.com/support/touchscreen-overview{Enter}https://www.teslamotors.com/support/tesla-mobile-app-walkthrough{Enter}
return

::-autopilot::		; Type links to autopilot videos
sendinput Autosteer: https://www.youtube.com/watch?v=4dv1LtEStu0{Enter}Auto Lane Change: https://www.youtube.com/watch?v=PUbk2Ko-VU4{Enter}Autopark: https://www.youtube.com/watch?v=Cn4X6nFsfo0{Enter}Summon: https://youtu.be/sQuJ8GKTjFM{Enter}
return 

::-leasing::	; Type leasing contact info
SendInput Phone: You can reach our Tesla Leasing team at 844-837-5285 or by email at teslafinance@tesla.com.
return 

::-finance::	; Type contact for finance team
Send autofinance@teslamotors.com or 650-681-6789
return

Return ; End of Script
