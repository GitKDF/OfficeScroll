#include "stdafx.h"
#include "debugtrace.h"
#include <oleacc.h>		//needed for AccessibleObjectFromWindow
#pragma comment(lib, "oleacc.lib")		//include reference to .lib.  This only works in MSVS, have to add it to linker info for other compilers


//from windowsx.h, don't see a need to include it otherwise
#define GET_X_LPARAM(lp)                        ((int)(short)LOWORD(lp))
#define GET_Y_LPARAM(lp)                        ((int)(short)HIWORD(lp))


//used to get scroll routing setting
#define SPI_GETMOUSEWHEELROUTING    0x201C


extern "C" IMAGE_DOS_HEADER __ImageBase;
HHOOK g_mouseHook;
HHOOK g_kbdHook;
IDispatch* g_pApplication;
//LONG g_AppID;//Excel=0, Word=1
BOOL g_bRecurse;

int suppressAlt = 0;				//flag/counter to supress the current alt keypress (or 2 if a sheet change in excel, since we'll trigger an extra by changing the keyboard state)
BOOL ignoreNextAlt = FALSE;			//flag to ignore the next alt keypress because it's one that we sent as an injected keypress (no way to identify injected keys on a WH_KEYBOARD hook)
BYTE kbdState[256];					//array to hold keyboard state to restore it after a ctrl+scroll event in excel
int restoreKbdState = 0;			//flag/counter to indicate that the keyboard state should be restored


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//   FUNCTION: AutoWrap(int, VARIANT*, IDispatch*, LPOLESTR, int,...)
//   PURPOSE: Automation helper function. It simplifies most of the low-level 
//      details involved with using IDispatch directly. Feel free to use it 
//      in your own implementations. One caveat is that if you pass multiple 
//      parameters, they need to be passed in reverse-order.
//   PARAMETERS:
//      * autoType - Could be one of these values: DISPATCH_PROPERTYGET, 
//      DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF, DISPATCH_METHOD.
//      * pvResult - Holds the return value in a VARIANT.
//      * pDisp - The IDispatch interface.
//      * ptName - The property/method name exposed by the interface.
//      * cArgs - The count of the arguments.
//   RETURN VALUE: An HRESULT value indicating whether the function succeeds or not.
//   EXAMPLE: 
//      AutoWrap(DISPATCH_METHOD, NULL, pDisp, L"call", 2, parm[1], parm[0]);
//
HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...)
{
	// Begin variable-argument list
	va_list marker;
	va_start(marker, cArgs);

	if (!pDisp) return E_INVALIDARG;

	// Variables used
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;

	// Get DISPID for name passed
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, /*LOCALE_USER_DEFAULT*/ LOCALE_SYSTEM_DEFAULT, &dispID);
	if (FAILED(hr))
	{
#ifdef _DEBUG
		OutputDebugString(_T("AutoWrap::IDispatch->GetIDsOfNames failed\n"));
		_com_error err(hr);
		OutputDebugString(err.ErrorMessage()); OutputDebugString(_T("\n"));
#endif//_DEBUG
		return hr;
	}

	// Allocate memory for arguments
	VARIANT *pArgs = new VARIANT[cArgs + 1];
	// Extract arguments...
	for(int i=0; i < cArgs; i++)
	{
		pArgs[i] = va_arg(marker, VARIANT);
	}

	// Build DISPPARAMS
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;

	// Handle special-case for property-puts
	if (autoType & DISPATCH_PROPERTYPUT)
	{
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	// Make the call
	hr = pDisp->Invoke(dispID, IID_NULL, /*LOCALE_USER_DEFAULT*/ LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
	if FAILED(hr) 
	{
#ifdef _DEBUG
		OutputDebugString(_T("AutoWrap::IDispatch->Invoke failed\n"));
		_com_error err(hr);
		OutputDebugString(err.ErrorMessage()); OutputDebugString(_T("\n"));
#endif//_DEBUG
		return hr;
	}

	// End variable-argument section
	va_end(marker);
	delete[] pArgs;
return hr;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Borrowed from somewhere on the web and modified to OutputDebugString
//Used to investigate keyboard events.  I'm not at all sure how stable it is,
//and I don't want it in anything I actually use, but I'll leave it here
//in case it's needed for investigative purposes in the future
/*
void printBits(size_t const size, void const * const ptr)
{
	unsigned char *b = (unsigned char*)ptr;
	char byte[2];
	byte[1] = '\0';
	int i, j;

	for (i = size - 1; i >= 0; i--)		//iterate through all bytes
	{
		for (j = 7; j >= 0; j--)		//iterate through 8 bits in each byte
		{
			byte[0] = ((b[i] >> j) & 1) + '0';
			OutputDebugStringA(byte);
		}
	}
	OutputDebugStringW(L"\n");
}
*/

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LRESULT CALLBACK KbdMsgProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	if (nCode < 0) return CallNextHookEx(g_kbdHook, nCode, wParam, lParam);
	
	if (VK_MENU == wParam)				//if the key is ALT
	{
		if (HIWORD(lParam) & KF_UP)		//alt keyup event
		{
			DBGTRACE("AltUp - nCode: %d - ", nCode);

			if (ignoreNextAlt)			//if flag is set to ignore the next alt press, since we sent it using sendinput below on the last call
			{
				DBGTRACE("ignored");
				//if the callback is called because the message is being polled by peek message rather than get message,
				//which Excel does several times, then nCode is HC_NOREMOVE instead of HC_ACTION, and we want to just pass it on
				//we want to pass on all of those until the key up event, with nCode == HC_ACTION meaning that the message will be removed
				//then we reset our flag and will process subsequent Alt presses
				if (HC_ACTION == nCode)
				{
					ignoreNextAlt = FALSE;			//reset our flag to ignore alt presses
					DBGTRACE(" - ignoreNextAlt reset");
				}
				DBGTRACE("\n");

				//ignore this alt keypress and let excel handle it
				return CallNextHookEx(g_kbdHook, nCode, wParam, lParam);
			}

			if (suppressAlt)						//if our flag to cancel the alt press is set because the user scrolled with the alt key down
			{
				DBGTRACE("Suppressed\n");
				return 1;							//eat the keypress
			}
			
			//if we get here, we're on a keyup, no cancel event, so we want to send excel a single alt press (down then up)
			ignoreNextAlt = TRUE;			//set flag to indicate that we're sending a keypress and that we should ignore subsequent alt keypresses until there's an alt up event with nCode HC_ACTION

			INPUT ip[2] = { 0 };											//2 for down then up events
			ip[0].type = INPUT_KEYBOARD;									//set events to kbd
			ip[1].type = INPUT_KEYBOARD;
			ip[0].ki.wScan = MapVirtualKey(VK_MENU, MAPVK_VK_TO_VSC);		//hardware scan code for key
			ip[1].ki.wScan = MapVirtualKey(VK_MENU, MAPVK_VK_TO_VSC);
			ip[0].ki.time = 0;												//let the system handle the time stamp
			ip[1].ki.time = 0;
			ip[0].ki.dwExtraInfo = 0;										//hardware use only
			ip[1].ki.dwExtraInfo = 0;
			ip[0].ki.dwFlags = KEYEVENTF_EXTENDEDKEY | KEYEVENTF_SCANCODE;	//extended key for alt key, ORd with scancode to indicate that we're passing by scancode, not VKCode
			ip[1].ki.dwFlags = KEYEVENTF_EXTENDEDKEY | KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;		//same but with key up

			SendInput(2, ip, sizeof(INPUT));								//send keypresses
						
			DBGTRACE("Sent New Alt Keypress\n");
			return 1;														//eat the current keypress event
		}
		else						//alt keydown event
		{
			DBGTRACE("AltDown - nCode: %d - ", nCode);

			if (ignoreNextAlt)		//if we're ignoring alt keypresses because we sent them with sendinput
			{
				DBGTRACE("ignored\n");
				return CallNextHookEx(g_kbdHook, nCode, wParam, lParam);		//just pass the event on
			}

			if (!(lParam & (1 << 30)))				//if lParam bit 30 is 0, which is the "Previous Key-State Flag", then this is the first keypress event, not a hold-down repeat event
			{
				if (suppressAlt)
				{
					//if suppressAlt is greater than 0, decrement it.  This is because in most cases we
					//can set it to true (1) to suppress the current alt keypress, but if we change sheets
					//we will trigger an additional keypress when we reset the keyboard state, so we need
					//to suppress 2 events in a row
					suppressAlt--;
					DBGTRACE("suppressAlt decremented - ");
				}
			}
			DBGTRACE("Suppressed\n");

			return 1;								//unless ignoreNextAlt is set, we eat all events
		}
	}
	else		//any other key
	{
		//if we press any other key while alt is down, we want to suppress the next alt up event
		//if alt is not down, this will be reset on the next altdown event (a few lines above)
		//if we don't do this, pressing e.g. Alt+F to open the file menu, then releasing alt
		//will send an altUP event, and disable allowing additional keypresses to continue
		//activating menus/commands by keyboard.
		suppressAlt = TRUE;
	}
		
	return CallNextHookEx(g_kbdHook, nCode, wParam, lParam);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LRESULT CALLBACK MouseHookProc(int nCode, WPARAM wMsg, LPARAM lParam)
{
	if (g_bRecurse || (nCode<0)) return CallNextHookEx(g_mouseHook, nCode, wMsg, lParam);
	//return CallNextHookEx(g_mouseHook, nCode, wMsg, lParam);

	if (1 == restoreKbdState)				//if flag is set to restore the keyboard state on this message
	{
		SetKeyboardState(kbdState);				//restore keyboard state
		restoreKbdState = 0;					//clear flag
	}
	else if (1 < restoreKbdState) {			//if count is greater than 1 because we want to process multiple sheet changes
		restoreKbdState--;						//decrement the counter and leave keyboard state for the current message
	}
	
	//if ((wMsg == WM_MOUSEWHEEL) && (HIWORD(GetKeyState(VK_SHIFT)) || HIWORD(GetKeyState(VK_MENU)))) //VK_MENU=ALT key
	//changed to handle all mousewheel events to also send non-modified events to the window under cursor rather than excel's default of active window
	if (wMsg == WM_MOUSEWHEEL)
	{
		//if (g_bRecurse) {DBGTRACE("prevent RECURSION-------\n");return 1;}
		g_bRecurse = TRUE;
		//DBGTRACE("WM_MOUSEWHEEL received\n");
		//DBGTRACE("hwnd=0x%x\n",((LPMOUSEHOOKSTRUCT)lParam)->hwnd);
		//DBGTRACE("WM_MOUSEWHEEL+VK_SHIFT|VK_MENU\n");

		//get key states
		bool ctrl = HIWORD(GetKeyState(VK_CONTROL));
		bool alt = HIWORD(GetKeyState(VK_MENU));
		bool shift = HIWORD(GetKeyState(VK_SHIFT));

		//set flag to suppress the current alt keypress if alt is down, even if we don't handle the event (e.g. cursor is over the ribbon)
		if (alt) suppressAlt = TRUE;

		//uncomment the following line to disable Ctrl + Alt + Scroll to change sheets in excel
		//if (ctrl) goto CallNext;

		//get handle to window receiving scroll message
		HWND msgHwnd = ((LPMOUSEHOOKSTRUCTEX)lParam)->hwnd; 

		//get system setting of whether to scroll focus window or window under cursor, or default to 0 if there's an error which implies pre windows 8
		//scroll routing is:
		//	MOUSEWHEEL_ROUTING_FOCUS     = 0 to focus window
		//	MOUSEWHEEL_ROUTING_HYBRID    = 1 to hybrid, focus window for desktop apps, window under cursor for store apps (windows 8.0 only)
		//	MOUSEWHEEL_ROUTING_MOUSE_POS = 2 to window under cursor (windows 8.1+)
		DWORD scrollRouting;
		if (!SystemParametersInfo(SPI_GETMOUSEWHEELROUTING, 0, &scrollRouting, 0))
			scrollRouting = 0;

		HWND checkwin;				//used to check what window we're over
		if (2 != scrollRouting)		//if scroll routing is not to the window under the cursor...
			checkwin = WindowFromPoint(((LPMOUSEHOOKSTRUCTEX)lParam)->pt);	//...get the window under the cursor
		else						//otherwise the window receiving the message is the same as the one under the cursor
			checkwin = msgHwnd;

		//check to make sure that we're over a control where we handle scroll events
		//we only care if we're over the main EXCEL7 or _WwG (word) panes (otherwise we can trap events if we're over the ribbon, which then prevents changing ribbons by scrolling)
		wchar_t className[50];
		if (!GetClassName(checkwin, className, 50))	//if we couldn't get the class name for some reason...
			goto CallNext;								//pass the message on

		//if we're not over the excel pane (EXCEL7), the excel edit-in-cell box (EXCEL6), or the Word pane (_WwG)
		if ((0 != lstrcmpi(className, L"EXCEL7")) && (0 != lstrcmpi(className, L"EXCEL6")) && (0 != lstrcmpi(className, L"_WwG")))
			if (GetAncestor(msgHwnd, GA_ROOT) == GetAncestor(checkwin, GA_ROOT))		//and we ARE over our own parent window
				goto CallNext;															//ignore the message
				//we don't care if we're not over the Excel/Word pane as long as we're over another window completely as can be the case with
				//scroll to focus window, but if we're not over the Excel/Word pane and we are over our window, then we're over the ribbon


		//if scroll is not to the window under the cursor, get the classname of the window receiving the message
		//we only need to do this if scroll routing is focus or hybrid, since otherwise they're the same window, set a few lines above (checkwin = msgHwnd)
		if (2 != scrollRouting)
			if (!GetClassName(msgHwnd, className, 50))	//if we couldn't get the class name for some reason...
				goto CallNext;								//pass the message on

		//if the window receiving the message is the excel edit-in-place box
		if (0 == lstrcmpi(className, L"EXCEL6")) {
			msgHwnd = GetParent(msgHwnd);								//get it's parent window ("XLDESK")
			msgHwnd = FindWindowEx(msgHwnd, NULL, L"EXCEL7", NULL);		//get it's child with class "EXCEL7", which is the excel pane
		}


		//get scroll distance; positive value is scroll up; divided by wheel_delta which is 120, since that is one "Click"
		//mousewheels with no click for finite control will not react if scrolled too slowly, since this will get truncated to 0 during integer division
		short zDelta = GET_WHEEL_DELTA_WPARAM(((LPMOUSEHOOKSTRUCTEX)lParam)->mouseData) / WHEEL_DELTA;



		//the following handles ctrl+alt+scroll events in excel to change sheets
		if (ctrl) {										//if ctrl is down
			if (!alt || shift)								//if alt is not down, or shift is also down
				goto CallNext;									//pass the message on since we only want ctrl+alt+scroll events, no other modifiers

			if (0 == lstrcmpi(className, L"_WwG"))			//if we're over a word window
				goto CallNext;									//pass the message on since we don't handle this event in word

			//get current focus handle
			HWND FocusHandle = GetFocus();

			//compare focus handle's root ancestor to the msg recipient's root ancestor
			//if they're different, this could be because the window under the mouse is:
			//	-not excel's topmost window, in which case we'll ignore the message
			//	-or because it IS excel's "ActiveWindow", but the application is not in focus, which we'll check for below
			//
			//but if they match already just move on with processing
			if (GetAncestor(msgHwnd, GA_ROOT) != GetAncestor(FocusHandle, GA_ROOT))
			{
				//they didn't match...
				VARIANT vtActiveWindow;
				VariantInit(&vtActiveWindow);
				VARIANT vtActiveWinHwnd;
				VariantInit(&vtActiveWinHwnd);

				//initialize flag to indicate match
				bool HwndMatch = FALSE;

				if SUCCEEDED(AutoWrap(DISPATCH_PROPERTYGET, &vtActiveWindow, g_pApplication, L"ActiveWindow", 0)) {			//get activewindow
					if SUCCEEDED(AutoWrap(DISPATCH_PROPERTYGET, &vtActiveWinHwnd, vtActiveWindow.pdispVal, L"Hwnd", 0)) {		//get it's hwnd
						HwndMatch = ((HWND)(LONG_PTR)vtActiveWinHwnd.lVal == GetAncestor(msgHwnd, GA_ROOT));						//compare it to the root handle of the window we're over
						VariantClear(&vtActiveWinHwnd);																				//we succeeded, so clear variable
					}
					VariantClear(&vtActiveWindow);																				//we succeeded, so clear variable
				}

				//if we're not over excel's "ActiveWindow" (even if it's not necessarily the active window with focus on the desktop) ignore the message
				if (!HwndMatch)	goto CallCancel;
			}

			//Set flag to suppress alt keypresses.  In the normal scroll case we can set it to true (1)
			//to suppress the current alt keypress, but if we change sheets we will trigger an additional
			//keypress when we reset the keyboard state, so we need to suppress 2 events in a row.
			suppressAlt = 2;

			//pick virtual key to send, pg up or down, based on scroll direction
			int VKey = (zDelta < 0) ? VK_NEXT : VK_PRIOR;
			
			//build lParam to send with WM_KEY* messages containing scan code of key to send
			int keylParam = MapVirtualKey(VKey, MAPVK_VK_TO_VSC);

			//save current keyboard state to global variable
			GetKeyboardState(kbdState);

			//build new keyboard state with just CTRL down
			BYTE newkbdState[256] = { 0 };
			newkbdState[VK_CONTROL] = 0x80;		//set high bit means key down
			SetKeyboardState(newkbdState);		//set keyboard state to our spoofed ctrl down state

			PostMessage(msgHwnd, WM_KEYDOWN, VKey, keylParam);						 //send keydown
			PostMessage(msgHwnd, WM_KEYUP, VKey, keylParam | (1 << 30) | (1 << 31)); //send keyup, OR keycode in keylParam to set bits 30 & 31 ('previous key-state' and 'transition-state' flags)

			//set flag to restore keyboard state on next call.  We have to do this because setting it will change it for the next input processing event,
			//which can't happen until we exit this procedure for the current event.  If we set it back here/now, the restored state will be seen at the
			//next input event after we exit this function and our modified state will never be seen.
			restoreKbdState = TRUE;

			//if the user scrolled quickly, we will send ourselves multiple scroll events with a zDelta of 120 (single scroll)
			if (abs(zDelta) > 1)
			{
				INPUT* ip;
				ip = new INPUT[abs(zDelta) - 1];							//one less than zDelta since we already handled one in this event
				int i;
				for (i = 0; i < (abs(zDelta) - 1); i++)						//iterate through array from a 0 index
				{
					restoreKbdState++;											//increment restoreKbdState flag to let us know we'll be processing multiple sheet changes
					ip[i].type = INPUT_MOUSE;									//set event to mouse
					ip[i].mi.dx = ((LPMOUSEHOOKSTRUCTEX)lParam)->pt.x;			//x coord from current event
					ip[i].mi.dy = ((LPMOUSEHOOKSTRUCTEX)lParam)->pt.y;			//y coord from current event
					ip[i].mi.dwFlags = MOUSEEVENTF_WHEEL;						//scroll event
					ip[i].mi.time = 0;											//let the system handle time stamp
					ip[i].mi.dwExtraInfo = 0;									//hardware use only
					ip[i].mi.mouseData = (zDelta > 0) ? 120 : -120;				//single scroll (120) positive or negative based on current event's sign
				}
				SendInput(i, ip, sizeof(INPUT));								//send keypresses (i will be incremented to 1 based index count of the array)

				delete[] ip;												//clean up our heap variable
			}

			//event handled, do not call next hook
			goto CallCancel;
		}


		//if we get here, we're over a word or excel window, and doing a scroll action other than ctrl+alt for a sheet change in excel


		//use an actual IDispatch rather than a variant to get window under cursor (rather than active window as previous version)
		//we don't need to worry about getting the active window if the system is set to scroll the focus window, as it will be the
		//active window that receives the scroll message anyway, not the window under cursor
		IDispatch* IDAppWin = NULL;
		
		if SUCCEEDED(AccessibleObjectFromWindow(msgHwnd, OBJID_NATIVEOM, IID_IDispatch, (void **)&IDAppWin))  //get IDispatch to window under cursor
		{
			//get system setting of how many lines to scroll, default to 3 if there's an error
			unsigned int scrollLines;
			if (!SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, &scrollLines, 0))
				scrollLines = 3;

			//if system scroll setting is by page, swap alt flag to swap alt functionality, i.e. holding alt is small scroll, large scroll without
			if (-1 == scrollLines) alt = !alt;

			//scroll by line/cell or by page depending on alt key
			LPOLESTR pstrMethodName = alt ? L"LargeScroll" : L"SmallScroll";
			
			VARIANT vtL, vtR, vtU, vtD;
			//set all direction variants to VB Long type
			vtL.vt = VT_I4;	vtR.vt = VT_I4;	vtU.vt = VT_I4;	vtD.vt = VT_I4;	//VB Long type
			
			//initialize all direction variants to 0
			vtL.lVal = 0;	vtR.lVal = 0;	vtU.lVal = 0;	vtD.lVal = 0;


			if (shift)	//if shift key means horizontal
			{
				if (zDelta < 0)	vtR.lVal = abs(zDelta);				//scroll down = right
				else vtL.lVal = zDelta;								//scroll up = left
			}
			else								//no shift key means vertical
			{
				if (!alt)											//if not a large scroll
					zDelta *= (-1 == scrollLines) ? 3: scrollLines;		//multiply zDelta by scrollLines, or if scroll by page then default to 3 (the windows default scroll speed)
				if (zDelta < 0)	vtD.lVal = abs(zDelta);				//scroll down
				else vtU.lVal = zDelta;								//scroll up
			}

			//send scroll command
			HRESULT scrollResult = AutoWrap(DISPATCH_METHOD, NULL, IDAppWin, pstrMethodName, 4, vtL, vtR, vtU, vtD); //Left, Right, Up, Down (reverse order!)

			//memory cleanup
			VariantClear(&vtL); VariantClear(&vtR); VariantClear(&vtU); VariantClear(&vtD);
			IDAppWin->Release();

			if SUCCEEDED(scrollResult)
				goto CallCancel;		//handle the event if the scroll succeeded
			else
				goto CallNext;			//let the application handle it if it didn't

		} //if we didn't succeed in getting the window, we'll fall through and the next hook will be called letting the system handle the message
	}


CallNext:
	g_bRecurse=FALSE;
	return CallNextHookEx(g_mouseHook, nCode, wMsg, lParam);

CallCancel:
	g_bRecurse=FALSE;
	return 1;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
STDAPI Disconnect()
{
	//_ASSERTE(g_mouseHook);
	if (g_mouseHook)
	{
		UnhookWindowsHookEx(g_mouseHook);
		g_mouseHook = NULL;
		DBGTRACE("MsoScroll::UnhookWindowsHookEx - Mouse\n");
	}
	//keyboard hook added to supress an alt keypress when page scrolling
	//before, doing so would show the keyboard shortcut hints, and pressing a key would activate them
	//if we scroll while holding alt, we suppress the event.  if we release alt without scrolling, we
	//send an alt keypress event.  This means that to see the keyboard shortcut hints, the user must
	//press and release alt, whereas the default behaviour is to show them after a second or so of
	//holding down alt.  The alt+key combinations all still work as expected without the hints showing,
	//you just have to release the key to see the hints.
	//I'm not 100% sure what this would do in pre 2007 office, but I suspect it would be essentially
	//the same behaviour but with the underlined letters on menus, i.e. keys would still work, but
	//underline hints wouldn't show until releasing alt key
	if (g_kbdHook)
	{
		UnhookWindowsHookEx(g_kbdHook);
		g_kbdHook = NULL;
		DBGTRACE("MsoScroll::UnhookWindowsHookEx - Keyboard\n");
	}
	//_ASSERTE(g_pApplication);
	if (g_pApplication)
	{
		g_pApplication->Release();
		g_pApplication = NULL;
		DBGTRACE("MsoScroll::g_pApplication->Release\n");
	}
	return S_OK;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
STDAPI Connect(IDispatch *pApplication)
{
	if (pApplication==NULL) return E_INVALIDARG;
	HRESULT hr=S_OK;

	_ASSERTE(g_pApplication==NULL);
	if (g_pApplication==NULL)
	{
		//get application name
//		VARIANT vtAppName;
//		VariantInit(&vtAppName);
//		if SUCCEEDED(AutoWrap(DISPATCH_PROPERTYGET, &vtAppName, pApplication, L"Name", 0))
//		{
//OutputDebugString(vtAppName.bstrVal);
//			if(0 == wcscmp(vtAppName.bstrVal, L"Microsoft Word")) g_AppID=1;
//			VariantClear(&vtAppName);
		g_pApplication=pApplication;
		g_pApplication->AddRef();
		DBGTRACE("MsoScroll::Connect\n");
		//}
	}
	else hr=ERROR_ALREADY_ASSIGNED;

	if SUCCEEDED(hr)
	{
		_ASSERTE(g_mouseHook==NULL);
		if (g_mouseHook==NULL)
		{
			g_mouseHook=SetWindowsHookEx(WH_MOUSE, MouseHookProc, (HINSTANCE)&__ImageBase, GetCurrentThreadId());
			DBGTRACE("MsoScroll::SetWindowsHookEx - Mouse\n");
		}
		else {hr=ERROR_ALREADY_EXISTS; DBGTRACE("Mousehook - ERROR_ALREADY_EXISTS\n");}

		_ASSERTE(g_kbdHook == NULL);
		if (g_kbdHook == NULL)
		{
			g_kbdHook = SetWindowsHookEx(WH_KEYBOARD, KbdMsgProc, (HINSTANCE)&__ImageBase, GetCurrentThreadId());
			DBGTRACE("MsoScroll::SetWindowsHookEx - Keyboard\n");
		}
		else { hr = ERROR_ALREADY_EXISTS; DBGTRACE("Keyboardhook - ERROR_ALREADY_EXISTS\n"); }
	}

	if (S_OK != hr)		//if there was an initialization error anywhere...
		Disconnect();		//call disconnect to unset hooks

	return hr;
}
