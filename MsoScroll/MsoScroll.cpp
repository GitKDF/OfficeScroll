#include "stdafx.h"
#include "debugtrace.h"
#include <oleacc.h>		//needed for AccessibleObjectFromWindow
#pragma comment(lib, "oleacc.lib")		//include reference to .lib.  This only works in MSVS, have to add it to linker info for other compilers


//from windowsx.h, don't see a need to include it otherwise
#define GET_X_LPARAM(lp)                        ((int)(short)LOWORD(lp))
#define GET_Y_LPARAM(lp)                        ((int)(short)HIWORD(lp))

//Using this rather than Succeeded since a return of S_FALSE is not considered a fail, but we certainly didn't get back what we wanted
#define IS_OK(hr) (((HRESULT)(hr)) == 0)

//used to get scroll routing setting
#define SPI_GETMOUSEWHEELROUTING    0x201C
#define SPI_GETWHEELSCROLLCHARS   0x006C
#define MOUSEWHEEL_ROUTING_FOCUS 0
#define MOUSEWHEEL_ROUTING_HYBRID 1
#define MOUSEWHEEL_ROUTING_MOUSE_POS 2

//only defined in WinUser.h for Windows Vista or higher, project defined for Win2000+ in stdafx.h
#define WM_MOUSEHWHEEL                  0x020E


extern "C" IMAGE_DOS_HEADER __ImageBase;
HHOOK g_mouseHook;
HHOOK g_kbdHook;
IDispatch* g_pApplication;
LONG g_AppID;						//Excel=0, Word=1
BOOL g_bRecurse;

int suppressAlt = 0;				//flag/counter to supress the current alt keypress (or 2 if a sheet change in excel, since we'll trigger an extra by changing the keyboard state)
BOOL ignoreNextAlt = FALSE;			//flag to ignore the next alt keypress because it's one that we sent as an injected keypress (no way to identify injected keys on a WH_KEYBOARD hook)
BYTE kbdState[256];					//array to hold keyboard state to restore it after a ctrl+scroll event in excel
int restoreKbdState = 0;			//flag/counter to indicate that the keyboard state should be restored
bool scrollSheets = FALSE;			//flag to indicate if scrolling sheets in excel is enabled; default to 0 so we don't handle any ctrl+scroll events in word
int verticalScrollValue = 0;		//vertical scroll setting; default to 0 to use system setting
int horizontalScrollValue = 1;		//default to 1 so we do a single horiz scroll in Word, and this was previously the only setting for excel and is a reasonable default

//define struct to hold dimensions of excel panes
struct paneRect {
	long Top;
	long Left;
	long Right;
	long Bottom;
};


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
// Borrowed from somewhere on the web and modified to OutputDebugString
// Used to investigate keyboard events.  I'm not at all sure how stable it is,
// and I don't want it in anything I actually use, but I'll leave it here
// in case it's needed for investigative purposes in the future
//void printBits(size_t const size, void const * const ptr)
/*{
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
// Keyboard hook callback procedure
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
			
			//check to see if there are any keys besides alt held down, in which case we don't actually want to send a new one,
			//since the user didn't press and release alt on it's own with no other keys down to activate the keyboard shortcut hints
			BYTE currentKbdState[256];
			GetKeyboardState(currentKbdState);
			BOOL otherKeys = FALSE;
			for (int i = 0; i < 256; i++) {								//loop through all keys
				if (currentKbdState[i] & 0x80) {							//if a key is down (top bit set)
					if ((VK_MENU == i) || (VK_LMENU == i) || (VK_RMENU == i))	//if it's an Alt key
						continue;													//ignore it
					otherKeys = TRUE;											//otherwise set key down flag
					break;														//leave loop
				}
			}
			
			if (otherKeys) {								//if any other keys are also down
				DBGTRACE("Suppressed - Other key down\n");
				return 1;										//eat the keypress since we only want to send an alt keypress if it was pushed and released with no other keys down
			}

			//if we get here, we're on a keyup, no other keys pressed, no cancel event, so we want to send excel a single alt press (down then up)
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
// Function to change sheets in topmost excel window
void ScrollSheets(HWND msgHwnd, short zDelta, long x, long y)
{
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
		BOOL HwndMatch = FALSE;

		if IS_OK(AutoWrap(DISPATCH_PROPERTYGET, &vtActiveWindow, g_pApplication, L"ActiveWindow", 0)) {			//get activewindow
			if IS_OK(AutoWrap(DISPATCH_PROPERTYGET, &vtActiveWinHwnd, vtActiveWindow.pdispVal, L"Hwnd", 0)) {		//get it's hwnd
				HwndMatch = ((HWND)(LONG_PTR)vtActiveWinHwnd.lVal == GetAncestor(msgHwnd, GA_ROOT));						//compare it to the root handle of the window we're over
				VariantClear(&vtActiveWinHwnd);																				//we succeeded, so clear variable
			}
			VariantClear(&vtActiveWindow);																				//we succeeded, so clear variable
		}

		//if we're not over excel's "ActiveWindow" (even if it's not necessarily the active window with focus on the desktop) ignore the message
		if (!HwndMatch)	return;
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
			ip[i].mi.dx = x;			//x coord from current event
			ip[i].mi.dy = y;			//y coord from current event
			ip[i].mi.dwFlags = MOUSEEVENTF_WHEEL;						//scroll event
			ip[i].mi.time = 0;											//let the system handle time stamp
			ip[i].mi.dwExtraInfo = 0;									//hardware use only
			ip[i].mi.mouseData = (zDelta > 0) ? 120 : -120;				//single scroll (120) positive or negative based on current event's sign
		}
		SendInput(i, ip, sizeof(INPUT));								//send keypresses (i will be incremented to 1 based index count of the array)

		delete[] ip;												//clean up our heap variable
	}
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Get the pane under the cursor in Excel, used when scrollRouting is active
void GetPaneExcel(HWND hwnd, IDispatch* &iDisp, long x, long y, int paneCount)
{
	paneRect PR[2] = { 0 }; // new paneRect[paneCount];

	VARIANT vtPanes;
	VariantInit(&vtPanes);
	VARIANT vtPaneNumber;
	vtPaneNumber.vt = VT_I4;

	//get panes object of the window we're over
	if IS_OK(AutoWrap(DISPATCH_PROPERTYGET, &vtPanes, iDisp, L"Panes", 0))
	{
		VARIANT vtPane;
		VariantInit(&vtPane);

		//HRESULT to track error status
		HRESULT hr = S_OK;

		//if panecount is 2, we'll loop twice, if it's not 2 (meaning it's 4), we'll only loop once since we only need pane 1's rectangle
		for (int i = 0; i <= ((2 == paneCount) ? 1 : 0); i++)
		{
			vtPaneNumber.lVal = i + 1;  //add 1 because panes in Excel are a 1 based index

			//get the pane object for vtPaneNumber
			//oddly, Item is a method in Word, and a Property in Excel
			if IS_OK(hr += AutoWrap(DISPATCH_PROPERTYGET, &vtPane, vtPanes.pdispVal, L"Item", 1, vtPaneNumber))
			{
				VARIANT vtVisibleRange;
				VariantInit(&vtVisibleRange);
				//get the VisibleRange object for the current pane
				if IS_OK(hr += AutoWrap(DISPATCH_PROPERTYGET, &vtVisibleRange, vtPane.pdispVal, L"VisibleRange", 0))
				{
					VARIANT vtTop;
					VariantInit(&vtTop);
					VARIANT vtHeight;
					VariantInit(&vtHeight);
					VARIANT vtLeft;
					VariantInit(&vtLeft);
					VARIANT vtWidth;
					VariantInit(&vtWidth);

					//get the top dimension (client referenced value)
					if IS_OK(hr += AutoWrap(DISPATCH_PROPERTYGET, &vtTop, vtVisibleRange.pdispVal, L"Top", 0)) {
						//get the height dimension
						if IS_OK(hr += AutoWrap(DISPATCH_PROPERTYGET, &vtHeight, vtVisibleRange.pdispVal, L"Height", 0)) {
							//add top and height to get bottom dimension
							vtHeight.dblVal = vtTop.dblVal + vtHeight.dblVal;

							//convert
							if IS_OK(hr += AutoWrap(DISPATCH_METHOD, &vtHeight, vtPane.pdispVal, L"PointsToScreenPixelsY", 1, vtHeight)) {
								PR[i].Bottom = vtHeight.lVal + 2;
							}
						}

						if IS_OK(hr += AutoWrap(DISPATCH_METHOD, &vtTop, vtPane.pdispVal, L"PointsToScreenPixelsY", 1, vtTop)) {
							PR[i].Top = vtTop.lVal - 2;
						}
					}


					if IS_OK(hr += AutoWrap(DISPATCH_PROPERTYGET, &vtLeft, vtVisibleRange.pdispVal, L"Left", 0)) {
						if IS_OK(hr += AutoWrap(DISPATCH_PROPERTYGET, &vtWidth, vtVisibleRange.pdispVal, L"Width", 0)) {
							vtWidth.dblVal = vtLeft.dblVal + vtWidth.dblVal;
							if IS_OK(hr += AutoWrap(DISPATCH_METHOD, &vtWidth, vtPane.pdispVal, L"PointsToScreenPixelsX", 1, vtWidth)) {
								PR[i].Right = vtWidth.lVal + 2;
							}
						}
						if IS_OK(hr += AutoWrap(DISPATCH_METHOD, &vtLeft, vtPane.pdispVal, L"PointsToScreenPixelsX", 1, vtLeft)) {
							PR[i].Left = vtLeft.lVal - 2;
						}
					}

					//clear variants
					VariantClear(&vtTop);
					VariantClear(&vtHeight);
					VariantClear(&vtLeft);
					VariantClear(&vtWidth);

					VariantClear(&vtVisibleRange);
				}
				VariantClear(&vtPane);
			}
		}

		//Make sure that we didn't get an error in any of the preceeding COM calls
		if (S_OK == hr)
		{
			if (2 == paneCount) {
				//if there are only two panes, figure out if they are split horizontally or vertically
				if (PR[0].Right == PR[1].Left)							//if panes 1 & 2 share a common vertical edge
					vtPaneNumber.lVal = (x < PR[0].Right) ? 1 : 2;			//1 if cursor left of edge, 2 if right of it
				else													//panes share a common horizontal edge
					vtPaneNumber.lVal = (y < PR[0].Bottom) ? 1 : 2;			//1 if cursor above bottom edge, 2 if below
			}
			else
			{
				//if there are 4 panes, figure out which we are over (with 4 panes, they are, from 1 to 4, Top-Left, Top-Right, Bottom-Left, Bottom-Right)
				if (x < PR[0].Right)									//cursor position is left of right edge of Pane 1
					vtPaneNumber.lVal = (y < PR[0].Bottom) ? 1 : 3;			//1 if cursor above bottom edge, 3 if below
				else													//cursor position is right of right edge of Pane 1
					vtPaneNumber.lVal = (y < PR[0].Bottom) ? 2 : 4;			//2 if cursor above bottom edge, 4 if below
			}

			//get the pane we're over
			if IS_OK(AutoWrap(DISPATCH_PROPERTYGET, &vtPane, vtPanes.pdispVal, L"Item", 1, vtPaneNumber))
			{
				iDisp->Release();				//release the window dispatch pointer
				iDisp = vtPane.pdispVal;		//set the iDisp to the pane IDispatch pointer
				vtPane.pdispVal->AddRef();		//add a ref, since we'll still call variantclear for good practice
				VariantClear(&vtPane);
			}
		}
		VariantClear(&vtPanes);
	}
	VariantClear(&vtPaneNumber);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Get the pane under the cursor in Word, used when scrollRouting is active
void GetPaneWord(HWND hwnd, IDispatch* &iDisp, long x, long y)
{
	// Calling AccessibleObjectFromWindow on a _WwG window from word returns an IDispatch to the window object, not the
	// pane under the cursor, so [small|large]scroll will scroll the active pane in the window.  There is no way to get
	// the pane rectangle information from within the Word VBA object model, however unlike Excel, the two panes have
	// their own unique window handles.  Since there can only be two panes in a word window, we first find the handle
	// to the pane we're not over, then compare the top of their screen rectangles to see if we're over the top or bottom
	// pane, then set iDisp to that pane accordingly.  There is also a hidden _WwG window for each pane that we ignore.
	BOOL topPane;
	RECT paneRect[2] = { NULL };

	//get the screen coordinates of the pane we're over
	if (!GetWindowRect(hwnd, &paneRect[0]))
		return;		//bail if we hit an error

	//get HWND of parent window
	HWND parentHwnd = GetParent(hwnd);

	//Find the first child window with class _WwG
	HWND otherPaneHwnd = FindWindowEx(parentHwnd, NULL, L"_WwG", NULL);

	//loop as long as we keep getting a return value
	while (otherPaneHwnd != NULL) {
		//if the window we found is not the window we're over
		if (otherPaneHwnd != hwnd)
		{
			//get window style and check for function success
			LONG winStyle = GetWindowLong(otherPaneHwnd, GWL_STYLE);
			if (0 == winStyle)
				return;

			//check window style for WS_VISIBLE and break if true
			//this is done because for each visible _WwG window, there is a corresponding invisible one,
			//we want to find the other visible one and ignore the invisible ones
			if (winStyle & WS_VISIBLE)
				break;
		}

		//if we're here, we haven't found the window we're looking for, so get the next one
		otherPaneHwnd = FindWindowEx(parentHwnd, otherPaneHwnd, L"_WwG", NULL);
	}

	//make sure that we found a window handle
	if (NULL == otherPaneHwnd)
		return;

	//get screen coordinates of the other pane
	if (!GetWindowRect(otherPaneHwnd, &paneRect[1]))
		return;		//bail if we hit an error

	//if the top of the pane we're over (0 in the array) is less than the top of the other pane,
	//then we're over the top pane, or pane 1 as Word calls it (1 on top, 2 on the bottom)
	topPane = (paneRect[0].top < paneRect[1].top);

	VARIANT vtPanes;
	VariantInit(&vtPanes);
	VARIANT vt1;
	vt1.vt = VT_I4;
	vt1.lVal = 1;
	
	//get panes object of current window
	if IS_OK(AutoWrap(DISPATCH_PROPERTYGET, &vtPanes, iDisp, L"Panes", 0))
	{
		VARIANT vtPane;
		VariantInit(&vtPane);
		//use item method to get pane 1
		//We can't use this to get pane 1 or 2 according to which one we're over, because there appears to be
		//a bug that while we receive a valid reference to the pane 2 object, calling the [small|large]scroll
		//method on it has no effect.  Oddly, getting the second pane using the "Next" property of pane 1 works.
		//This is true not just in automation, but in VBA itself.
		if IS_OK(AutoWrap(DISPATCH_METHOD, &vtPane, vtPanes.pdispVal, L"Item", 1, vt1))
		{
			//if we're over the top pane, this is the one we want 
			if (topPane)
			{
				iDisp->Release();				//release the window dispatch pointer
				iDisp = vtPane.pdispVal;		//set the iDisp to the pane IDispatch pointer
				vtPane.pdispVal->AddRef();		//add a ref, since we'll still call variantclear for good practice
			}
			else
			{ //if we're over the bottom pane
				VARIANT vtNextPane;
				VariantInit(&vtNextPane);
				//use the "Next" property of pane 1 to get pane 2 (see note above as to why rather than getting pane 2 in the previous AutoWrap call)
				if IS_OK(AutoWrap(DISPATCH_PROPERTYGET, &vtNextPane, vtPane.pdispVal, L"Next", 0))
				{
					iDisp->Release();
					iDisp = vtNextPane.pdispVal;		//release the window dispatch pointer
					vtNextPane.pdispVal->AddRef();		//set the iDisp to the pane IDispatch pointer
					VariantClear(&vtNextPane);			//add a ref, since we'll still call variantclear for good practice
				}
			}
			VariantClear(&vtPane);
		}
		VariantClear(&vtPanes);
	}
	VariantClear(&vt1);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Function to get the pane under the cursor when scroll routing is active
void GetPane(HWND msgHwnd, IDispatch* &iDisp, long x, long y)
{
	//First we're going to make sure that we're not scrolling in a window where we're editing a cell in Excel,
	//in which case we revert to scrolling the active pane.  This is because Excel does strange things when
	//scrolling the non-active pane of the active window when doing an in place edit.
	//(Inserts a '+' character into the cell, sometimes scrolls the wrong pane, and doesn't redraw properly.)

	//ignore all of this if g_AppID is 1 means we're in word
	if (!g_AppID) {
		//get the process focus window
		HWND FocusWin = GetFocus();

		//check to see if the focus window returned non-NULL and has the same root window as the one we're over
		//there's no issue if we're scrolling panes in a window that's not the activeWindow
		if (FocusWin && (GetAncestor(FocusWin, GA_ROOT) == GetAncestor(msgHwnd, GA_ROOT))) {
			//if we're over the same root window as the focuswindow...
			wchar_t className[50];
			if (GetClassName(FocusWin, className, 50))	//if we successfully got the class name
				if ((0 == lstrcmpi(className, L"EXCEL6")) || (0 == lstrcmpi(className, L"EXCEL<")))
					return;		//return without doing anything if editing a cell, since we have the activePnae by default
		}
	}

	int paneCount = 0;
	VARIANT vtPanes;
	VariantInit(&vtPanes);
	
	//get Panes object of current window
	if IS_OK(AutoWrap(DISPATCH_PROPERTYGET, &vtPanes, iDisp, L"Panes", 0))
	{
		VARIANT vtPaneCount;
		VariantInit(&vtPaneCount);
		//get Count property of Panes object
		if IS_OK(AutoWrap(DISPATCH_PROPERTYGET, &vtPaneCount, vtPanes.pdispVal, L"Count", 0))
		{
			//store the retreived value
			paneCount = vtPaneCount.intVal;
			VariantClear(&vtPaneCount);
		}
		VariantClear(&vtPanes);
	}

	//if there is more than one pane (or still 0 because there was an error getting the pane count)
	if (paneCount > 1)
	{
		//call application specific GetPane function
		if (g_AppID)
			GetPaneWord(msgHwnd, iDisp, x, y);
		else
			GetPaneExcel(msgHwnd, iDisp, x, y, paneCount);
	}
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Function to verify that the window handle is to the EXCEL7 or _WwG pane, not another control (e.g. the ribbon)
// (if the handle is the EXCEL6 inPlaceEdit box, set msgHwnd passed by reference to the EXCEL7 pane for the window)
BOOL VerifyRelevantHwnd(const DWORD &scrollRouting, const POINT &pt, HWND &msgHwnd)
{
	HWND checkwin;				//Handle used to check what window we're over
	if (MOUSEWHEEL_ROUTING_MOUSE_POS != scrollRouting)		//if scroll routing is not to the window under the cursor...
		checkwin = WindowFromPoint(pt);	//...get the window under the cursor
	else						//otherwise the window receiving the message is the same as the one under the cursor
		checkwin = msgHwnd;

	//check to make sure that we're over a control where we handle scroll events
	//we only care if we're over the main EXCEL7 or _WwG (word) panes (otherwise we can trap events if we're over the ribbon, which then prevents changing ribbons by scrolling)
	wchar_t className[50];
	if (!GetClassName(checkwin, className, 50))		//if we couldn't get the class name for some reason...
		return FALSE;									//don't handle the message


	if (g_AppID)
	{		//if g_AppID is 1 means we're in Word
		if (0 != lstrcmpi(className, L"_WwG"))		//if we're not over a Word pane (_WwG)
				return FALSE;							//don't handle the message
	}
	else
	{		//otherwise we're in Excel
		//If scrollRouting is to the cursor and we're over an EXCEL; control (the box to the left of the formula bar),
		//we just want to ditch these events all together since if we pass it on, excel will scroll the active window.
		//We do this by returning false and setting msgHwnd to NULL, which we'll check for when calling this function.
		//If scrollrouting is to focus, then passing it on is fine since an activeWindow scroll is expected behavior.
		if ((MOUSEWHEEL_ROUTING_MOUSE_POS == scrollRouting) && (0 == lstrcmpi(className, L"EXCEL;"))) {
			msgHwnd = NULL;
			return FALSE;
		}

		//if we're not over the excel pane (EXCEL7), or the excel edit-in-cell box (EXCEL6)
		if ((0 != lstrcmpi(className, L"EXCEL7")) && (0 != lstrcmpi(className, L"EXCEL6")))
			return FALSE;				//don't handle the message

		//the following is to check if the window receiving the message is the excel edit-in-place
		//box and if so change msgHwnd (passed by reference) to the main Excel Pane

		//if scroll is not to the window under the cursor, first we need to get the classname of the window receiving the message
		//we only need to do this if scroll routing is focus or hybrid, since otherwise they're the same window, set a few lines above (checkwin = msgHwnd)
		if (MOUSEWHEEL_ROUTING_MOUSE_POS != scrollRouting)
			if (!GetClassName(msgHwnd, className, 50))		//if we couldn't get the class name for some reason...
				return FALSE;									//don't handle the message

		if (0 == lstrcmpi(className, L"EXCEL6")) {					//if we're over the editInPlace box
			msgHwnd = GetParent(msgHwnd);								//get it's parent window ("XLDESK")
			msgHwnd = FindWindowEx(msgHwnd, NULL, L"EXCEL7", NULL);		//get it's child with class "EXCEL7", which is the excel pane
		}
	}

	return TRUE;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Mouse hook callback procedure
#define CallNext return g_bRecurse = FALSE, CallNextHookEx(g_mouseHook, nCode, wMsg, lParam)
#define CallCancel return g_bRecurse = FALSE, 1
LRESULT CALLBACK MouseHookProc(int nCode, WPARAM wMsg, LPARAM lParam)
{
	if (g_bRecurse || (nCode < 0)) return CallNextHookEx(g_mouseHook, nCode, wMsg, lParam);

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
	if ((wMsg == WM_MOUSEWHEEL) || (wMsg == WM_MOUSEHWHEEL))
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

		//get handle to window receiving scroll message
		HWND msgHwnd = ((LPMOUSEHOOKSTRUCTEX)lParam)->hwnd; 

		//set flag to suppress the current alt keypress if alt is down, even if we don't handle the event (e.g. cursor is over the ribbon)
		if (alt) suppressAlt = TRUE;

		//if ctrl is down, check to see whether we're handling scroll sheets events, as that's the only time we act on a ctrl+scroll event
		if (ctrl && !scrollSheets)
				CallNext;

		if (WM_MOUSEHWHEEL == wMsg) {		//if it's a horizontal scroll message
			if (shift || ctrl)					//and shoft or ctrl is down
				CallNext;						//pass the message on since we don't handle those events
		}

		//get system setting of whether to scroll focus window or window under cursor, or default to 0 if there's an error which implies pre windows 8
		//scroll routing is:
		//	MOUSEWHEEL_ROUTING_FOCUS     = 0 to focus window
		//	MOUSEWHEEL_ROUTING_HYBRID    = 1 to hybrid, focus window for desktop apps, window under cursor for store apps (windows 8.0 only)
		//	MOUSEWHEEL_ROUTING_MOUSE_POS = 2 to window under cursor (windows 8.1+)
		DWORD scrollRouting;
		if (!SystemParametersInfo(SPI_GETMOUSEWHEELROUTING, 0, &scrollRouting, 0))
			scrollRouting = MOUSEWHEEL_ROUTING_FOCUS;

		POINT pt = ((LPMOUSEHOOKSTRUCTEX)lParam)->pt;

		//get scroll distance; positive value is scroll up; divide by wheel_delta which is 120, since that is one "Click"
		//mousewheels with no click for finite control will not react if scrolled too slowly, since this will get truncated to 0 during integer division
		short zDelta = GET_WHEEL_DELTA_WPARAM(((LPMOUSEHOOKSTRUCTEX)lParam)->mouseData) / WHEEL_DELTA;
		
		//the following handles ctrl+alt+scroll events in excel to change sheets
		if (ctrl) {										//if ctrl is down
			if (!alt || shift)	//if alt is not down, or shift is also down, or it's a horizontal scroll message
				CallNext;									//pass the message on since we only want ctrl+alt+scroll events, no other modifiers

			ScrollSheets(msgHwnd, zDelta, pt.x, pt.y);

			//event handled, do not call next hook
			CallCancel;
		}

		//if we get here, we're over a word or excel window, and doing a scroll action other than ctrl+alt for a sheet change in excel

		//call function to make sure that we're over a window where we want to handle scroll events
		if (!VerifyRelevantHwnd(scrollRouting, pt, msgHwnd))
			if (NULL == msgHwnd)		//if the function returned false, and set msgHwnd to NULL...
				CallCancel;					//then we want to suppress this scroll message completely
			else						//if it returned false without changing msgHwnd...
				CallNext;					//then pass the message on

		//use an actual IDispatch rather than a variant to get window under cursor (rather than active window as previous version)
		//we don't need to worry about getting the active window if the system is set to scroll the focus window, as it will be the
		//active window that receives the scroll message anyway, not the window under cursor
		IDispatch* IDAppWin = NULL;
		if IS_OK(AccessibleObjectFromWindow(msgHwnd, OBJID_NATIVEOM, IID_IDispatch, (void **)&IDAppWin))  //get IDispatch to window under cursor
		{
			//if scrollrouting is to the window under the cursor, we're going to get the active pane of the window under the cursor
			if (MOUSEWHEEL_ROUTING_MOUSE_POS == scrollRouting)
				GetPane(msgHwnd, IDAppWin, pt.x, pt.y);

			VARIANT vtL, vtR, vtU, vtD;
			//set all direction variants to VB Long type
			vtL.vt = VT_I4;	vtR.vt = VT_I4;	vtU.vt = VT_I4;	vtD.vt = VT_I4;
			//initialize all direction variants to 0
			vtL.lVal = 0;	vtR.lVal = 0;	vtU.lVal = 0;	vtD.lVal = 0;


			if (shift || (WM_MOUSEHWHEEL == wMsg)) {	//if shift key or horizontal scroll message
				if (0 == horizontalScrollValue)	{			//if horizontal scroll value is set to 0 to use system setting
					//get system setting of how many characters/columns to scroll, default to 3 if there's an error
					unsigned int systemScrollColumns;
					if (!SystemParametersInfo(SPI_GETWHEELSCROLLCHARS, 0, &systemScrollColumns, 0))
						systemScrollColumns = 1;					//default to 1 if we couldn't get the system setting
					zDelta *= systemScrollColumns;				//multiply zDelta by systemScrollColumns to scroll in increments of systemScrollColumns
				} else {									//otherwise there is a custom value
					zDelta *= horizontalScrollValue;			//multiply zDelta by horizontalScrollValue to scroll in increments of horizontalScrollValue
				}

				if (WM_MOUSEHWHEEL == wMsg)		//if it's a horizontal scroll message
					zDelta *= -1;					//negate zDelta since a positive value is a right scroll, not left

				//zDelta is positive for scroll left, if negative for scroll right then excel will scroll right with negative left value
				vtL.lVal = zDelta;
			} else {
				//vertical scroll
				//get system setting of how many lines to scroll, default to 3 if there's an error getting the value
				unsigned int systemScrollLines;
				if (!SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, &systemScrollLines, 0))
					systemScrollLines = 3;

				//if system scroll setting is by page, swap alt flag to swap alt functionality, i.e. holding alt is small scroll, large scroll without
				if (-1 == systemScrollLines) alt = !alt;

				if (!alt)											//if not a large scroll
					//multiply zDelta to scroll by scroll lines setting
					zDelta *= (0 != verticalScrollValue)				//if custom setting is not equal to use system setting
									? verticalScrollValue					//then use custom value
									: (-1 == systemScrollLines)				//otherwise, if system setting is scroll by page
										? 3										//then default to 3
										: systemScrollLines;						//otherwise use system setting

				//zDelta is positive for scroll up, if negative for scroll down then excel will scroll down with negative up value	
				vtU.lVal = zDelta;
			}

			//scroll by line/cell or by page depending on alt key
			LPOLESTR pstrMethodName = alt ? L"LargeScroll" : L"SmallScroll";
			
			//send scroll command
			HRESULT scrollResult = AutoWrap(DISPATCH_METHOD, NULL, IDAppWin, pstrMethodName, 4, vtL, vtR, vtU, vtD); //Left, Right, Up, Down (reverse order!)

			//memory cleanup
			VariantClear(&vtL); VariantClear(&vtR); VariantClear(&vtU); VariantClear(&vtD);
			IDAppWin->Release();

			if IS_OK(scrollResult)
				CallCancel;		//handle the event if the scroll succeeded
			else
				CallNext;			//let the application handle it if it didn't
		} 
		else
		{	//if we didn't succeed in getting the window...
			//don't pass on the scroll message if shift is down to avoid bug where Excel 97-2003 crashes with a Shift+Scroll with no workbook open
			if (shift)
				CallCancel;
		}	//otherwise we'll fall through and the next hook will be called letting the system handle the message
	}

	CallNext;
}
#undef CallNext
#undef CallCancel

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Called by Excel to set custom settings within the mousehook
EXTERN_C void STDAPICALLTYPE SendSettings(int vert, int horiz, bool shts)
{
	verticalScrollValue = vert;
	horizontalScrollValue = horiz;
	scrollSheets = shts;
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
	//keyboard hook added to supress an alt keypress when page scrolling.  Without this,
	//doing so would show the keyboard shortcut hints, and pressing a key would activate them
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
	if (pApplication == NULL) return E_INVALIDARG;
	HRESULT hr = S_OK;

	_ASSERTE(g_pApplication == NULL);
	if (g_pApplication == NULL)
	{
		//get application name
		VARIANT vtAppName;
		VariantInit(&vtAppName);
		if SUCCEEDED(AutoWrap(DISPATCH_PROPERTYGET, &vtAppName, pApplication, L"Name", 0))
		{
			DBGTRACE("%s\n", vtAppName.bstrVal);
			if (0 == wcscmp(vtAppName.bstrVal, L"Microsoft Excel"))
				g_AppID = 0;
			else if (0 == wcscmp(vtAppName.bstrVal, L"Microsoft Word"))
				g_AppID=1;
			else hr = STG_E_UNIMPLEMENTEDFUNCTION;
			VariantClear(&vtAppName);
			g_pApplication=pApplication;
			g_pApplication->AddRef();
			DBGTRACE("MsoScroll::Connect\n");
		}
	}
	else hr = ERROR_ALREADY_ASSIGNED;

	if IS_OK(hr)
	{
		_ASSERTE(g_mouseHook == NULL);
		if (g_mouseHook == NULL)
		{
			g_mouseHook = SetWindowsHookEx(WH_MOUSE, MouseHookProc, (HINSTANCE)&__ImageBase, GetCurrentThreadId());
			DBGTRACE("MsoScroll::SetWindowsHookEx - Mouse\n");
		}
		else { hr = ERROR_ALREADY_EXISTS; DBGTRACE("Mousehook - ERROR_ALREADY_EXISTS\n"); }

		_ASSERTE(g_kbdHook == NULL);
		if (IS_OK(hr) && g_kbdHook == NULL)
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
