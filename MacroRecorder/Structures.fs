module Structures

open System
open System.Collections.Generic
open System.Runtime.InteropServices
open Microsoft.FSharp.Collections
open Microsoft.FSharp.Core
open System.Diagnostics
open System.Threading
open System.Collections.ObjectModel
open System.Runtime.InteropServices

type HookProc = delegate of int * IntPtr * IntPtr -> IntPtr  

[<DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, UInt32 dwThreadId)

[<DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
extern bool UnhookWindowsHookEx(IntPtr hhk)

[<DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam)

[<DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
extern IntPtr GetModuleHandle(string lpModuleName)

[<DllImport("user32.dll")>]
extern int GetAsyncKeyState(int vKey)

[<DllImport("user32.dll", SetLastError = true)>]
extern IntPtr SendMessage(IntPtr hWnd, uint32 Msg, IntPtr wParam, IntPtr lParam)


[<DllImport("user32.dll")>]
extern bool GetMessage(IntPtr lpMsg, IntPtr hWnd, uint32 wMsgFilterMin, uint32 wMsgFilterMax)

[<DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
extern int ToUnicode(
    System.UInt32 vkCode,
    System.UInt32 scanCode,
    byte[] keyState,
    [<Out>] char[] result,
    int resultSize,
    System.UInt32 flags
)

[<StructLayout(LayoutKind.Sequential)>]
type MSLLHOOKSTRUCT = struct
    val mutable pt: System.Drawing.Point
    val mutable mouseData: uint32
    val mutable flags: uint32
    val mutable time: uint32
    val mutable dwExtraInfo: IntPtr
end


[<DllImport("user32.dll")>]
extern bool GetKeyboardState(byte[] lpKeyState)

[<DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
extern System.UInt32 MapVirtualKey(System.UInt32 uCode, System.UInt32 uMapType)

[<DllImport("user32.dll")>]
extern bool GetCursorPos([<Out>] System.Drawing.Point& lpPoint)

[<DllImport("user32.dll", SetLastError = true)>]
extern void mouse_event(uint32 dwFlags, uint32 dx, uint32 dy, uint32 dwData, IntPtr dwExtraInfo)

[<DllImport("user32.dll", SetLastError = true)>]
extern void keybd_event(byte bVk, byte bScan, uint32 dwFlags, IntPtr dwExtraInfo)

let noASCIkeys = Dictionary<int, string>()
noASCIkeys.Add(220, "\\")
noASCIkeys.Add(189, "-")
noASCIkeys.Add(187, "=")
noASCIkeys.Add(219, "[")
noASCIkeys.Add(221, "]")
noASCIkeys.Add(186, ";")
noASCIkeys.Add(222, "'")
noASCIkeys.Add(188, ",")
noASCIkeys.Add(190, ".")
noASCIkeys.Add(191, "/")
noASCIkeys.Add(192, "`")
noASCIkeys.Add(0x20, " ") // spacja

let sysKeys = Dictionary<int, string>()
sysKeys.Add(0x08, "Backspace")
sysKeys.Add(0x09, "Tab")
sysKeys.Add(0x0D,"Enter")
sysKeys.Add(160, "lShift")
sysKeys.Add(161, "rShift")
sysKeys.Add(162, "lCtrl")
sysKeys.Add(163, "rCtrl")
sysKeys.Add(0x14, "Caps Lock")
sysKeys.Add(0x1B, "Esc")
sysKeys.Add(0x2E, "Delete")
sysKeys.Add(91, "Win")
sysKeys.Add(93, "Options")
sysKeys.Add(164, "Alt")
sysKeys.Add(165, "Alt Gr")
sysKeys.Add(0x70, "F1")
sysKeys.Add(0x71, "F2")
sysKeys.Add(0x72, "F3")
sysKeys.Add(0x73, "F4")
sysKeys.Add(0x74, "F5")
sysKeys.Add(0x75, "F6")
sysKeys.Add(0x76, "F7")
sysKeys.Add(0x77, "F8")
sysKeys.Add(0x78, "F9")
sysKeys.Add(0x79, "F10")
sysKeys.Add(0x7B, "F12")
sysKeys.Add(0x2C, "PrtSc")

// stałe do odtwarzania nagranych zdarzen
let KEYEVENTF_KEYUP = 0x0002u
let MOUSEEVENTF_LEFTDOWN = 0x0002u
let MOUSEEVENTF_LEFTUP = 0x0004u
let MOUSEEVENTF_RIGHTDOWN = 0x0008u
let MOUSEEVENTF_RIGHTUP = 0x0010u
let MOUSEEVENTF_WHEEL = 0x0800u
let MOUSEEVENTF_MIDDLEDOWN = 0x0020
let MOUSEEVENTF_MIDDLEUP = 0x0040

// Stałe dla globalnego hooka myszy
let WH_MOUSE_LL = 14

// Stałe dla komunikatów związanych z naciśnięciem i zwolnieniem przycisków myszy
let WM_LBUTTONDOWN = 0x0201   // Lewy przycisk myszy - naciśnięcie
let WM_LBUTTONUP = 0x0202     // Lewy przycisk myszy - zwolnienie
let WM_RBUTTONDOWN = 0x0204   // Prawy przycisk myszy - naciśnięcie
let WM_RBUTTONUP = 0x0205     // Prawy przycisk myszy - zwolnienie
let WM_MBUTTONDOWN = 0x0207   // Środkowy przycisk myszy - naciśnięcie
let WM_MBUTTONUP = 0x0208     // Środkowy przycisk myszy - zwolnienie
let WM_MOUSEWHEEL = 0x020A
let WM_MOUSEHWHEEL = 0x020E

// stałe dla globalnego hooka klawiatury
let WH_KEYBOARD_LL = 13
let WM_KEYDOWN = 0x0100
let WM_SYSKEYDOWN = 0x0104
let WM_KEYUP = 0x0101
let WM_SYSKEYUP = 0x0105

[<DllImport("user32.dll")>]
extern bool RegisterHotKey(IntPtr hWnd, int id, uint32 fsModifiers, uint32 vk)

[<DllImport("user32.dll")>]
extern bool UnregisterHotKey(IntPtr hWnd, int id)

[<Literal>]
let MOD_ALT = 0x0001
[<Literal>]
let MOD_CONTROL = 0x0002
[<Literal>]
let MOD_SHIFT = 0x0004
[<Literal>]
let WM_HOTKEY = 0x0312

let mutable mouseHookId = IntPtr.Zero
let mutable keyboardHookId = IntPtr.Zero
let mutable stopwatch = Stopwatch.StartNew()  // Rozpoczynamy pomiar czasu

// Definicja delegata dla klawiatury
type LowLevelKeyboardProc = delegate of int * IntPtr * IntPtr -> IntPtr


