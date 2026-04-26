using System;
using System.Runtime.InteropServices;

namespace MarvelHeroes.DpsMeter.Interop;

/// <summary>
/// Minimal P/Invoke surface for the standalone DPS overlay — just the bits needed to make a
/// borderless WPF window non-activating (so a click on the overlay never steals focus from
/// Marvel Heroes' fullscreen window) and to extend its window-style flags.
///
/// <para>
/// This is a hand-trimmed subset of the parent comporator's <c>User32</c> wrapper — we ship
/// only the constants + entry points actually called from <c>DpsOverlayWindow</c>, so the
/// standalone exe doesn't drag in unrelated Win32 surfaces (hotkey registration, DPI helpers,
/// process-id lookups, screen-capture exclusion).  Add to this file as new overlay features
/// require additional Win32 calls; do NOT pull the whole parent file in wholesale.
/// </para>
/// </summary>
internal static class User32
{
    /// <summary>WM_MOUSEACTIVATE — sent before WM_*BUTTONDOWN so we can intercept and return
    /// MA_NOACTIVATE, telling Windows "process the click but do not bring me to foreground".
    /// Without this, every drag of the overlay would yank focus away from the game.</summary>
    public const int WM_MOUSEACTIVATE = 0x0021;

    /// <summary>Return value for WM_MOUSEACTIVATE: hit-test passes through and the click is
    /// processed but the window is NOT activated and NOT brought to the foreground.</summary>
    public const int MA_NOACTIVATE = 3;

    /// <summary>Index for GetWindowLongPtr / SetWindowLongPtr — extended window style bitfield.</summary>
    public const int GWL_EXSTYLE = -20;

    /// <summary>WS_EX_NOACTIVATE — window does not get foreground activation when clicked.
    /// Combined with the WM_MOUSEACTIVATE hook gives us the "draggable but non-activating"
    /// behaviour expected of a HUD overlay.</summary>
    public const nint WS_EX_NOACTIVATE = 0x08000000;

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW", SetLastError = true)]
    public static extern nint GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
    public static extern nint SetWindowLongPtr(IntPtr hWnd, int nIndex, nint dwNewLong);

    /// <summary>Handle of the window currently in the foreground (the one receiving keyboard
    /// input).  Used by the foreground watcher to auto-hide the overlay while Marvel Heroes
    /// isn't the active window.  Returns <c>IntPtr.Zero</c> during transient states (e.g. an
    /// app is in the middle of activating) — callers should treat zero as "indeterminate" and
    /// keep the previous decision rather than flipping visibility on a glitch.</summary>
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetForegroundWindow();

    /// <summary>Maps an HWND to the OS process id of its owning thread.  Combined with
    /// <see cref="System.Diagnostics.Process.GetProcessById"/> this gives us the process name,
    /// which is how the foreground watcher decides whether the active window is the game,
    /// our own overlay (right-click menu open), or something else entirely.</summary>
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    // ── Low-level mouse hook (for context-menu outside-click dismissal) ───────────────────────
    // WS_EX_NOACTIVATE breaks WPF's standard ContextMenu / Popup outside-click dismissal because
    // the Popup relies on Win32 SetCapture, which only holds while a window in our process has
    // foreground.  An overlay that can't activate can't keep capture, so when the user clicks
    // anywhere outside the popup (especially on the game's window in another process) the popup
    // never receives the dismissal signal and stays open forever.
    //
    // Solution: install a global low-level mouse hook (WH_MOUSE_LL) only while the menu is
    // open.  The hook fires for every mouse-down system-wide, regardless of which window
    // owns the click.  We just check if the click HWND is the popup HWND; if not, close the
    // menu.  Hook is uninstalled the moment the menu closes — overhead is bounded to the
    // brief lifetime of a user-initiated context menu.
    public const int WH_MOUSE_LL = 14;
    public const int WM_LBUTTONDOWN = 0x0201;
    public const int WM_RBUTTONDOWN = 0x0204;
    public const int WM_MBUTTONDOWN = 0x0207;
    public const int WM_NCLBUTTONDOWN = 0x00A1;
    public const int WM_NCRBUTTONDOWN = 0x00A4;

    /// <summary>Hook proc signature for WH_MOUSE_LL.  Must call <see cref="CallNextHookEx"/>
    /// to allow the click to reach the actual target window — DO NOT suppress mouse events
    /// from the hook (would break literally every click in the system while the menu is open).</summary>
    public delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr SetWindowsHookExW(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    /// <summary>Returns the topmost HWND at <paramref name="point"/> (screen coordinates).
    /// Used by the menu-dismissal hook to decide whether a click landed inside the popup or
    /// somewhere else (game window, other overlay, desktop).</summary>
    [DllImport("user32.dll")]
    public static extern IntPtr WindowFromPoint(POINT point);

    /// <summary>Returns the root ancestor of <paramref name="hwnd"/>.  Combined with
    /// <see cref="WindowFromPoint"/> this lets us walk a child HWND back up to its top-level
    /// window for popup-membership testing — relevant for tooltip / sub-popup HWNDs that
    /// belong to the menu but aren't the popup HWND itself.</summary>
    [DllImport("user32.dll")]
    public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

    /// <summary>GA_ROOT — top-level ancestor (skipping owner relationships).</summary>
    public const uint GA_ROOT = 2;

    /// <summary>Module handle for the calling process — required by <see cref="SetWindowsHookExW"/>
    /// when installing a global hook.  For WH_MOUSE_LL specifically the hook isn't actually injected
    /// into other processes, but the API still requires a valid hModule (any module in our process
    /// works — we use the main exe via <c>GetModuleHandleW(null)</c>).</summary>
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr GetModuleHandleW(string? lpModuleName);
}
