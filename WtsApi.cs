using System;
using System.Runtime.InteropServices;
using System.Text;

namespace RDPManager
{
    internal static class WtsApi
    {
        internal const int WTS_CURRENT_SERVER_HANDLE = 0;
        internal const int WTS_CURRENT_SESSION = -1;

        [Flags]
        internal enum HotkeyModifiers : ushort
        {
            None = 0x0,
            Alt = 0x1,
            Ctrl = 0x2,
            Shift = 0x4,
            Win = 0x8
        }

        internal enum WTS_INFO_CLASS
        {
            WTSInitialProgram,
            WTSApplicationName,
            WTSWorkingDirectory,
            WTSOEMId,
            WTSSessionId,
            WTSUserName,
            WTSWinStationName,
            WTSDomainName,
            WTSConnectState,
            WTSClientBuildNumber,
            WTSClientName,
            WTSClientDirectory,
            WTSClientProductId,
            WTSClientHardwareId,
            WTSClientAddress,
            WTSClientDisplay,
            WTSClientProtocolType
        }

        internal enum WTS_CONNECTSTATE_CLASS
        {
            WTSActive,
            WTSConnected,
            WTSConnectQuery,
            WTSShadow,
            WTSDisconnected,
            WTSIdle,
            WTSListen,
            WTSReset,
            WTSDown,
            WTSInit
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct WTS_SESSION_INFO
        {
            public int SessionID;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pWinStationName;
            public WTS_CONNECTSTATE_CLASS State;
        }

        [DllImport("wtsapi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern IntPtr WTSOpenServer(string pServerName);

        [DllImport("wtsapi32.dll", SetLastError = true)]
        internal static extern void WTSCloseServer(IntPtr hServer);

        [DllImport("wtsapi32.dll", SetLastError = true)]
        internal static extern bool WTSEnumerateSessions(
            IntPtr hServer,
            int Reserved,
            int Version,
            out IntPtr ppSessionInfo,
            out int pCount);

        [DllImport("wtsapi32.dll", SetLastError = true)]
        internal static extern void WTSFreeMemory(IntPtr pMemory);

        [DllImport("wtsapi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern bool WTSQuerySessionInformation(
            IntPtr hServer,
            int sessionId,
            WTS_INFO_CLASS wtsInfoClass,
            out IntPtr ppBuffer,
            out int pBytesReturned);

        [DllImport("wtsapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool WTSStartRemoteControlSession(
            string pTargetServerName,
            int TargetLogonId,
            byte HotkeyVk,
            HotkeyModifiers HotkeyModifiers);

        [DllImport("wtsapi32.dll", SetLastError = true)]
        internal static extern bool WTSStopRemoteControlSession(int LogonId);

        [DllImport("wtsapi32.dll", SetLastError = true)]
        internal static extern bool WTSLogoffSession(IntPtr hServer, int SessionId, bool bWait);

        [DllImport("wtsapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool WTSSendMessage(
            IntPtr hServer,
            int SessionId,
            string pTitle,
            int TitleLength,
            string pMessage,
            int MessageLength,
            uint Style,
            int Timeout,
            out int pResponse,
            bool bWait);

        internal static string PtrToStringAndFree(IntPtr buffer)
        {
            if (buffer == IntPtr.Zero)
            {
                return string.Empty;
            }

            try
            {
                return Marshal.PtrToStringAuto(buffer) ?? string.Empty;
            }
            finally
            {
                WTSFreeMemory(buffer);
            }
        }

        internal static WTS_CONNECTSTATE_CLASS PtrToConnectStateAndFree(IntPtr buffer)
        {
            if (buffer == IntPtr.Zero)
            {
                return WTS_CONNECTSTATE_CLASS.WTSDown;
            }

            try
            {
                int raw = Marshal.ReadInt32(buffer);
                return (WTS_CONNECTSTATE_CLASS)raw;
            }
            finally
            {
                WTSFreeMemory(buffer);
            }
        }

        internal static string GetWin32ErrorMessage(int error)
        {
            return new System.ComponentModel.Win32Exception(error).Message;
        }

        internal static string GetStateDisplayName(WTS_CONNECTSTATE_CLASS state)
        {
            switch (state)
            {
                case WTS_CONNECTSTATE_CLASS.WTSActive:
                    return "Active";
                case WTS_CONNECTSTATE_CLASS.WTSConnected:
                    return "Connected";
                case WTS_CONNECTSTATE_CLASS.WTSDisconnected:
                    return "Disconnected";
                case WTS_CONNECTSTATE_CLASS.WTSIdle:
                    return "Idle";
                case WTS_CONNECTSTATE_CLASS.WTSShadow:
                    return "Shadow";
                case WTS_CONNECTSTATE_CLASS.WTSListen:
                    return "Listen";
                case WTS_CONNECTSTATE_CLASS.WTSReset:
                    return "Reset";
                case WTS_CONNECTSTATE_CLASS.WTSDown:
                    return "Down";
                case WTS_CONNECTSTATE_CLASS.WTSInit:
                    return "Init";
                case WTS_CONNECTSTATE_CLASS.WTSConnectQuery:
                    return "ConnectQuery";
                default:
                    return state.ToString();
            }
        }
    }
}
