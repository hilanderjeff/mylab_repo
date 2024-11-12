using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace WinAudioMixerLib
{
    public class WindowsAudioMixer : IDisposable
    {
        #region Native COM Interfaces
        [ComImport]
        [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        private class MMDeviceEnumerator { }

        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceEnumerator
        {
            [PreserveSig]
            int EnumAudioEndpoints(EDataFlow dataFlow, DeviceState stateMask, out IMMDeviceCollection deviceCollection);

            [PreserveSig]
            int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
        }

        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDevice
        {
            [PreserveSig]
            int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        }

        [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioSessionEnumerator
        {
            [PreserveSig]
            int GetCount(out int sessionCount);

            [PreserveSig]
            int GetSession(int sessionCount, out IAudioSessionControl session);
        }

        [Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioSessionControl
        {
            int GetState(out AudioSessionState state);
            int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string displayName);
            // ... other methods not used in this implementation
        }

        [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ISimpleAudioVolume
        {
            [PreserveSig]
            int SetMasterVolume(float level, ref Guid eventContext);

            [PreserveSig]
            int GetMasterVolume(out float level);

            [PreserveSig]
            int SetMute(bool mute, ref Guid eventContext);

            [PreserveSig]
            int GetMute(out bool mute);
        }

        [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioSessionManager2
        {
            int GetSessionEnumerator(out IAudioSessionEnumerator sessionEnum);
            // ... other methods not used in this implementation
        }

        [Guid("BFA971F1-4D5E-40BB-935E-967039BFBEE4"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioSessionControl2 : IAudioSessionControl
        {
            // Inherit base interface methods

            [PreserveSig]
            int GetSessionIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string sessionId);

            [PreserveSig]
            int GetSessionInstanceIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string instanceId);

            [PreserveSig]
            int GetProcessId(out uint processId);
        }

        private enum EDataFlow
        {
            eRender = 0,
            eCapture = 1,
            eAll = 2
        }

        private enum ERole
        {
            eConsole = 0,
            eMultimedia = 1,
            eCommunications = 2
        }

        private enum DeviceState
        {
            DEVICE_STATE_ACTIVE = 0x00000001,
            DEVICE_STATE_DISABLED = 0x00000002,
            DEVICE_STATE_NOTPRESENT = 0x00000004,
            DEVICE_STATE_UNPLUGGED = 0x00000008,
            DEVICE_STATEMASK_ALL = 0x0000000F
        }

        private enum AudioSessionState
        {
            Inactive = 0,
            Active = 1,
            Expired = 2
        }

        [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
        private interface IMMDeviceCollection
        {
            [PreserveSig]
            int GetCount(out uint count);
        }
        #endregion

        private IMMDeviceEnumerator deviceEnumerator;
        private IMMDevice defaultDevice;
        private IAudioSessionManager2 sessionManager;

        public WindowsAudioMixer()
        {
            Initialize();
        }

        private void Initialize()
        {
            deviceEnumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
            deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out defaultDevice);

            Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
            defaultDevice.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out object sessionManagerObj);
            sessionManager = (IAudioSessionManager2)sessionManagerObj;
        }

        public bool SetApplicationVolume(string processName, float volume)
        {
            var session = GetSessionControl(processName);
            if (session == null)
                return false;

            Guid guid = Guid.Empty;
            var audioVolume = session as ISimpleAudioVolume;
            audioVolume?.SetMasterVolume(volume / 100f, ref guid);

            Marshal.ReleaseComObject(session);
            return true;
        }

        public float? GetApplicationVolume(string processName)
        {
            var session = GetSessionControl(processName);
            if (session == null)
                return null;

            var audioVolume = session as ISimpleAudioVolume;
            audioVolume.GetMasterVolume(out float volume);

            Marshal.ReleaseComObject(session);
            return volume * 100f;
        }

        public bool SetApplicationMute(string processName, bool mute)
        {
            var session = GetSessionControl(processName);
            if (session == null)
                return false;

            Guid guid = Guid.Empty;
            var audioVolume = session as ISimpleAudioVolume;
            audioVolume?.SetMute(mute, ref guid);

            Marshal.ReleaseComObject(session);
            return true;
        }

        private object GetSessionControl(string processName)
        {
            sessionManager.GetSessionEnumerator(out IAudioSessionEnumerator sessionEnumerator);
            sessionEnumerator.GetCount(out int sessionCount);

            for (int i = 0; i < sessionCount; i++)
            {
                sessionEnumerator.GetSession(i, out IAudioSessionControl sessionControl);
                var sessionControl2 = sessionControl as IAudioSessionControl2;

                if (sessionControl2 != null)
                {
                    sessionControl2.GetProcessId(out uint processId);
                    var process = System.Diagnostics.Process.GetProcessById((int)processId);

                    if (process.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase))
                    {
                        Marshal.ReleaseComObject(sessionEnumerator);
                        return sessionControl;
                    }

                    Marshal.ReleaseComObject(sessionControl);
                }
            }

            Marshal.ReleaseComObject(sessionEnumerator);
            return null;
        }

        public void Dispose()
        {
            if (sessionManager != null)
            {
                Marshal.ReleaseComObject(sessionManager);
                sessionManager = null;
            }
            if (defaultDevice != null)
            {
                Marshal.ReleaseComObject(defaultDevice);
                defaultDevice = null;
            }
            if (deviceEnumerator != null)
            {
                Marshal.ReleaseComObject(deviceEnumerator);
                deviceEnumerator = null;
            }
        }
    }
}
