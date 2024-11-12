using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

[ComImport]
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
internal class MMDeviceEnumerator
{
}

[Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator
{
    int EnumAudioEndpoints(EDataFlow dataFlow, DeviceState stateMask, out IMMDeviceCollection devices);
    int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
}

[Guid("D666063F-1587-4E43-81F1-B948E807363F"),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice
{
    int Activate(ref Guid iid, ClsCtx clsCtx, IntPtr activationParams, out IAudioSessionManager2 sessionManager);
}

[Guid("0BD7A1BE-7A1A-44DB-8397-C0A4A1F9CB02"),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceCollection
{
    int GetCount(out uint count);
    int Item(uint index, out IMMDevice device);
}

[Guid("bfb7ff88-7239-4fc9-8fa2-07c950be9c6d"),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioSessionManager2
{
    int GetSessionEnumerator(out IAudioSessionEnumerator sessionEnum);
}

[Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioSessionEnumerator
{
    int GetCount(out int sessionCount);
    int GetSession(int sessionIndex, out IAudioSessionControl2 sessionControl);
}

[Guid("bfb7ff88-7239-4fc9-8fa2-07c950be9c6d"),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioSessionControl2
{
    int GetProcessId(out int processId);
    int QueryInterface(ref Guid riid, out ISimpleAudioVolume simpleVolume);
}

[Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8"),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface ISimpleAudioVolume
{
    int SetMasterVolume(float fLevel, ref Guid EventContext);
    int GetMasterVolume(out float pfLevel);
    int SetMute(bool bMute, ref Guid EventContext);
    int GetMute(out bool pbMute);
}

enum EDataFlow
{
    eRender,
    eCapture,
    eAll
}

enum ERole
{
    eConsole,
    eMultimedia,
    eCommunications
}

enum ClsCtx
{
    Inproc = 0x1
}

[Flags]
enum DeviceState
{
    Active = 0x00000001,
}

namespace WinAudioMixerLib
{
    public class VolumeMixer
    {
        private static readonly Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
        private static readonly Guid IID_ISimpleAudioVolume = typeof(ISimpleAudioVolume).GUID;
        private static readonly Guid EVENT_CONTEXT = Guid.NewGuid();

        public static void SetApplicationVolume(int processId, float volume)
        {
            var deviceEnumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
            deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out var device);

            var iidAudioSessionManager2 = IID_IAudioSessionManager2;
            var iidSimpleAudioVolume    = IID_ISimpleAudioVolume;
            var eventContext            = EVENT_CONTEXT;

            device.Activate(ref iidAudioSessionManager2, ClsCtx.Inproc, IntPtr.Zero, out var sessionManager);
            sessionManager.GetSessionEnumerator(out var sessionEnumerator);
            sessionEnumerator.GetCount(out var count);

            for (int i = 0; i < count; i++)
            {
                sessionEnumerator.GetSession(i, out var sessionControl);
                sessionControl.GetProcessId(out int sessionId);

                if (sessionId == processId)
                {
                    sessionControl.QueryInterface(ref iidSimpleAudioVolume, out var simpleVolume);
                    simpleVolume.SetMasterVolume(volume, ref eventContext);
                    Marshal.ReleaseComObject(simpleVolume);
                }

                Marshal.ReleaseComObject(sessionControl);
            }

            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(sessionManager);
            Marshal.ReleaseComObject(device);
            Marshal.ReleaseComObject(deviceEnumerator);
        }

        public static float? GetApplicationVolume(int processId)
        {
            var deviceEnumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
            deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out var device);

            var iidAudioSessionManager2 = IID_IAudioSessionManager2;
            var iidSimpleAudioVolume = IID_ISimpleAudioVolume;
            var eventContext = EVENT_CONTEXT;

            device.Activate(ref iidAudioSessionManager2, ClsCtx.Inproc, IntPtr.Zero, out var sessionManager);
            sessionManager.GetSessionEnumerator(out var sessionEnumerator);
            sessionEnumerator.GetCount(out var count);

            for (int i = 0; i < count; i++)
            {
                sessionEnumerator.GetSession(i, out var sessionControl);
                sessionControl.GetProcessId(out int sessionId);

                if (sessionId == processId)
                {
                    sessionControl.QueryInterface(ref iidSimpleAudioVolume, out var simpleVolume);
                    simpleVolume.GetMasterVolume(out float volume);
                    Marshal.ReleaseComObject(simpleVolume);
                    return volume;
                }

                Marshal.ReleaseComObject(sessionControl);
            }

            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(sessionManager);
            Marshal.ReleaseComObject(device);
            Marshal.ReleaseComObject(deviceEnumerator);

            return null;
        }
    }

}
