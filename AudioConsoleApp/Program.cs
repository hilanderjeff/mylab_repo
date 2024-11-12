// See https://aka.ms/new-console-template for more information
using System.Diagnostics;
using WinAudioMixerLib;

class Program
{
    static void Main()
    {
        int processId = Process.GetProcessesByName("ConnectAgent").FirstOrDefault()?.Id ?? -1;

        if (processId != -1)
        {
            VolumeMixer.SetApplicationVolume(processId, 0.5f);
            // Set volume to 50%
            float? currentVolume = VolumeMixer.GetApplicationVolume(processId);
            Console.WriteLine($"Current Volume: {currentVolume}");
        }
        else
        {
            Console.WriteLine("Application not found.");
        }
    }
}

