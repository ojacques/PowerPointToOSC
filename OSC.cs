using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Sockets;

namespace PowerPointToOSC
{
    public class OscLocal
    {
        // Default OSC Host, works with DMXIS
        private string oscHost = "127.0.0.1 8000";

        public OscLocal() { }
      
        public string OscHost
        {
            get { return oscHost; }
            set
            {
                oscHost = value;
                Console.WriteLine($"OSC Host set to {value}");
            }
        }

        public bool sendOSC(string oscString)
        {
            Console.WriteLine($"  Sending OSC command to {oscHost}: {oscString}");
            // Current implementation: leverage sendosc.exe - https://github.com/yoggy/sendosc
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "sendosc.exe";
            startInfo.Arguments = oscHost + " " + oscString;
            process.StartInfo = startInfo;
            process.Start();
            return true;
        }
    }
}
