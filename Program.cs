using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
//Thanks to CSharpFritz and EngstromJimmy for their gists, snippets, and thoughts.

namespace PowerPointToOSC
{
    class Program
    {
        private static Application ppt = new Microsoft.Office.Interop.PowerPoint.Application();
        private static OscLocal OSC;
        static async Task Main(string[] args)
        {
            Console.Write("Connecting to PowerPoint...");
            ppt.SlideShowNextSlide += App_SlideShowNextSlide;
            ppt.SlideShowNextClick += App_SlideShowNextClick;
            Console.WriteLine("Connected. Ready to send OSC commands.");
            OSC = new OscLocal();
            Console.ReadLine();
        }

        // Future use: multiple OSC command, associated to Animations in a single slide
        async static void App_SlideShowNextClick(SlideShowWindow Wn, Effect effect)
        {
            if (Wn != null)
            {
                if (effect != null)
                {
                    Console.WriteLine($"  Next Effect: {effect.DisplayName}");
                }
                
            }
        }

        async static void App_SlideShowNextSlide(SlideShowWindow Wn)
        {
            if (Wn != null)
            {
                Console.WriteLine($"Moved to Slide Number {Wn.View.Slide.SlideNumber}");
                //Text starts at Index 2 ¯\_(ツ)_/¯
                var note = String.Empty;
                try { note = Wn.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text; }
                catch { /*no notes*/ }

                bool oscCommandHandled = false;


                var notereader = new StringReader(note);
                string line;
                while ((line = notereader.ReadLine()) != null)
                {
                    if (line.StartsWith("OSC:"))
                    {
                        line = line.Substring(4).Trim();

                        if (!oscCommandHandled)
                        {
                            // Console.WriteLine($"Found OSC command in slide: \"{line}\"");
                            try
                            {
                                oscCommandHandled = OSC.sendOSC(line);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"  ERROR: {ex.Message.ToString()}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"  WARNING: Multiple OSC commands found. I used the first and have ignored \"{line}\"");
                        }
                    }

                    if (line.StartsWith("OSCHOST:"))
                    {
                        OSC.OscHost = line.Substring(8).Trim();
                        Console.WriteLine($"Setting the OSC host to \"{OSC.OscHost}\"");
                    }
                }
            }
        }

    }
}