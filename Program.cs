using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
//Thanks to CSharpFritz and EngstromJimmy for their gists, snippets, and thoughts.

namespace PowerPointToOSC
{
    class Program
    {
        private static Application ppt = new Microsoft.Office.Interop.PowerPoint.Application();
        private static OscLocal OSC;
        private static List<string> OSCCommands = new List<string>();    // List of OSC commands in a slide
        private static int slideClick;  // hold the nb of clicks in a given slide
        static async Task Main(string[] args)
        {
            Console.Write("Connecting to PowerPoint...");
            ppt.SlideShowNextSlide += App_SlideShowNextSlide;
            ppt.SlideShowNextClick += App_SlideShowNextClick;
            ppt.SlideShowOnNext += App_SlideShowOnNext;
            ppt.SlideShowOnPrevious += App_SlideShowOnPrevious;
            Console.WriteLine("Connected. Ready to send OSC commands.");
            OSC = new OscLocal();
            Console.ReadLine();
        }

        // Future use
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

        // Clicked next on current slide
        async static void App_SlideShowOnNext(SlideShowWindow Wn)
        {
            if (Wn != null)
            {
                Console.WriteLine($"  Next click on current slide");
                slideClick++;
                if (OSCCommands.Count > slideClick)
                {
                    OSC.sendOSC(OSCCommands[slideClick]);
                }

            }
        }

        // Clicked previous on current slide
        async static void App_SlideShowOnPrevious(SlideShowWindow Wn)
        {
            if (Wn != null)
            {
                Console.WriteLine($"  Prev click on current slide");
                if (slideClick>0) slideClick--;
                if (OSCCommands.Count > slideClick)
                {
                    OSC.sendOSC(OSCCommands[slideClick]);
                }

            }
        }

        // Next slide
        async static void App_SlideShowNextSlide(SlideShowWindow Wn)
        {
            if (Wn != null)
            {
                Console.WriteLine($"Moved to Slide Number {Wn.View.Slide.SlideNumber}");
                //Text starts at Index 2 ¯\_(ツ)_/¯
                var note = String.Empty;
                OSCCommands.Clear();    // Start with a fresh list of commands on new slide
                slideClick = 0;
                try { note = Wn.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text; }
                catch { /*no notes*/ }

                bool oscCommandHandled = false; // Only handle the 1st command on slide change

                var notereader = new StringReader(note);
                string line;
                while ((line = notereader.ReadLine()) != null)
                {
                    if (line.StartsWith("OSC:"))
                    {
                        line = line.Substring(4).Trim();
                        OSCCommands.Add(line); // Add OSC command to the list of commands in that slide
                        if (!oscCommandHandled)
                        {
                            try
                            {
                                oscCommandHandled = OSC.sendOSC(line);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"  ERROR: {ex.Message.ToString()}");
                            }
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