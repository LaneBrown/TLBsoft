using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows.Forms;

namespace SleepingPill
{
    /// <summary>
    /// This app was developed to help fix a Windows 10 computer that would
    /// start the screen saver reliably but would not go to sleep reliably.
    /// It assumes that PowerCfg /requestsoverride has been properly done by the user.
    /// 
    /// This program monitors screen saver processes (.SCR) and measures the running time.
    /// When a screen saver has been running long enough, it starts invoking POWERCFG.EXE
    /// and looking at the SYSTEM requests. These requests, if not overridden, tell the OS
    /// to keep the computer awake. If there are no requests for a time, the computer
    /// is told to go to sleep.
    /// </summary>
    /// <remarks>Invoking PowerCfg.exe requires the app to be run as administrator -AND-
    /// the build configuration must match the operating system. Build two release versions:
    /// One for the x64 target and a second for the x86 target (AnyCPU won't do).
    /// Rename the x86 version from SleepingPill.exe to SleepingPill_x86.exe. Make both
    /// versions available to users so they can choose the one that matches their OS.</remarks>
    class Program
    {
        /**************************** IMPORTANT *****************************************
        * When debugging, run VisualStudio as administrator so POWERCFG calls will work *
        * -OR- edit the app.manifest and change the requestExecutionLevel node to       *
        * <requestedExecutionLevel  level="requireAdministrator" uiAccess="false" />    *
        ********************************************************************************/
        private const string scrExtension = ".SCR";            // Extension of screen saver processes; should be upper-case
        private static string localUserAppDataPath = null;     // AppData folder
        // Argument values after parsing input args
        private static string specificScreenSaver;             // Screen saver process argument
        private static int screenTime;                         // Time argument in milliseconds
        private static int idleTime;                           // Idle argument in milliseconds; this is a kind of de-bounce strategy
        private static bool verbose;                           // Usually only used for initial testing/debugging
        private static int samplePeriod;                       // Sample argument in milliseconds

#if DEBUG
        // P-Invoke stuff is only needed for debugging to make the app appear on top of the screen saver
        // See: https://stackoverflow.com/questions/53026524/how-to-make-c-sharp-console-application-to-always-be-in-front
        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd,
                                        IntPtr hWndInsertAfter,
                                        int X,
                                        int Y,
                                        int cx,
                                        int cy,
                                        uint uFlags);
        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        const uint SWP_NOSIZE = 0x0001, SWP_NOMOVE = 0x0002, SWP_SHOWWINDOW = 0x0040;

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();
#endif

        /// <summary>
        /// Main program for SleepingPill.exe. This program detects how long a screen saver has been running
        /// and puts the computer to sleep if the screen saver has been running more than a certain time
        /// -AND- there are no SYSTEM requests that haven't been overriden
        /// </summary>
        /// <param name="args">See help.txt for argument details</param>
        static void Main(string[] args)
        {
            if (!ParseArguments(args))
            {
                // Arguments didn't parse - assume the user is asking for help
                using (StreamReader reader = new StreamReader(GetResourceStream("help.txt"))) // Embedded resource
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                        Console.WriteLine(line);
                }
                PressAnyKeyWithTimout();
                return;
            }

            // Always report settings even if verbose==false
            Console.WriteLine(string.Format(StringTable.ProcessEquals, string.IsNullOrEmpty(specificScreenSaver) ? StringTable.AnyScreenSaver : specificScreenSaver));
            Console.WriteLine(string.Format(StringTable.TimeEquals, screenTime / 1000));
            Console.WriteLine(string.Format(StringTable.IdleEquals, idleTime / 1000));
            Console.WriteLine(string.Format(StringTable.Verbose, verbose.ToString()));
            Console.WriteLine(string.Format(StringTable.SamplePeriod, samplePeriod / 1000));
            
            // Test to see if the permission levels are right for invoking PowerCfg
            if (!RunningAsAdmin())
            {
                Console.WriteLine(StringTable.PowerCfgProblem);
                Console.WriteLine(StringTable.RunAsAdmin);
                PressAnyKeyWithTimout();
                return;
            }

            // Test to see if the build configuration is right for invoking PowerCfg
            string useOtherVersionMessage = null;
            if (System.Environment.Is64BitOperatingSystem)
            {
                if (IntPtr.Size != 8)
                    useOtherVersionMessage = StringTable.Use64bitVersion;
            }
            else
            {
                if (IntPtr.Size != 4)
                    useOtherVersionMessage = StringTable.Use32bitVersion;
            }
            if (useOtherVersionMessage != null)
            {
                Console.WriteLine(StringTable.PowerCfgProblem);
                Console.WriteLine(useOtherVersionMessage);
                PressAnyKeyWithTimout();
                return;
            }

            // At this point, settings have been parsed and conditions have been vetted
            // Enter the main sample loop
            MonitorLoop();
        }

        /// <summary>
        /// Measures screen saver run time and invokes PowerCfg
        /// Under the right conditions, puts the computer to sleep
        /// </summary>
        private static void MonitorLoop()
        {
            Console.WriteLine(StringTable.Running);
            int timeInScreenSaver = 0;   // Time the screen saver has been running
            int timeNoSystemRequest = 0; // Time with no system requests
            for (; ; )
            {
                Thread.Sleep(samplePeriod);
                if (verbose)
                    Console.Clear();
                // There are two ways to determine if the screen saver is running:
                // 1) Use a specific process name
                // 2) Look for any process name that ends with .SCR
                bool screenSaverRunning;
                Process[] processes;
                if (!string.IsNullOrEmpty(specificScreenSaver))
                {
                    // This method works better if the user wants to be able to run a screen saver on demand
                    // that isn't the one defined in Windows. If the user changes the Windows screen saver,
                    // he will have to change the arguments to the SleepingPill
                    processes = Process.GetProcessesByName(specificScreenSaver);
                    screenSaverRunning = processes != null && processes.Length > 0;
                }
                else
                {
                    // This method works better generally because the user can change the screen saver without
                    // re-doing the arguments to the SleepingPill. However, if he runs ANY screen saver long enough
                    // it will cause the computer to go to sleep.
                    processes = Process.GetProcesses();
                    screenSaverRunning = processes.Any(x => x.ProcessName.ToUpper().EndsWith(scrExtension));
                }
                if (processes != null)
                    foreach (Process process in processes)
                        process.Dispose();

                if (screenSaverRunning)
                {
#if DEBUG
                    // When debugging, it's helpful to see the output so we need to make the
                    // app top-most to show over top of the screen saver. In release mode, we won't.
                    if (verbose && timeInScreenSaver <= samplePeriod)
                    {
                        // Make this app top-most so it shows in front of the screen saver
                        Console.WriteLine(StringTable.MakingTopMost);
                        IntPtr handle = GetConsoleWindow();
                        // See: https://stackoverflow.com/questions/53026524/how-to-make-c-sharp-console-application-to-always-be-in-front
                        SetWindowPos(handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                    }
#endif

                    // Increment timeInScreenSaver
                    timeInScreenSaver += samplePeriod;
                    if (verbose)
                        Console.WriteLine(string.Format(StringTable.ScreenSaverRunTime, timeInScreenSaver / 1000));
                    if (timeInScreenSaver >= screenTime - idleTime)
                    {
                        // Screen saver has been running long enough, start checking system requests
                        bool validRequests = ReadPowercfgRequests();
                        if (validRequests)
                        {
                            // There are valid requests, reset the timer
                            timeNoSystemRequest = 0;
                        }
                        else
                        {
                            // No valid system requests, increment timeNoSystemRequest
                            timeNoSystemRequest += samplePeriod;
                            if (timeNoSystemRequest >= idleTime)
                            {
                                if (verbose)
                                {
                                    Console.WriteLine(StringTable.PuttingComputerToSleep);
                                    Thread.Sleep(100);
                                }
                                
                                // If running from Task Scheduler, there will be no UI so the only way
                                // to report that something was done is through a mechanism like this
                                IncrementSleepCountFile();

                                // Prevent going to sleep until the whole sequence of events happens again
                                timeInScreenSaver = 0;
                                timeNoSystemRequest = 0;
                                // Make the PC go to sleep
                                Application.SetSuspendState(PowerState.Suspend, true, false);
                            }
                        }
                    }
                    else
                    {
                        // Screen saver hasn't been on long enough yet
                        timeNoSystemRequest = 0;
                    }
                }
                else
                {
                    // Screen saver isn't running
                    if (verbose)
                        Console.WriteLine(string.Format(StringTable.ScreenSaverNotRunning));
                    timeInScreenSaver = 0;
                    timeNoSystemRequest = 0;
                }
            }
        }

        /// <summary>
        /// Returns true if running as administrator
        /// </summary>
        /// <returns></returns>
        private static bool RunningAsAdmin()
        {
            try
            {
                using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
                {
                    return (new WindowsPrincipal(identity)).IsInRole(WindowsBuiltInRole.Administrator);
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Parse arguments, return true if successful, false if not
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private static bool ParseArguments(string[] args)
        {
            string crlf = new string(new char[] { '\r', '\n' });
            // Argument keys, if these are changed then help.txt needs to change as well
            const string screenSaverProcessKey = "PROCESS";// Argument key for the screen saver process to monitor
            const string screenTimeKey = "TIME";           // Argument key for the screen saver running time
            const string idleTimeKey = "IDLE";             // Argument key for the idle time
            const string verboseKey = "VERBOSE";           // Argument key for the verbose switch
            const string samplePeriodKey = "SAMPLE";       // Argument key for the sample period
            // Argument defaults, minimums and maximums
            string defaultScreenSaver = string.Empty;      // All/any screen saver
            const int defaultScreenTime = 180;             // Seconds
            const int defaultIdleTime = 20;                // Seconds
            const int defaultSamplePeriod = 5;             // Seconds
            const int minimumTime = 30;                    // Seconds, applies to both screenTime and idleTime
            const int maximumTime = 1800;                  // Seconds, applies to both screenTime and idleTime
            const int minimumSamplePeriod = 1;             // Seconds
            const int maximumSamplePeriod = 60;            // Seconds

            // Initialize with defaults in case there are no args to parse
            specificScreenSaver = defaultScreenSaver;
            screenTime = defaultScreenTime * 1000;
            idleTime = defaultIdleTime * 1000;
            verbose = false;
            samplePeriod = defaultSamplePeriod * 1000;
            
            // Parse the args
            if (args != null && args.Length > 0)
            {
                // Parse each argument
                foreach (string s in args)
                {
                    // Remove CR-LF sequences, if any
                    string arg = s.Replace(crlf, string.Empty);
                    if (arg.Equals(string.Empty))
                        continue; // Skip empty arguments
                    bool argParsed = false;
                    string argValue = null;
                    if ((argValue = Extract(screenSaverProcessKey, arg, defaultScreenSaver)) != null)
                    {
                        // Parse the screen saver process name, require the extension to be correct
                        if (argValue.Length > 4 && argValue.ToUpper().EndsWith(scrExtension))
                        {
                            specificScreenSaver = argValue;
                            argParsed = true;
                            continue;
                        }
                        else if (argValue == defaultScreenSaver)
                        {
                            argParsed = true;
                            continue;
                        }

                    }
                    if ((argValue = Extract(screenTimeKey, arg, defaultScreenTime.ToString())) != null)
                    {
                        // Parse screenTime
                        int parsedScreenTime;
                        if (int.TryParse(argValue, out parsedScreenTime))
                        {
                            parsedScreenTime = Math.Min(maximumTime, Math.Max(minimumTime, parsedScreenTime));
                            screenTime = parsedScreenTime * 1000; // Convert to milliseconds
                            argParsed = true;
                            continue;
                        }
                    }
                    if ((argValue = Extract(idleTimeKey, arg, defaultIdleTime.ToString())) != null)
                    {
                        // Parse idleTime
                        int parsedIdleTime;
                        if (int.TryParse(argValue, out parsedIdleTime))
                        {
                            parsedIdleTime = Math.Min(maximumTime, Math.Max(minimumTime, parsedIdleTime));
                            idleTime = parsedIdleTime * 1000; // Convert to milliseconds
                            argParsed = true;
                            continue;
                        }
                    }
                    if ((argValue = Extract(verboseKey, arg, true.ToString())) != null)
                    {
                        // Parse Verbose
                        bool parsedVerbose;
                        if (bool.TryParse(argValue, out parsedVerbose))
                        {
                            verbose = parsedVerbose;
                            argParsed = true;
                            continue;
                        }
                    }
                    if ((argValue = Extract(samplePeriodKey, arg, defaultSamplePeriod.ToString())) != null)
                    {
                        // Parse samplePeriod
                        int parsedSamplePeriod;
                        if (int.TryParse(argValue, out parsedSamplePeriod))
                        {
                            parsedSamplePeriod = Math.Min(maximumSamplePeriod, Math.Max(minimumSamplePeriod, parsedSamplePeriod));
                            samplePeriod = parsedSamplePeriod * 1000; // Convert to milliseconds
                            argParsed = true;
                            continue;
                        }
                    }
                    if (!argParsed)
                        return false; // Anything that doesn't parse will be lumped in with ?, HELP and HLP which should trigger help 
                }
            }
            // The idle time should be less than or equal to the screen time
            idleTime = Math.Min(screenTime, idleTime);
            
            return true;
        }

        /// <summary>
        /// Extracts the value of the argument as a string
        /// </summary>
        /// <param name="key">Possible prefix</param>
        /// <param name="arg">E.g. /Process="Mystify.scr", -time:180, idle=20, etc.</param>
        /// <param name="defaultValue">Value to use if none supplied</param>
        /// <returns></returns>
        private static string Extract(string key, string arg, string defaultValue)
        {
            int index = arg.ToUpper().IndexOf(key.ToUpper());
            if (index == 0 || (index == 1 && (arg.StartsWith("/") || arg.StartsWith(":"))))
            {
                // Remove the key
                arg = arg.Remove(0, key.Length + index);
                if (arg.StartsWith("=") || arg.StartsWith(":"))
                {
                    // Strip = or :
                    arg = arg.Remove(0, 1);
                    // Strip quotes
                    if (arg.EndsWith("\""))
                        arg = arg.Remove(arg.Length - 1, 1);
                    if (arg.StartsWith("\""))
                        arg = arg.Remove(0, 1);
                    if (arg.Equals(string.Empty))
                        return defaultValue;
                    else
                        return arg;
                }
                else if (arg.Equals(string.Empty))
                {
                    // Empty value is OK
                    return defaultValue;
                }
            }
            return null;
        }

        /// <summary>
        /// Reads all of the lines of text from the source stream
        /// </summary>
        /// <param name="source"></param>
        /// <returns>A list of string</returns>
        private static List<string> Read(StreamReader source)
        {
            List<string> text = new List<string>();
            string s;
            while ((s = source.ReadLine()) != null)
                text.Add(s);
            return text;
        }

        /// <summary>
        /// Invokes POWERCFG.EXE with an argument and returns the text it produces
        /// </summary>
        /// <param name="argument">E.g. /requests or /requestsoverride</param>
        /// <returns>A list of string</returns>
        private static List<string> InvokePowercfg(string argument)
        {
            const string powerCfgExe = "powercfg.exe"; // The name of the PowerCfg.exe
            Process p = new Process();
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.FileName = powerCfgExe;
            p.StartInfo.Arguments = argument;
            p.Start();
            // Read the reply
            List<string> output = Read(p.StandardOutput);
            p.WaitForExit();
            p.Dispose();
            return output;
        }

        /// <summary>
        /// Returns true if there are SYSTEM requests that haven't been overridden
        /// </summary>
        /// <returns>true if there are SYSTEM requests that are preventing sleep</returns>
        private static bool ReadPowercfgRequests()
        {
            // Constants specific to working with PowerCfg.exe
            const string requestsArg = "/requests";        // PowerCfg argument to ask for PowerCfg requests
            const string overrideArg = "/requestsoverride";// PowerCfg argument to ask for PowerCfg overrides
            const string nullRequest = "None.";            // The "empty set" of request types
            const string systemType = "SYSTEM";            // A particular requestType having to do with sleep requests
            // Types of callers that can make requests
            List<string> requestCallers = new List<string>(new string[] { "[DRIVER]", "[PROCESS]", "[SERVICE]" });

            // Start with the requests
            List<string> requests = InvokePowercfg(requestsArg);

            // Parse the requests
            Dictionary<string, List<string>> typeDictionary = new Dictionary<string, List<string>>();
            string type = null;
            foreach (string s in requests)
            {
                if (s.EndsWith(":"))
                {
                    // We have a new type
                    type = s.Remove(s.Length - 1, 1);
                    if (!typeDictionary.ContainsKey(type))
                        typeDictionary.Add(type, new List<string>());
                }
                else if (type != null && !string.IsNullOrEmpty(s) && string.Compare(s, nullRequest, true) != 0)
                {
                    // Filter out comments to get just the real requests
                    foreach (string caller in requestCallers)
                    {
                        if (s.StartsWith(caller))
                            typeDictionary[type].Add(s);
                    }
                }
            }

            // Focus on just the SYSTEM requests
            if (!typeDictionary.ContainsKey(systemType))
                return false;
            List<string> systemRequests = typeDictionary[systemType];
            if (systemRequests.Count > 0)
            {
                // There are SYSTEM requests, now we need to get the overrides to see which ones to honor
                if (verbose)
                {
                    Console.WriteLine(StringTable.Requests);
                    foreach (string s in systemRequests)
                        Console.WriteLine("  " + s);
                }
                List<string> overrides = InvokePowercfg(overrideArg);

                // Parse the overrides
                Dictionary<string, List<string>> overridesDictionary = new Dictionary<string, List<string>>();
                bool overridesPresent = false;
                string caller = null;
                string systemCaller = " " + systemType;
                foreach (string s in overrides)
                {
                    if (s.StartsWith("[") && s.EndsWith("]") && requestCallers.Contains(s))
                    {
                        caller = s;
                        if (!overridesDictionary.ContainsKey(caller))
                            overridesDictionary.Add(caller, new List<string>());
                    }
                    else if (caller != null && !string.IsNullOrEmpty(s) && s.Contains(systemCaller))
                    {
                        string requestOverride = s.Replace(systemCaller, string.Empty);
                        overridesDictionary[caller].Add(requestOverride);
                        overridesPresent = true;

                    }
                }

                if (verbose && overridesPresent)
                {
                    Console.WriteLine(StringTable.Overrides);
                    foreach (string key in overridesDictionary.Keys)
                    {
                        foreach (string s in overridesDictionary[key])
                            Console.WriteLine("  " + s);
                    }
                }

                // Count how many of the systemRequests are overridden
                int[] overridenRequests = new int[systemRequests.Count];
                for (int i = 0; i < systemRequests.Count; i++)
                {
                    overridenRequests[i] = i;
                    string request = systemRequests[i];
                    string requestAllCaps = request.ToUpper();
                    foreach (string key in overridesDictionary.Keys)
                    {
                        if (request.StartsWith(key))
                        {
                            foreach (string s in overridesDictionary[key])
                            {
                                string callerAllCaps = s.ToUpper();
                                if (requestAllCaps.Contains(callerAllCaps))
                                    overridenRequests[i] = -1;
                            }
                        }
                    }
                }

                // Determine which systemRequests are not overridden, these are the ones that must be honored
                List<string> validRequests = new List<string>(overridenRequests.Where(x => x >= 0).Select(x => systemRequests[x]));
                if (validRequests.Count > 0)
                {
                    if (verbose)
                    {
                        Console.WriteLine(StringTable.RequestsToHonor);
                        foreach (string request in validRequests)
                            Console.WriteLine("  " + request);
                    }
                    return true;
                }
                else
                {
                    if (verbose)
                        Console.WriteLine(StringTable.RequestsOverridden);
                    return false;
                }
            }
            else
            {
                // There are no requests
                if (verbose)
                    Console.WriteLine(StringTable.NoRequests);

                return false;
            }
        }

        /// <summary>
        /// Get a resource stream
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static Stream GetResourceStream(string key)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string name = string.Format("{0}.{1}", "SleepingPill", key); // Should match namespace
            return assembly.GetManifestResourceStream(name);
        }

        /// <summary>
        /// Displays 'Press any key...' and but will not wait forever for a key stroke
        /// </summary>
        private static void PressAnyKeyWithTimout()
        {
            Console.Write(StringTable.PressAnyKey);
            Thread timeoutThread = new Thread(new ThreadStart(TimedTermination));
            timeoutThread.Start();
            Console.Read();
            timeoutThread.Abort();
        }

        /// <summary>
        /// If SleepingPill is run from TaskScheduler, there won't be any UI for the user to
        /// see or respond to. Therefore, we must implement a timeout.
        /// </summary>
        private static void TimedTermination()
        {
            try
            {
                Thread.Sleep(60000); // Wait for a minute then terminate
                Environment.Exit(-1);
            }
            catch (ThreadAbortException)
            {
            }
        }

        /// <summary>
        /// Builds the path for storing local program data - creates the path if it doesn't exist
        /// </summary>
        public static string GetLocalUserAppDataPath()
        {
            if (localUserAppDataPath == null)
            {
                string localAppDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                AssemblyProductAttribute productAttribute = null;
                AssemblyCompanyAttribute companyAttribute = null;
                foreach (object o in typeof(Program).Assembly.GetCustomAttributes())
                {
                    if (o is AssemblyProductAttribute)
                        productAttribute = (AssemblyProductAttribute)o;
                    if (o is AssemblyCompanyAttribute)
                        companyAttribute = (AssemblyCompanyAttribute)o;
                }
                localUserAppDataPath = Path.Combine(Path.Combine(localAppDataPath, companyAttribute.Company), productAttribute.Product);
                if (!Directory.Exists(localUserAppDataPath))
                    Directory.CreateDirectory(localUserAppDataPath);
            }
            return localUserAppDataPath;
        }

        /// <summary>
        /// Increment the sleep count file
        /// </summary>
        private static void IncrementSleepCountFile()
        {
            // Increment the sleep count
            string countFile = Path.Combine(GetLocalUserAppDataPath(), "SleepCount.txt");
            StreamReader reader = null;
            try
            {
                using (FileStream stream = new FileStream(countFile, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    int count = 0;
                    reader = new StreamReader(stream);
                    string s = reader.ReadLine();
                    if (s != null)
                    {
                        int endOfInt = s.IndexOf(" ");
                        if (endOfInt > 0)
                            s = s.Substring(0, endOfInt);
                        int.TryParse(s, out count);
                    }
                    stream.Position = 0;
                    count++;
                    using (StreamWriter writer = new StreamWriter(stream))
                        writer.WriteLine(string.Format("{0}  {1}", count.ToString(), DateTime.Now.ToString()));
                }
            }
            catch (IOException error)
            {
                Console.WriteLine(string.Format(StringTable.FailedToUpdateSleepCountFile, countFile, error.Message));
            }
            finally
            {
                if (reader != null)
                    reader.Dispose();
            }
        }
    }
}
