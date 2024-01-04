The problem: my Windows 10 computer won't reliably go to sleep automatically.
It does, however, reliably start the screen saver.
I tried on-line suggestions and recommendations for 2 days; finally, I decided
it would be faster to just write a C# console application. To anyone reading
this, please don't suggest things to try in Windows; I've already spent too
much time trying this and that.

Solution:
This C# console app puts the computer to sleep if the screen saver has been running
for 'Time' seconds or more and PowerCfg reports no SYSTEM requests
for 'Idle' seconds or more (i.e. requests that are not overridden). See:
https://learn.microsoft.com/en-us/windows-hardware/design/device-experiences/powercfg-command-line-options#option_requestsoverride
Unfortunately, invoking PowerCfg requires that the app be in a 32-bit or 64-bit version
that matches the operating system)and that it be run as administrator.
In this repository, I've included both:
SleepingPill.exe is for 64 bit Windows, SleepingPill_x86.exe is for 32-bit
Windows.

Testing:
I used Task Scheduler to run the app at startup. I then ran 'Groove Music' and noted
that as long it's playing music, it makes SYSTEM requests that are visible to PowerCfg.
The computer didn't go to sleep as long as it was playing music but did go to sleep
shortly after it stopped. This is the behavior I wanted to see.
I also ran tests where I copied many large files (requiring at least 30 minutes) in
both a push and pull test configuration. For the push test, I initiated the copy
from the Windows 10 computer to another computer on the network. For the pull test,
I initiated the copy from another computer on the network. It was done this way because
the PowerCfg SYSTEM requests are different - in both cases the computer stayed awake
while copying files and went to sleep automatically afterwards (as it should).
I also tried using PowerCfg /requestsoverride to override a process (like Groove Music)
and the computer did go sleep even with the process running. This is the desired result.
In all of my testing, I was satisfied that SleepingPill.exe fixed the insomia problems
my computer was having without introducing un-desired side-effects. I've also successfully
tested it on a Windows 7 computer.

Recommendation: copy SleepingPill.exe to a secure location that requires
  admin priviledges such as somewhere in ProgramFiles. Use Task Scheduler
  to run SleepingPill.exe with admin priviledges at startup + 1 minute.
  Note: admin priviledges are required to properly invoke PowerCfg.exe
  SleepingPill should run whether user is logged on or not.
  For the action's arguments, leave blank or see the next section:

Arguments (space delimited, order doesn't matter, default used if omitted)
  /Process=<string>       Process name e.g. Mystify.scr (default = any)
                          Enclose name in quotes if it contains spaces
  /Time=<integer>         Screen saver must be running for this much time
                          (default = 180 seconds)
  /Idle=<integer>         No PowerCfg SYSTEM requests for this much time
                          (default = 20 seconds)
  /Sample=<integer>       How often to test screen saver and PowerCfg
                          (default = 5 seconds)
  /Verbose=<boolean>      Reports/displays details while running
                          Good for testing but useless for normal usage.
                          (default = false)
Example: SleepingPill /process=Mystify.scr -TIME=300 idle=10 /Verbose=true
  Put a 64-bit computer to sleep when Mystify.scr has been running for
  5 min. and there are no SYSTEM requests for 10 seconds; report details.
Example: SleepingPill
  Put a 64-bit computer to sleep when any screen saver has been running for
  3 minutes and there are no SYSTEM requests for 20 seconds; run silently.
Example: SleepingPill_x86 /time:120 -idle:15 verbose
  Put a 32-bit computer to sleep when a screen saver has been running for
  2 min. and there are no SYSTEM requests for 15 seconds, report details.

For more details, see the 'SleepingPill and TaskScheduler.zip' file in this repository.
It contains screen shots of Task Scheduler and more detailed instructions for set up.

The program.cs file contains the bulk of the software for anyone interested in seeing how I did it. I welcome comments/suggestions. The main thing that's not included is the StringTable.resx file and the icon file. If someone wishes to build it himself, I can
include all of the ancillary project files. I used VisualStudio 2013 (a bit old, I know).
