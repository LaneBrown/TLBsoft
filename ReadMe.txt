This program puts the computer to sleep if the screen saver has been running
for 'Time' seconds or more and PowerCfg reports no SYSTEM requests
for 'Idle' seconds or more (i.e. requests that are not overridden). See:
https://learn.microsoft.com/en-us/windows-hardware/design/device-experiences/powercfg-command-line-options#option_requestsoverride
SleepingPill.exe is for 64 bit Windows, SleepingPill_x86.exe is for 32-bit
Windows.

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
