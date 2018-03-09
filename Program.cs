using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookFormsManager
{
    class Program
    {

        public struct RuntimeOptions
        {
            public ulong mode;
            public string formPath;
            public string formName;
            public string formClass;
        }

        const ulong ImportForm = 0x0010;
        const ulong RemoveForm = 0x0020;
        const ulong ClearFormCache = 0x0040;
        const ulong SetDefaultInboxForm = 0x0080;
        const ulong SetDefaultCalendarForm = 0x0100;
        const ulong ResetDefaultForms = 0x0200;
    
        static void PrintHelp()
        {
            Console.WriteLine("OutlookFormsManager");
            Console.WriteLine("    Outlook forms management utility.");
            Console.WriteLine();
            Console.WriteLine("Parameters");
            Console.WriteLine("    -m    : Running mode. Possible values are: importform, removeform, clearformcache");
            Console.WriteLine("            setdefaultinboxform, setdefaultcalendarform, and resetdefaultforms");
            Console.WriteLine("    -m    : Path to the .ost exported form file.");
            Console.WriteLine("    -n    : Custom form name. ");
            Console.WriteLine("    -c    : Custom form message class. Only useable in conjunction with setdefaultinboxform");
            Console.WriteLine("            and setdefaultcalendarform,");
            Console.WriteLine("    -?    : Displays this information.");
            Console.WriteLine();
            Console.WriteLine("Example 1: Import a calendar custom form and set it as default form for new appointments:");
            Console.WriteLine("    OutlookFormsManager.exe -m importform -p C:\\Forms\\CustomMeeting.oft -n CustomMetting2 -c IPM.Appointment.CustomMetting -m setdefaultcalendarform ");
            Console.WriteLine();
            Console.WriteLine("Example 2: Reset the Inbox and Calendar folders to their default forms");
            Console.WriteLine("    OutlookFormsManager.exe -m resetdefaultforms");
            Console.WriteLine();
            Console.WriteLine("Example 3: Removes a specified form from the Personal Forms Library");
            Console.WriteLine("    OutlookFormsManager.exe -m removeform -n CustomMeeting");
            Console.WriteLine();
            Console.WriteLine("Example 4: Clear the local forms cache");
            Console.WriteLine("    OutlookFormsManager.exe -m clearformcache");
            Console.WriteLine();
        }


        static bool ParseArguments(int argc, string[] argv, ref RuntimeOptions runtimeOptions)
        {
            runtimeOptions = new RuntimeOptions();
            for (int i = 0; i < argc; i++)
            {
                switch (argv[i][0])
                {
                    case '-':
                    case '/':
                    case '\\':
                        if (0 == argv[i][1])
                        {
                            return false;
                        }
                        switch (Char.ToLower(argv[i][1]))
                        {
                            case 'm':
                                if (i + 1 < argc)
                                {
                                    if (argv[i + 1].ToLower() == "importform")
                                    {
                                        runtimeOptions.mode = runtimeOptions.mode | ImportForm;
                                        i++;
                                    }
                                    else if (argv[i + 1].ToLower() == "removeform")
                                    {
                                        runtimeOptions.mode = runtimeOptions.mode | RemoveForm;
                                        i++;
                                    }
                                    else if (argv[i + 1].ToLower() == "clearformcache")
                                    {
                                        runtimeOptions.mode = runtimeOptions.mode | ClearFormCache;
                                        i++;
                                    }
                                    else if (argv[i + 1].ToLower() == "setdefaultinboxform")
                                    {
                                        runtimeOptions.mode = runtimeOptions.mode | SetDefaultInboxForm;
                                        i++;
                                    }
                                    else if (argv[i + 1].ToLower() == "setdefaultcalendarform")
                                    {
                                        runtimeOptions.mode = runtimeOptions.mode | SetDefaultCalendarForm;
                                        i++;
                                    }
                                    else if (argv[i + 1].ToLower() == "resetdefaultforms")
                                    {
                                        runtimeOptions.mode = runtimeOptions.mode | ResetDefaultForms;
                                        i++;
                                    }
                                    else
                                    { return false; }
                                }
                                else
                                    return false;
                                break;
                            case 'p':
                                if (i + 1 < argc)
                                {
                                    runtimeOptions.formPath = argv[i + 1];
                                    i++;
                                }
                                else
                                    return false;
                                break;
                            case 'n':
                                if (i + 1 < argc)
                                {
                                    runtimeOptions.formName = argv[i + 1];
                                    i++;
                                }
                                else
                                    return false;
                                break;
                            case 'c':
                                if (i + 1 < argc)
                                {
                                    runtimeOptions.formClass = argv[i + 1];
                                    i++;
                                }
                                else
                                    return false;
                                break;
                            case '?':
                                return false;
                            default:
                                return false;
                        }
                        break;
                }
            }

            if (runtimeOptions.mode == ImportForm)
            {
                if (runtimeOptions.formName != null && runtimeOptions.formPath != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else if ((runtimeOptions.mode & RemoveForm) == RemoveForm)
            {
                if (runtimeOptions.formName != null)
                {
                    return true;
                }
                else
                    return false;
            }
            else if (((runtimeOptions.mode & SetDefaultInboxForm) == SetDefaultInboxForm) || ((runtimeOptions.mode & SetDefaultCalendarForm) == SetDefaultCalendarForm))
            {
                if ((runtimeOptions.formName != null) && (runtimeOptions.formClass != null))
                {
                    return true;
                }
                else
                    return false;
            }
            return true;
        }

        [STAThread]
        static void Main(string[] args)
        {
            RuntimeOptions runtimeOptions = new RuntimeOptions();
            if (!ParseArguments(args.Length, args, ref runtimeOptions))
            {
                PrintHelp();
                return;
            }

            if ((runtimeOptions.mode & ImportForm) == ImportForm)
                    OutlookFormsProxy.ImportFormToPersonalFormsLibrary(runtimeOptions.formPath, runtimeOptions.formName);
            if ((runtimeOptions.mode & RemoveForm) == RemoveForm)
                    OutlookFormsProxy.RemoveFormFromPersonalFormsLibrary(runtimeOptions.formName);
            if ((runtimeOptions.mode & ClearFormCache) == ClearFormCache)
                    OutlookFormsProxy.ClearLocalFormCache();
            if ((runtimeOptions.mode & SetDefaultCalendarForm) == SetDefaultCalendarForm)
                OutlookFormsProxy.SetDefaultCalendarForm(runtimeOptions.formName, runtimeOptions.formClass);
            if ((runtimeOptions.mode & SetDefaultInboxForm) == SetDefaultInboxForm)
                OutlookFormsProxy.SetDefaultInboxForm(runtimeOptions.formName, runtimeOptions.formClass);
            if ((runtimeOptions.mode & ResetDefaultForms) == ResetDefaultForms)
                Console.WriteLine("Reset logic not yet implemented!");  
                //OutlookFormsProxy.ResetDefaultForms();
            OutlookFormsProxy.Exit();
        }

    }
}
