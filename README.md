OutlookFormsManager
    Outlook forms management utility.

Parameters
    -m    : Running mode. Possible values are: importform, removeform, clearformcache
            setdefaultinboxform, setdefaultcalendarform, and resetdefaultforms
    -m    : Path to the .ost exported form file.
    -n    : Custom form name.
    -c    : Custom form message class. Only useable in conjunction with setdefaultinboxform
            and setdefaultcalendarform,
    -?    : Displays this information.

Example 1: Import a calendar custom form and set it as default form for new appointments:
    OutlookFormsManager.exe -m importform -p C:\Forms\CustomMeeting.oft -n CustomMetting2 -c IPM.Appointment.CustomMetting -m setdefaultcalendarform

Example 2: Reset the Inbox and Calendar folders to their default forms
    OutlookFormsManager.exe -m resetdefaultforms

Example 3: Removes a specified form from the Personal Forms Library
    OutlookFormsManager.exe -m removeform -n CustomMeeting

Example 4: Clear the local forms cache
    OutlookFormsManager.exe -m clearformcache
