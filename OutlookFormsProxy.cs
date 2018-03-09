using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Permissions;

namespace OutlookFormsManager
{
    public static class OutlookFormsProxy
    {
        private static Outlook.Application m_Application = null;
        private static Outlook.NameSpace m_NameSpace = null;
        private static bool newOutlookInstance = false;
        const string PR_COMMON_VIEWS_ENTRY_ID = "http://schemas.microsoft.com/mapi/proptag/0x35E60102";
        const string SEARCH_FORM_MESSAGECLASS = "http://schemas.microsoft.com/mapi/proptag/0x6800001E";
        const string PR_DISPLAY_NAME = "http://schemas.microsoft.com/mapi/proptag/0x3001001E";
        const string PR_ENTRY_ID = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102";
        const string PR_LONG_TERM_ENTRYID_FROM_TABLE = "http://schemas.microsoft.com/mapi/proptag/0x66700102";
        const string prDefMsgClass = "http://schemas.microsoft.com/mapi/proptag/0x36E5001E";
        const string prDefFormName = "http://schemas.microsoft.com/mapi/proptag/0x36E6001E";
        private static void Initialise()
        {
            if (null == m_Application)
            {
                try
                {
                    SecurityPermission secper = new SecurityPermission(PermissionState.Unrestricted);
                    secper.Assert();
                    object olApplicationInstance = null;
                    try
                    {
                        olApplicationInstance = Marshal.GetActiveObject("Outlook.Application");
                    }
                    catch (COMException comException)
                    {
                        if ((uint)comException.ErrorCode == 0x800401e3)
                        {
                            Console.WriteLine("No running instance of Outlook found.");
                        }
                    }
                    SecurityPermission.RevertAssert();

                    if (null != olApplicationInstance)
                    {
                        if (olApplicationInstance is Outlook.Application)
                        {
                            Console.WriteLine("Binding to existing Outlook instance.");
                            m_Application = olApplicationInstance as Outlook.Application;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Creating a new Outlook instance.");
                        m_Application = new Outlook.Application();
                        newOutlookInstance = true;
                    }
                    if (null != m_Application)
                    {
                        m_NameSpace = m_Application.GetNamespace("MAPI");
                    }
                    else
                    {
                        Console.WriteLine("ERROR: Unable to start Outlook!");
                    }

                }
                catch (Exception exception)
                {
                    Console.WriteLine("Unable to initialise application object. Error: " + exception.ToString());
                }
            }
        }

        public static void Exit()
        {
            if (null != m_Application)
            {
                m_NameSpace = null;
                if (newOutlookInstance)
                {
                    m_Application.Quit();
                }
                m_Application = null;
            }
        }

        public static void ImportFormToPersonalFormsLibrary(string path, string name)
        {
            if (null == m_Application)
            {
                Initialise();
            }
            try
            {
                Outlook.FormDescription formDescription = null;
                Outlook.MailItem mailItem = null;
                Outlook.AppointmentItem appointmentItem = null;
                dynamic formTemplate = m_Application.CreateItemFromTemplate(path);
                if (formTemplate is Outlook.MailItem)
                {
                    mailItem = formTemplate as Outlook.MailItem;
                    formDescription = mailItem.FormDescription;
                }
                else if (formTemplate is Outlook.AppointmentItem)
                {
                    appointmentItem = formTemplate as Outlook.AppointmentItem;
                    formDescription = appointmentItem.FormDescription;
                };
                if (formDescription != null)
                {
                    formDescription.Name = name;
                    formDescription.PublishForm(Outlook.OlFormRegistry.olPersonalRegistry);
                    ReleaseObject(formDescription);
                    formDescription = null;
                }
                
                if (mailItem != null)
                {
                    ReleaseObject(mailItem);
                }
                if (appointmentItem != null)
                {
                    ReleaseObject(appointmentItem);
                }

            }
            catch (Exception exception)
            {
                Console.WriteLine("Unable to import custom form. Error: " + exception.ToString());
            }
            finally
            {

            }

        }

        public static void ReleaseObject(object obj)
        {
            Marshal.ReleaseComObject(obj);
        }

        private static Outlook.MAPIFolder GetCommonViewsFolder()
        {
            try
            {
                if (null == m_NameSpace)
                {
                    Initialise();
                }
                Outlook.Store olStore = m_NameSpace.DefaultStore;
                Outlook.PropertyAccessor olPropertyAccessor = olStore.PropertyAccessor;
                string commonViewsEntryId = olPropertyAccessor.BinaryToString(olPropertyAccessor.GetProperty(PR_COMMON_VIEWS_ENTRY_ID));
                if (commonViewsEntryId != string.Empty)
                {
                    return m_NameSpace.GetFolderFromID(commonViewsEntryId);
                }
                return null;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Unable to get common views folder. Error: " + exception.ToString());
                return null;
            }
        }

        public static void RemoveFormFromPersonalFormsLibrary(string name)
        {
            bool formFound = false;
            Outlook.Table olTable;
            Outlook.Row olRow;
            string searchFilter;
            if (null == m_Application)
            {
                Initialise();
            }
            try
            {
                Outlook.MAPIFolder olFolder = GetCommonViewsFolder();
                if (null != olFolder)
                {
                    searchFilter = "[MessageClass] = \"IPM.Microsoft.FolderDesign.FormsDescription\"";
                    olTable = olFolder.GetTable(searchFilter, Outlook.OlTableContents.olHiddenItems);
                    olTable.Columns.Add(PR_DISPLAY_NAME);
                    olTable.Columns.Add(SEARCH_FORM_MESSAGECLASS);
                    olTable.Columns.Add(PR_LONG_TERM_ENTRYID_FROM_TABLE);
                    olTable.Restrict(searchFilter);
                    while (!olTable.EndOfTable)
                    {
                        olRow = olTable.GetNextRow();
                        if (name.ToLower() == olRow[PR_DISPLAY_NAME].ToString().ToLower())
                        {
                            formFound = true;
                            byte[] entryId = olRow[PR_LONG_TERM_ENTRYID_FROM_TABLE];
                            string temp = "";
                            for (int i = 0; i < entryId.Length; i++)
                            {
                                temp += entryId[i].ToString("X2");
                            }
                            object item = m_NameSpace.GetItemFromID(temp, olFolder.StoreID);
                            if (item is Outlook.StorageItem)
                            {
                                Outlook.StorageItem storageItem = item as Outlook.StorageItem;
                                storageItem.Delete();
                                Console.WriteLine("Form succesfully deleted. You might need to restart Outlook.");
                            }
                        }
                    } 
                    if (!formFound)
                    {
                        Console.WriteLine("The form couldn't be found in the Personal Forms Library.");
                    }
                }
                ClearLocalFormCache();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Unable to remove form. Error: " + exception.ToString());
            }
            finally
            {
       
            }
        }

        public static void ClearLocalFormCache()
        {
            try
            {
                string LocalAppDataPath = Environment.ExpandEnvironmentVariables("%localappdata%");
                string FormCacheDatPath = LocalAppDataPath + "\\Microsoft\\Forms\\FRMCACHE.DAT";
                try
                {
                    File.Delete(FormCacheDatPath);
                    Console.WriteLine("Local forms cache succesfully deleted.");
                }
                catch (Exception exception)
                {
                    Console.WriteLine("Unable to delete local forms cache. Exception was: " + exception.ToString());
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("Unable to clear local forms cache. Error: " + exception.ToString());
            }
        }

        public static void SetDefaultInboxForm(string formName, string formClass)
        {
            if (null == m_NameSpace)
            {
                Initialise();
            }
            try
            {
                Outlook.MAPIFolder inboxFolder = m_NameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                SetDefaultFormOnFolder(inboxFolder, formName, formClass);
                Marshal.ReleaseComObject(inboxFolder);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Unable to set default inbox form. Error: " + exception.ToString());
            }

        }

        public static void SetDefaultCalendarForm(string formName, string formClass)
        {
            if (null == m_NameSpace)
            {
                Initialise();
            }
            try
            {
                Outlook.MAPIFolder calendarFolder = m_NameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                SetDefaultFormOnFolder(calendarFolder, formName, formClass);
                Marshal.ReleaseComObject(calendarFolder);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Unable to set default calendar form. Error: " + exception.ToString());
            }
        }

        public static void ResetDefaultForms()
        {
            if (null == m_NameSpace)
            {
                Initialise();
            }
            Outlook.MAPIFolder inboxFolder = null;
            Outlook.MAPIFolder calendarFolder = null;
            try
            {
                inboxFolder = m_NameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                SetDefaultFormOnFolder(inboxFolder, "IPM.Note", "Message");
                calendarFolder = m_NameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                SetDefaultFormOnFolder(inboxFolder, "IPM.Appointment", "Appointment");
            }
            catch (Exception exception)
            {
                Console.WriteLine("Unable to set reset to default forms. Error: " + exception.ToString());
            }
            finally
            {
                if (inboxFolder != null)
                    ReleaseObject(inboxFolder);
                if (calendarFolder != null)
                    ReleaseObject(calendarFolder);
            }
        }

        public static void SetDefaultFormOnFolder(Outlook.MAPIFolder mapiFolder, string formName, string formClass)
        {
            string prDefMsgClass = "http://schemas.microsoft.com/mapi/proptag/0x36E5001E";
            string prDefFormName = "http://schemas.microsoft.com/mapi/proptag/0x36E6001E";
            try
            {
                Outlook.PropertyAccessor propertyAccessor = mapiFolder.PropertyAccessor;
                propertyAccessor.SetProperty(prDefMsgClass, formClass);
                propertyAccessor.SetProperty(prDefFormName, formName);
                Marshal.ReleaseComObject(propertyAccessor);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Unable to set the default form. Error: " + exception.ToString());
            }
        }
    }
}
