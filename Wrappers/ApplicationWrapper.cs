using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using Microsoft.Office.Core;

namespace VSTO_AddIn
{

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(Outlook.ApplicationEvents_11)), Guid("0a9069af-8b5b-46ff-8182-6493b8f8f2fe")]
    public class ApplicationWrapper : Outlook.ApplicationEvents_11, IDisposable
    {
        private readonly Guid _sourceIntfGuid = new Guid("0006302C-0000-0000-C000-000000000046");
        private Int32 _cookie;
        private IConnectionPoint _pConnectionPoint = null;
        private IConnectionPointContainer _pConnectionPointContainer = null;
        private bool _isDisposed = false;
        object _myObject = null;
        Outlook.Application application;

        public ApplicationWrapper(object obj)
        {
            Utility.LogApplicationEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

            this._myObject = obj;
            _pConnectionPointContainer = _myObject as IConnectionPointContainer;
            if (_pConnectionPointContainer != null)
            {
                _pConnectionPoint = null;
                _pConnectionPointContainer.FindConnectionPoint(_sourceIntfGuid, out _pConnectionPoint);
                if (_pConnectionPoint != null)
                {
                    _pConnectionPoint.Advise(this, out _cookie);
                    application = _myObject as Outlook.Application;
                    if (application != null)
                    {
                        new ExplorersWrapper(application.Explorers); // application.Explorers returns a COM object that is released when the ExplorersWrapper instance is disposed of so no need to release here
                        new ExplorerWrapper(application.ActiveExplorer()); // application.ActiveExplorer() returns a COM object that is released when the ExplorerWrapper instance is disposed of so no need to release here
                        new InspectorsWrapper(application.Inspectors); // application.Inspectors returns a COM object that is released when the ExplorersWrapper instance is disposed of so no need to release here
                        new NameSpaceWrapper(application.Session);

                    }
                }
            }
        }

        public void Dispose()
        {
            Utility.LogApplicationEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            Utility.LogApplicationEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
            if (!_isDisposed)
            {
                if (disposing)
                {
                    if (_cookie != -1 && _pConnectionPoint != null)
                    {
                        _pConnectionPoint.Unadvise(_cookie);
                        _pConnectionPoint = null;
                        _cookie = -1;
                    }
                }

                if (_pConnectionPointContainer != null)
                {
                    // there's an implicit QueryInterface cast when calling 
                    //_pConnectionPointContainer = _myObject as IConnectionPointContainer;
                    int refCount = Marshal.ReleaseComObject(_pConnectionPointContainer);
                    Utility.LogApplicationEvent(LogType.Diagnostic, (string.Format("Reference count for _pConnectionPointContainer  in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    _pConnectionPointContainer = null;
                }

                if (_myObject != null)
                {
                    // when we're disposing of instance we should also dispose of the source object 
                    int refCount = Marshal.ReleaseComObject(_myObject);
                    Utility.LogApplicationEvent(LogType.Diagnostic, (string.Format("Reference count for _myObject  in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    _myObject = null;
                }

                if (application != null)
                {
                    // when we're disposing of instance we should also dispose of the source object 
                    int refCount = Marshal.ReleaseComObject(application);
                    Utility.LogApplicationEvent(LogType.Diagnostic, (string.Format("Reference count for application  in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    application = null;
                }

                _isDisposed = true;
            }
        }

        public void ItemSend(object Item, ref bool Cancel)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void NewMail()
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void Reminder(object Item)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void OptionsPagesAdd(Outlook.PropertyPages Pages)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void Startup()
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void Quit()
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void AdvancedSearchComplete(Outlook.Search SearchObject)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void AdvancedSearchStopped(Outlook.Search SearchObject)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void MAPILogonComplete()
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void NewMailEx(string EntryIDCollection)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void AttachmentContextMenuDisplay(CommandBar CommandBar, Outlook.AttachmentSelection Attachments)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void FolderContextMenuDisplay(CommandBar CommandBar, Outlook.MAPIFolder Folder)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void StoreContextMenuDisplay(CommandBar CommandBar, Outlook.Store Store)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void ShortcutContextMenuDisplay(CommandBar CommandBar, Outlook.OutlookBarShortcut Shortcut)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

        }

        public void ViewContextMenuDisplay(CommandBar CommandBar, Outlook.View View)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void ItemContextMenuDisplay(CommandBar CommandBar, Outlook.Selection Selection)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void ContextMenuClose(Outlook.OlContextMenu ContextMenu)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void ItemLoad(object Item)
        {
            new ItemWrapper(Item);
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void BeforeFolderSharingDialog(Outlook.MAPIFolder FolderToShare, ref bool Cancel)
        {
            Utility.LogApplicationEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }
    }
}