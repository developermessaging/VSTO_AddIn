using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using Microsoft.Office.Core;

namespace VSTO_AddIn
{
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(Outlook.NameSpaceEvents)), Guid("f161aa73-49f5-433e-8343-04a24a86d120")]
    public class NameSpaceWrapper : Outlook.NameSpaceEvents, IDisposable
    {
        private readonly Guid _sourceIntfGuid = new Guid("0006308C-0000-0000-C000-000000000046");
        private Int32 _cookie;
        private IConnectionPoint _pConnectionPoint = null;
        private IConnectionPointContainer _pConnectionPointContainer = null;
        private bool _isDisposed = false;
        object _myObject = null;

        Outlook.NameSpace nameSpace = null;
        Outlook.MAPIFolder inboxMAPIFolder = null;
        Outlook.MAPIFolder outboxMAPIFolder = null;
        Outlook.MAPIFolder sentItemsMAPIFolder = null;

        public NameSpaceWrapper(object obj)
        {
            Utility.LogNameSpaceEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

            this._myObject = obj;
            _pConnectionPointContainer = _myObject as IConnectionPointContainer;
            if (_pConnectionPointContainer != null)
            {
                _pConnectionPoint = null;
                _pConnectionPointContainer.FindConnectionPoint(_sourceIntfGuid, out _pConnectionPoint);
                if (_pConnectionPoint != null)
                {
                    _pConnectionPoint.Advise(this, out _cookie);

                    nameSpace = _myObject as Outlook.NameSpace;
                    inboxMAPIFolder = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); // inboxMAPIFolder is released within the wrapper
                    outboxMAPIFolder = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox); // outboxMAPIFolder is released within the wrapper
                    sentItemsMAPIFolder = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail); // sentItemsMAPIFolder is released within the wrapper
                    if (ControlPanel.TrackInboxEnabled) new FolderWrapper(inboxMAPIFolder);
                    if (ControlPanel.TrackOutboxEnabled) new FolderWrapper(outboxMAPIFolder);
                    if (ControlPanel.TrackSentItemsEnabled) new FolderWrapper(sentItemsMAPIFolder);
                }
            }
        }

        public void Dispose()
        {
            Utility.LogNameSpaceEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            Utility.LogNameSpaceEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
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
                    Utility.LogNameSpaceEvent(LogType.Diagnostic, (string.Format("Reference count for _pConnectionPointContainer in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    _pConnectionPointContainer = null;
                }

                if (_myObject != null)
                {
                    // when we're disposing of instance we should also dispose of the source object 
                    int refCount = Marshal.ReleaseComObject(_myObject);
                    Utility.LogNameSpaceEvent(LogType.Diagnostic, (string.Format("Reference count for _myObject in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    _myObject = null;
                }

                if (nameSpace != null)
                {
                    // when we're disposing of instance we should also dispose of the source object 
                    int refCount = Marshal.ReleaseComObject(nameSpace);
                    Utility.LogNameSpaceEvent(LogType.Diagnostic, (string.Format("Reference count for nameSpace in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    nameSpace = null;
                }

                _isDisposed = true;
            }
        }

        public void OptionsPagesAdd(Outlook.PropertyPages Pages, Outlook.MAPIFolder Folder)
        {
            Utility.LogNameSpaceEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void AutoDiscoverComplete()
        {
            Utility.LogNameSpaceEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }
    }
}
