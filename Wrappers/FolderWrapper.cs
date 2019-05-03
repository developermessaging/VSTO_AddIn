using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace VSTO_AddIn
{
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(Outlook.MAPIFolderEvents_12)), Guid("7c05484d-92b9-43d3-88b4-ab3a4acf5ca6")]
    public class FolderWrapper : Outlook.MAPIFolderEvents_12, IDisposable
    {
        private readonly Guid _sourceIntfGuid = new Guid("000630F7-0000-0000-C000-000000000046");
        private Int32 _cookie;
        private IConnectionPoint _pConnectionPoint = null;
        private IConnectionPointContainer _pConnectionPointContainer = null;
        private bool _isDisposed = false;
        object _myObject = null;

        Outlook.MAPIFolder mAPIFolder = null;
        Outlook.Items items = null;
        string folderName;

        public FolderWrapper(object obj)
        {
            Utility.LogFolderEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

            this._myObject = obj;
            _pConnectionPointContainer = _myObject as IConnectionPointContainer;
            if (_pConnectionPointContainer != null)
            {
                _pConnectionPoint = null;
                _pConnectionPointContainer.FindConnectionPoint(_sourceIntfGuid, out _pConnectionPoint);
                if (_pConnectionPoint != null)
                {
                    _pConnectionPoint.Advise(this, out _cookie);
                    mAPIFolder = obj as Outlook.MAPIFolder;
                    items = mAPIFolder.Items;
                    folderName = mAPIFolder.Name;
                    new ItemsWrapper(items, folderName); // mAPIFolder.Items should be released within the wrapper
                }
            }
        }

        public void Dispose()
        {
            Utility.LogFolderEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            Utility.LogFolderEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
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
                    Utility.LogFolderEvent(LogType.Diagnostic, (string.Format("Reference count for _pConnectionPointContainer  in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    _pConnectionPointContainer = null;
                }

                if (_myObject != null)
                {
                    // when we're disposing of instance we should also dispose of the source object 
                    int refCount = Marshal.ReleaseComObject(_myObject);
                    Utility.LogFolderEvent(LogType.Diagnostic, (string.Format("Reference count for _myObject  in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    _myObject = null;
                }

                if (mAPIFolder != null)
                {
                    // when we're disposing of instance we should also dispose of the source object 
                    int refCount = Marshal.ReleaseComObject(mAPIFolder);
                    Utility.LogFolderEvent(LogType.Diagnostic, (string.Format("Reference count for mAPIFolder  in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    mAPIFolder = null;
                }

                _isDisposed = true;
            }
        }

        public void BeforeFolderMove(Outlook.MAPIFolder MoveTo, ref bool Cancel)
        {
            Utility.LogFolderEvent(LogType.Diagnostic, (string.Format("{0} in folder {1}{2}", new StackTrace().GetFrame(0).GetMethod().Name, folderName, System.Environment.NewLine)));
        }

        public void BeforeItemMove(object Item, Outlook.MAPIFolder MoveTo, ref bool Cancel)
        { 
            Utility.LogFolderEvent(LogType.Diagnostic, (string.Format("{0} in folder {1}{2}", new StackTrace().GetFrame(0).GetMethod().Name, folderName, System.Environment.NewLine)));
        }
}
}
