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
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(Outlook.StoresEvents_12)), Guid("f771ef1f-ea50-48e6-ac46-3354162c1c4f")]
    public class StoresWrapper : Outlook.StoresEvents_12, IDisposable
    {
        private readonly Guid _sourceIntfGuid = new Guid("000630F8-0000-0000-C000-000000000046");
        private Int32 _cookie;
        private IConnectionPoint _pConnectionPoint = null;
        private IConnectionPointContainer _pConnectionPointContainer = null;
        private bool _isDisposed = false;
        object _myObject = null;

        public StoresWrapper(object obj)
        {
            Utility.LogStoresEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

            this._myObject = obj;
            _pConnectionPointContainer = _myObject as IConnectionPointContainer;
            if (_pConnectionPointContainer != null)
            {
                _pConnectionPoint = null;
                _pConnectionPointContainer.FindConnectionPoint(_sourceIntfGuid, out _pConnectionPoint);
                if (_pConnectionPoint != null)
                {
                    _pConnectionPoint.Advise(this, out _cookie);
                }
            }
        }

        public void Dispose()
        {
            Utility.LogStoresEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            Utility.LogStoresEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
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
                    Utility.LogStoresEvent(LogType.Diagnostic, (string.Format("Reference count for _pConnectionPointContainer  in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    _pConnectionPointContainer = null;
                }

                if (_myObject != null)
                {
                    // when we're disposing of instance we should also dispose of the source object 
                    int refCount = Marshal.ReleaseComObject(_myObject);
                    Utility.LogStoresEvent(LogType.Diagnostic, (string.Format("Reference count for _myObject  in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    _myObject = null;
                }

                _isDisposed = true;
            }
        }

        public void BeforeStoreRemove(Outlook.Store Store, ref bool Cancel)
        {
            Utility.LogStoresEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void StoreAdd(Outlook.Store Store)
        {
            Utility.LogStoresEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }
    }
}
