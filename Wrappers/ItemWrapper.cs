using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Reflection;

namespace VSTO_AddIn
{
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(Outlook.ItemEvents_10)), Guid("82e1b877-a7d8-48de-9b24-5985482b5a14")]
    public class ItemWrapper : Outlook.ItemEvents_10, IDisposable
    {
        private readonly Guid _sourceIntfGuid = new Guid("0006302B-0000-0000-C000-000000000046");
        private Int32 _cookie;
        private IConnectionPoint _pConnectionPoint = null;
        private IConnectionPointContainer _pConnectionPointContainer = null;
        private bool _isDisposed = false;
        object _myObject = null;

        
        bool dispidFDirty = false;

        string subject;
        public ItemWrapper(object obj)
        {
            Utility.LogItemEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

            this._myObject = obj;
            _pConnectionPointContainer = _myObject as IConnectionPointContainer;
            if (_pConnectionPointContainer != null)
            {
                _pConnectionPoint = null;
                _pConnectionPointContainer.FindConnectionPoint(_sourceIntfGuid, out _pConnectionPoint);
                if (_pConnectionPoint != null)
                {
                    _pConnectionPoint.Advise(this, out _cookie);
                    //bool state = IsDirty();
                }
            }
        }

        //public bool IsDirty()
        //{
        //    //Type type = _myObject.GetType();
        //    //type.InvokeMember("[DispId(0x0000f024)]", System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Default, null, dispidFDirty, new object[] {});
        //    //return dispidFDirty;
        //}

        public void Dispose()
        {
            Utility.LogItemEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            Utility.LogItemEvent(LogType.Diagnostic, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
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
                    Utility.LogItemEvent(LogType.Diagnostic, (string.Format("Reference count for _pConnectionPointContainer  in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    _pConnectionPointContainer = null;
                }

                if (_myObject != null)
                {
                    // when we're disposing of instance we should also dispose of the source object 
                    int refCount = Marshal.ReleaseComObject(_myObject);
                    Utility.LogItemEvent(LogType.Diagnostic, (string.Format("Reference count for _myObject  in {0} is {1}{2}", this.GetType().Name, refCount, System.Environment.NewLine)));
                    _myObject = null;
                }

                _isDisposed = true;
            }
        }

        public void Open(ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void CustomAction(object Action, object Response, ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void Forward(object Forward, ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void Close(ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
            this.Dispose();
        }

        public void Reply(object Response, ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void ReplyAll(object Response, ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void Send(ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void Write(ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void BeforeCheckNames(ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void AttachmentAdd(Outlook.Attachment Attachment)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void AttachmentRead(Outlook.Attachment Attachment)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void BeforeAttachmentSave(Outlook.Attachment Attachment, ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void BeforeDelete(object Item, ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void AttachmentRemove(Outlook.Attachment Attachment)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void BeforeAttachmentAdd(Outlook.Attachment Attachment, ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void BeforeAttachmentPreview(Outlook.Attachment Attachment, ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void BeforeAttachmentRead(Outlook.Attachment Attachment, ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void BeforeAttachmentWriteToTempFile(Outlook.Attachment Attachment, ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void BeforeAutoSave(ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void CustomPropertyChange(string Name)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

        }

        public void PropertyChange(string Name)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

        }

        public void Read()
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

        }

        public void Unload()
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));
        }

        public void BeforeRead()
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

        }

        public void AfterWrite()
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

        }

        public void ReadComplete(ref bool Cancel)
        {
            Utility.LogItemEvent(LogType.Event, (string.Format("{0}{1}", new StackTrace().GetFrame(0).GetMethod().Name, System.Environment.NewLine)));

        }
    }
}

