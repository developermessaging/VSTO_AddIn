using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace VSTO_AddIn
{
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(ItemEvents10Ex)), Guid("82e1b877-a7d8-48de-9b24-5985482b5a14")]
    public class ItemEventWrapper : ItemEvents10Ex, IDisposable
    {
        private readonly Guid _sourceIntfGuid = new Guid("0006302B-0000-0000-C000-000000000046");
        private Int32 _cookie;
        private IConnectionPoint _pConnectionPoint = null;
        private IConnectionPointContainer _pConnectionPointContainer = null;
        private bool _isDisposed = false;
        object _myObject = null;

        public ItemEventWrapper(object obj)
        {
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
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
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
                    System.Diagnostics.Trace.WriteLine(string.Format("Reference count for _pConnectionPointContainer is {0}", refCount));
                    _pConnectionPointContainer = null;
                }

                if (_myObject != null)
                {
                    // when we're disposing of instance we should also dispose of the source object 
                    int refCount = Marshal.ReleaseComObject(_myObject);
                    System.Diagnostics.Trace.WriteLine(string.Format("Reference count for _myObject is {0}", refCount));
                    _myObject = null;
                }

                _isDisposed = true;
            }
        }

        [DispId(0x0000fc8f)]
        public void ReadComplete([MarshalAs(UnmanagedType.VariantBool), In, Out] ref bool cancel)
        {
        }

        [DispId(0x0000f004)]
        public void Close([In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000f003)]
        public void Open([In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000f006)]
        public void CustomAction([In] object Action, [In] object Response, [In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000f008)]
        public void CustomPropertyChange([In] string Name)
        {
        }

        [DispId(0x0000f468)]
        public void Forward([In] object Forward, [In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000f009)]
        public void PropertyChange([In] string Name)
        {
        }

        [DispId(0x0000f001)]
        public void Read()
        {
        }

        [DispId(0x0000f466)]
        public void Reply([In] object Response, [In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000f467)]
        public void ReplyAll([In] object Response, [In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000f005)]
        public void Send([In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000f002)]
        public void Write([In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000f00a)]
        public void BeforeCheckNames([In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000f00b)]
        public void AttachmentAdd([In] object Attachment)
        {
        }

        [DispId(0x0000f00c)]
        public void AttachmentRead([In] object Attachment)
        {
        }

        [DispId(0x0000f00d)]
        public void BeforeAttachmentSave([In] object Attachment, [In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000fa75)]
        public void BeforeDelete([In] object Item, [In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000fbae)]
        public void AttachmentRemove([In] object Attachment)
        {
        }

        [DispId(0x0000fbb0)]
        public void BeforeAttachmentAdd([In] object Attachment, [In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000fbaf)]
        public void BeforeAttachmentPreview([In] object Attachment, [In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000fbab)]
        public void BeforeAttachmentRead([In] object Attachment, [In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000fbb2)]
        public void BeforeAttachmentWriteToTempFile([In] object Attachment, [In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000fbad)]
        public void Unload()
        {
        }

        [DispId(0x0000fc02)]
        public void BeforeAutoSave([In, Out] ref object Cancel)
        {
        }

        [DispId(0x0000fc8c)]
        public void BeforeRead()
        {
        }

        [DispId(0x0000fc8d)]
        public void AfterWrite()
        {
        }

        [DispId(0x0000fc8e)]
        public void BeforePrint([MarshalAs(UnmanagedType.VariantBool), In, Out] ref bool unnamed1, [In] object reserved)
        {
        }

    }
}

