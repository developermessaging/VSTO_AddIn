using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace VSTO_AddIn
{
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), Guid("0006302B-0000-0000-C000-000000000046")]
    internal interface ItemEvents10Ex
    {
        [DispId(0x0000fc8e)]
        void BeforePrint([MarshalAs(UnmanagedType.VariantBool), In, Out] ref bool unnamed1, [In] object reserved);

        [DispId(0x0000fc8f)]
        void ReadComplete([MarshalAs(UnmanagedType.VariantBool), In, Out] ref bool Cancel);

        [DispId(0x0000f003)]
        void Open([In, Out] ref object Cancel);

        [DispId(0x0000f006)]
        void CustomAction([In] object Action, [In] object Response, [In, Out] ref object Cancel);

        [DispId(0x0000f008)]
        void CustomPropertyChange([In] string Name);

        [DispId(0x0000f468)]
        void Forward([In] object Forward, [In, Out] ref object Cancel);

        [DispId(0x0000f004)]
        void Close([In, Out] ref object Cancel);

        [DispId(0x0000f009)]
        void PropertyChange([In] string Name);

        [DispId(0x0000f001)]
        void Read();

        [DispId(0x0000f466)]
        void Reply([In] object Response, [In, Out] ref object Cancel);

        [DispId(0x0000f467)]
        void ReplyAll([In] object Response, [In, Out] ref object Cancel);

        [DispId(0x0000f005)]
        void Send([In, Out] ref object Cancel);

        [DispId(0x0000f002)]
        void Write([In, Out] ref object Cancel);

        [DispId(0x0000f00a)]
        void BeforeCheckNames([In, Out] ref object Cancel);

        [DispId(0x0000f00b)]
        void AttachmentAdd([In] object Attachment);

        [DispId(0x0000f00c)]
        void AttachmentRead([In] object Attachment);

        [DispId(0x0000f00d)]
        void BeforeAttachmentSave([In] object Attachment, [In, Out] ref object Cancel);

        [DispId(0x0000fa75)]
        void BeforeDelete([In] object Item, [In, Out] ref object Cancel);

        [DispId(0x0000fbae)]
        void AttachmentRemove([In] object Attachment);

        [DispId(0x0000fbb0)]
        void BeforeAttachmentAdd([In] object Attachment, [In, Out] ref object Cancel);

        [DispId(0x0000fbaf)]
        void BeforeAttachmentPreview([In] object Attachment, [In, Out] ref object Cancel);

        [DispId(0x0000fbab)]
        void BeforeAttachmentRead([In] object Attachment, [In, Out] ref object Cancel);

        [DispId(0x0000fbb2)]
        void BeforeAttachmentWriteToTempFile([In] object Attachment, [In, Out] ref object Cancel);

        [DispId(0x0000fbad)]
        void Unload();

        [DispId(0x0000fc02)]
        void BeforeAutoSave([In, Out] ref object Cancel);

        [DispId(0x0000fc8c)]
        void BeforeRead();

        [DispId(0x0000fc8d)]
        void AfterWrite();
    }

}
