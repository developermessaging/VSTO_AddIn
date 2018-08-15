using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace VSTO_AddIn
{
	#region Not relevant to this proto

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

	#endregion Not relevant to this proto

	[ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(ItemEvents10Ex)), Guid("2CA74A5A-FA89-4769-9B97-6185370709D0")]
	public class MailItemEventWrapper : ItemEvents10Ex, IDisposable
	{
		private readonly Guid _sourceIntfGuid = new Guid("0006302B-0000-0000-C000-000000000046");
		private Int32 _cookie;
		private IConnectionPoint _pConnectionPoint;

		public MailItemEventWrapper(Outlook.MailItem mailItem)
		{

			IConnectionPointContainer pConnectionPointContainer = mailItem as IConnectionPointContainer;
			if (pConnectionPointContainer != null)
			{
				_pConnectionPoint = null;
				pConnectionPointContainer.FindConnectionPoint(_sourceIntfGuid, out _pConnectionPoint);
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
			if (disposing)
			{
				if (_cookie != -1 && _pConnectionPoint != null)
				{
					_pConnectionPoint.Unadvise(_cookie);
					_pConnectionPoint = null;
					_cookie = -1;
				}
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
            Globals.ThisAddIn.replyAllAttempt++;
            System.Windows.Forms.MessageBox.Show(String.Format("Reply All attempt #{0}", Globals.ThisAddIn.replyAllAttempt));
            Cancel = true;
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
