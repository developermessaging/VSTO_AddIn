using System;
using System.Text;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace VSTO_AddIn
{
	public partial class ThisAddIn
	{
        public int replyAllAttempt = 0;
		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			Application.ItemLoad += new Outlook.ApplicationEvents_11_ItemLoadEventHandler(ItemLoad);
		}

		private void ItemLoad(object item)
		{
			Outlook.MailItem mailItem = item as Outlook.MailItem;
			if (mailItem != null)
			{
				new MailItemEventWrapper(mailItem);
			}
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
