using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
namespace VSTO_AddIn
{
    enum LogType
    {
        Information,
        Warning,
        Error,
        Event,
        Generic,
        Diagnostic
    }
    class Utility
    {
        internal static int ReleaseObject(object obj)
        {
            int iReleased = Marshal.ReleaseComObject(obj);
            if (0 < iReleased)
            {
                return iReleased;
            }
            else
            {
                obj = null;
            }
            return iReleased;
        }

        internal static void LogItemEvent(LogType logType, string message)
        {
            if (ControlPanel.ItemLoggingEnabled)
            {

                Log(logType, "Item", message);
            }
        }

        internal static void LogApplicationEvent(LogType logType, string message)
        {
            if (ControlPanel.ApplicationLoggingEnabled)
            {

                Log(logType, "Application", message);
            }
        }

        internal static void LogExplorerEvent(LogType logType, string message)
        {
            if (ControlPanel.ExplorerLoggingEnabled)
            {

                Log(logType, "Explorer", message);
            }
        }

        internal static void LogExplorersEvent(LogType logType, string message)
        {
            if (ControlPanel.ExplorersLoggingEnabled)
            {

                Log(logType, "Explorers", message);
            }
        }

        internal static void LogInspectorsEvent(LogType logType, string message)
        {
            if (ControlPanel.InspectorsLoggingEnabled)
            {

                Log(logType, "Inspectors", message);
            }
        }

        internal static void LogInspectorEvent(LogType logType, string message)
        {
            if (ControlPanel.InspectorLoggingEnabled)
            {

                Log(logType, "Inspector", message);
            }
        }

        internal static void LogFoldersEvent(LogType logType, string message)
        {
            if (ControlPanel.FoldersLoggingEnabled)
            {

                Log(logType, "Folders", message);
            }
        }

        internal static void LogFolderEvent(LogType logType, string message)
        {
            if (ControlPanel.FolderLoggingEnabled)
            {

                Log(logType, "Folder", message);
            }
        }

        internal static void LogItemsEvent(LogType logType, string message)
        {
            if (ControlPanel.ItemsLoggingEnabled)
            {

                Log(logType, "Items", message);
            }
        }

        internal static void LogNameSpaceEvent(LogType logType, string message)
        {
            if (ControlPanel.NameSpaceLoggingEnabled)
            {

                Log(logType, "NameSpace", message);
            }
        }

        internal static void LogStoresEvent(LogType logType, string message)
        {
            if (ControlPanel.StoresLoggingEnabled)
            {

                Log(logType, "Stores", message);
            }
        }

        static void Log(LogType logType, string type, string message)
        {
            switch (logType)
            {
                case LogType.Event:
                    {
                        ControlPanel.AddText(string.Format("{0} {1} {2} {3} {4}", System.DateTime.Now.ToShortDateString(), System.DateTime.Now.ToShortTimeString(), type, "Event", message));
                        break;
                    }
                case LogType.Diagnostic:
                    {
                        ControlPanel.AddText(string.Format("{0} {1} {2} {3} {4}", System.DateTime.Now.ToShortDateString(), System.DateTime.Now.ToShortTimeString(), type, "Diagnostic", message));
                        break;
                    }
                case LogType.Error:
                    {
                        ControlPanel.AddText(string.Format("{0} {1} {2} {3} {4}", System.DateTime.Now.ToShortDateString(), System.DateTime.Now.ToShortTimeString(), type, "Error", message));
                        break;
                    }
                case LogType.Information:
                    {
                        ControlPanel.AddText(string.Format("{0} {1} {2} {3} {4}", System.DateTime.Now.ToShortDateString(), System.DateTime.Now.ToShortTimeString(), type, "Information", message));
                        break;
                    }
                case LogType.Warning:
                    {
                        ControlPanel.AddText(string.Format("{0} {1} {2} {3} {4}", System.DateTime.Now.ToShortDateString(), System.DateTime.Now.ToShortTimeString(), type, "Warning", message));
                        break;
                    }
                default:
                    {
                        ControlPanel.AddText(string.Format("{0} {1} {2} {3} {4}", System.DateTime.Now.ToShortDateString(), System.DateTime.Now.ToShortTimeString(), type, "Generic", message));
                        break;
                    }
            };
        }
    }
}
