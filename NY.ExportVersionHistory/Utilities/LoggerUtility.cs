using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;

namespace NY.ExportVersionHistory.Utilities
{
    public class LoggerUtility
    {
        private static readonly SPDiagnosticsService service;
        private static readonly SPDiagnosticsCategory category;
        static LoggerUtility()
        {
            service = SPDiagnosticsService.Local;
            category = service.Areas["SharePoint Foundation"].Categories["General"];
        }
        public static void LogToULS(TraceSeverity severity,string message)
        {
            service.WriteTrace(101, category, severity, message, null);
        }

        public static void LogToEventViewer(EventSeverity severity, string message)
        {
            service.WriteEvent(101, category, severity, message, null);
        } 
    }
}
