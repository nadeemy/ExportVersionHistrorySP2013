using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.Web;
using System.Web.UI;
using System.IO;
using System.Web.UI.HtmlControls;
using System.Text;
using NY.ExportVersionHistory.Utilities;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;
using System.Linq;
using Microsoft.SharePoint.Taxonomy;


namespace NY.ExportVersionHistory.Layouts.NY.ExportVersionHistory
{
    public partial class ExportVersionHistory : LayoutsPageBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string listID = string.Empty;
            String[] itemIDs = null;
            try
            {
                if (Request["List"] != null && Request["ID"] != null)
                {
                    listID = Request["List"].ToString();
                    itemIDs = Request["ID"].ToString().Split(new Char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                }
                else if (Request["List"] != null && Request["View"] != null)
                {                    
                    listID = Request["List"].ToString();
                    string viewId = Request["View"].ToString();
                    Guid listGuid = new Guid(listID);
                    Guid viewGuid = new Guid(viewId);
                    SPList list = SPContext.Current.Web.Lists[listGuid];
                    SPView view = list.Views[viewGuid];
                    SPQuery query = new SPQuery(view);
                    query.RowLimit = 0;
                    itemIDs = list.GetItems(query).Cast<SPListItem>().Select(i => i.ID.ToString()).ToArray();
                }
                else if (Request["List"] != null)
                {                    
                    listID = Request["List"].ToString();
                    Guid listGuid = new Guid(listID);
                    SPList list = SPContext.Current.Web.Lists[listGuid];
                    itemIDs = list.GetItems().Cast<SPListItem>().Select(i => i.ID.ToString()).ToArray();
                }
            }
            catch (Exception ex)
            {
                LoggerUtility.LogToULS(TraceSeverity.Medium, "Error reading query string parameters: " + ex.Message);
                LoggerUtility.LogToEventViewer(EventSeverity.ErrorCritical, "Error reading query string parameters: " + ex.Message);
            }
            if (null != itemIDs && !string.IsNullOrEmpty(listID))
                ExportHistory(itemIDs, listID);
        }
        private void ExportHistory(string[] items, string listID)
        {
            SPTimeZone serverzone = SPContext.Current.Web.RegionalSettings.TimeZone; 
            StringBuilder sb = new StringBuilder();
            SPList list = SPContext.Current.Web.Lists[new Guid(listID)];
            bool isLibrary = false;
            if (list.BaseType == SPBaseType.DocumentLibrary)
                isLibrary = true;
            HtmlTable versionTable = new HtmlTable();
            versionTable.Border = 1;
            versionTable.CellPadding = 3;
            versionTable.CellSpacing = 3;
            HtmlTableRow htmlrow;
            HtmlTableCell htmlcell;

            // Add header row in HTML table
            htmlrow = new HtmlTableRow();
            htmlcell = new HtmlTableCell();
            htmlcell.InnerHtml = "Item ID";
            htmlrow.Cells.Add(htmlcell);            
            if (isLibrary)
            {
                htmlcell = new HtmlTableCell();
                htmlcell.InnerHtml = "File Name";
                htmlrow.Cells.Add(htmlcell);
                htmlcell = new HtmlTableCell();
                htmlcell.InnerHtml = "Comment";
                htmlrow.Cells.Add(htmlcell);
                htmlcell = new HtmlTableCell();
                htmlcell.InnerHtml = "Size";
                htmlrow.Cells.Add(htmlcell); 
            }
            htmlcell = new HtmlTableCell();
            htmlcell.InnerHtml = "Version No.";
            htmlrow.Cells.Add(htmlcell);
            htmlcell = new HtmlTableCell();
            htmlcell.InnerHtml = "Modified Date";
            htmlrow.Cells.Add(htmlcell);
            htmlcell = new HtmlTableCell();
            htmlcell.InnerHtml = "Modified By";
            htmlrow.Cells.Add(htmlcell);

            foreach (SPField field in list.Fields)
            {
                if (field.ShowInVersionHistory)
                {
                    htmlcell = new HtmlTableCell();
                    htmlcell.InnerHtml = field.Title;
                    htmlrow.Cells.Add(htmlcell);
                }
            }
            versionTable.Rows.Add(htmlrow);
            foreach (string item in items)
            {
                SPListItem listItem = list.GetItemById(Convert.ToInt32(item));
                SPListItemVersionCollection itemVersions = listItem.Versions;
                SPFileVersionCollection fileVersions = null;
                if (isLibrary && listItem.FileSystemObjectType == SPFileSystemObjectType.File)
                    fileVersions = listItem.File.Versions;
                for (int i = 0; i < itemVersions.Count; i++)
                {
                    SPListItemVersion currentVersion = itemVersions[i];
                    SPListItemVersion previousVersion = itemVersions.Count > i + 1 ? itemVersions[i + 1] : null;
                    htmlrow = new HtmlTableRow();
                    if (i == 0)
                    {
                        htmlcell = new HtmlTableCell();
                        htmlcell.RowSpan = itemVersions.Count;
                        htmlcell.InnerHtml = listItem.ID.ToString();
                        htmlrow.Cells.Add(htmlcell);
                    }                    
                    if (isLibrary)
                    {
                        if (i == 0)
                        {
                            htmlcell = new HtmlTableCell();
                            htmlcell.RowSpan = itemVersions.Count;
                            htmlcell.InnerHtml = listItem.File.Name;
                            htmlrow.Cells.Add(htmlcell);
                        }  

                        htmlcell = new HtmlTableCell();                       
                        HtmlTableCell sizeCell = new HtmlTableCell();
                        if (i == 0 && listItem.FileSystemObjectType == SPFileSystemObjectType.File)
                        {
                            htmlcell.InnerHtml = currentVersion.ListItem.File.CheckInComment;

                            // Implicit conversion from long to double
                            double bytes = currentVersion.ListItem.File.Length;                            
                            sizeCell.InnerHtml = Convert.ToString(Math.Round((bytes / 1024) / 1024, 2)) + " MB";
                        }
                        else
                        {
                            if (null != fileVersions)
                            {
                                foreach (SPFileVersion fileVersion in fileVersions)
                                {
                                    if (fileVersion.VersionLabel == currentVersion.VersionLabel)
                                    {
                                        htmlcell.InnerHtml = fileVersion.CheckInComment;
                                        
                                        // Implicit conversion from long to double
                                        double bytes = fileVersion.Size;
                                        sizeCell.InnerHtml = Convert.ToString(Math.Round((bytes / 1024) / 1024, 2)) + " MB";
                                        break;
                                    }
                                }
                            }
                        }
                        htmlrow.Cells.Add(htmlcell);
                        htmlrow.Cells.Add(sizeCell);                                              
                    }
                    htmlcell = new HtmlTableCell();
                    htmlcell.InnerHtml = currentVersion.VersionLabel;
                    htmlrow.Cells.Add(htmlcell);

                    htmlcell = new HtmlTableCell();
                    DateTime localDateTime = serverzone.UTCToLocalTime(currentVersion.Created);
                    htmlcell.InnerHtml = localDateTime.ToShortDateString() + " " + localDateTime.ToLongTimeString();                    
                    htmlrow.Cells.Add(htmlcell);
                    htmlcell = new HtmlTableCell();
                    htmlcell.InnerHtml = currentVersion.CreatedBy.User.Name;
                    htmlrow.Cells.Add(htmlcell);
                    foreach (SPField field in currentVersion.Fields)
                    {
                        if (field.ShowInVersionHistory)
                        {
                            htmlcell = new HtmlTableCell();
                            htmlcell.Attributes.Add("class", "textmode");
                            if (null != currentVersion[field.StaticName])
                            {
                                if (null == previousVersion)
                                {
                                    htmlcell.InnerHtml = GetFieldValue(field, currentVersion);
                                }
                                else
                                {
                                    if (null != previousVersion[field.StaticName] && currentVersion[field.StaticName].ToString().Equals(previousVersion[field.StaticName].ToString()))
                                    {
                                        htmlcell.InnerHtml = string.Empty;
                                    }

                                    else
                                    {
                                        htmlcell.InnerHtml = GetFieldValue(field, currentVersion);
                                    }
                                }
                            }
                            else
                            {
                                htmlcell.InnerHtml = string.Empty;
                            }
                            htmlrow.Cells.Add(htmlcell);
                        }
                    }
                    versionTable.Rows.Add(htmlrow);
                }
            }

            ExportTableToExcel(versionTable, list.Title);

        }
        private void ExportTableToExcel(HtmlTable table, string title)
        {
            using (StringWriter stringWriter = new StringWriter())
            {
                using (HtmlTextWriter textWriter = new HtmlTextWriter(stringWriter))
                {
                    table.RenderControl(textWriter);
                    Response.Clear();
                    Response.ContentType = "application/vnd.ms-excel";
                    Response.Charset = "65001";
                    byte[] b = new byte[] { 0xEF, 0xBB, 0xBF };
                    Response.BinaryWrite(b);
                    Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", title + ".xls"));

                    // style to format numbers to string
                    string style = @"<style> .textmode { mso-number-format:\@; } </style>";
                    Response.Write(style);
                    Response.Write(stringWriter.ToString());
                    Response.End();
                }
            }
        }
        private string GetFieldValue(SPField field, SPListItemVersion version)
        {
            string fieldValue = string.Empty;
            SPFieldType fieldType = field.Type;
            switch (fieldType)
            {
                case SPFieldType.Lookup:
                    SPFieldLookup newField = (SPFieldLookup)field;
                    fieldValue = newField.GetFieldValueAsText(version[field.StaticName]);
                    break;
                case SPFieldType.User:
                    SPFieldUser newUser = (SPFieldUser)field;
                    fieldValue = newUser.GetFieldValueAsText(version[field.StaticName]);
                    break;
                case SPFieldType.ModStat:
                    SPFieldModStat modStat = (SPFieldModStat)field;
                    fieldValue = modStat.GetFieldValueAsText(version[field.StaticName]);
                    break;
                case SPFieldType.URL:
                    SPFieldUrl urlField = (SPFieldUrl)field;
                    fieldValue = urlField.GetFieldValueAsHtml(version[field.StaticName]);
                    break;
                case SPFieldType.DateTime:
                    SPFieldDateTime newDateField = (SPFieldDateTime)field;
                    if (!string.IsNullOrEmpty(newDateField.GetFieldValueAsText(version[field.StaticName])))
                    {
                        if (newDateField.DisplayFormat == SPDateTimeFieldFormatType.DateTime)
                        {
                            fieldValue = DateTime.Parse(newDateField.GetFieldValueAsText(version[field.StaticName])).ToString();
                        }
                        else
                        {
                            fieldValue = DateTime.Parse(newDateField.GetFieldValueAsText(version[field.StaticName])).ToShortDateString();
                        }
                    }
                    break;
                case SPFieldType.Invalid:

                    // http://sharepointnadeem.blogspot.com/2013/09/sharepoint-spfieldtype-is-invalid-for.html
                    if (field.TypeAsString.Equals("TaxonomyFieldType") || field.TypeAsString.Equals("TaxonomyFieldTypeMulti"))
                    {
                        TaxonomyField taxonomyField = field as TaxonomyField;
                        fieldValue = taxonomyField.GetFieldValueAsText(version[field.StaticName]);
                    }
                    else
                    {
                        fieldValue = version[field.StaticName].ToString();
                    }
                    break;
                default:
                    fieldValue = version[field.StaticName].ToString();
                    break;
            }

            return fieldValue;
        }
    }
}
