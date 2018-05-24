using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Web.Controllers
{
    public class HomeController : Controller
    {
        private string destination = "~/Export";

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExportExcell()
        {
            string destination = Server.MapPath(this.destination + "/" + Guid.NewGuid() + ".xlsx");

            CreateFileExport(destination);
            ExportFollowSheet(GetMockupDataset(), destination, "ExportOne");
            ExportFollowSheet(GetMockupDataset(), destination, "ExportTwo");
            ExportFollowSheet(GetMockupDataset(), destination, "ExportThree");

            return Json(new { msg = "Ok" }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [NonAction]
        private DataSet GetMockupDataset()
        {
            var ds = new DataSet();

            for (int i = 0; i < 3; i++)
            {
                ds.Tables.Add(GetMockupDataTable().Copy());
            }

            return ds;
        }

        [NonAction]
        private DataTable GetMockupDataTable()
        {
            var dt = new DataTable();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Description", typeof(string));

            dt.Rows.Add("1", "One", "One");
            dt.Rows.Add("2", "One", "One");
            dt.Rows.Add("3", "One", "One");
            dt.Rows.Add("4", "One", "One");
            dt.Rows.Add("5", "One", "One");

            return dt;
        }

        [NonAction]
        void CreateFileExport(string destination)
        {
            using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                workbook.Clone();
            }
        }

        [NonAction]
        private void ExportFollowSheet(DataSet ds, string destination, string nameSheet)
        {
            using (var workbook = SpreadsheetDocument.Open(destination, true))
            {
                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;
                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = nameSheet };
                sheets.Append(sheet);
                foreach (System.Data.DataTable table in ds.Tables)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                    List<String> columns = new List<string>();
                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }


                    sheetData.AppendChild(headerRow);

                    foreach (System.Data.DataRow dsrow in table.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        foreach (String col in columns)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }
                }

                workbook.Clone();
            }
        }
    }
}