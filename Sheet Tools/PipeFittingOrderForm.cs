using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;



namespace SheetTools
{
    [Transaction(TransactionMode.Manual)]
    public class PipeFittingOrderForm : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects


            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                FilteredElementCollector col = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));

                ViewScheduleExportOptions opt = new ViewScheduleExportOptions();

                foreach (ViewSchedule vs in col)
                {
                    Directory.CreateDirectory(@"C:\Users\Brian\Desktop\Schedules");

                    if (vs.ViewName == "Plumbing - Pipe Fitting Order Sheet")
                    {
                        vs.Export(@"C:\Users\Brian\Desktop\Schedules",
                               "Plumbing - Pipe Fitting Order Sheet" + ".csv", opt);
                    }
                }

                string mySheet = @"C:\Users\Brian\Desktop\Schedules\Plumbing - Pipe Fitting Order Sheet.csv";

                var excelApp = new Excel.Application();
                excelApp.Visible = true;

                Excel.Workbooks books = excelApp.Workbooks;

                Excel.Workbook sheet = books.Open(mySheet);

                Excel.Worksheet ws = excelApp.ActiveWorkbook.Worksheets[1];
                ws.Columns.AutoFit();
                ws.Rows.AutoFit();

                // Assuming that everything went right return Result.Succeeded

                return Result.Succeeded;
            }



            // This is where we "catch" potential error and define how to deal with them
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // If user decided to cancel the operation return Result.Cancelled
                return Result.Cancelled;
            }

            catch (System.Exception ex)
            {
                // If something went wrong return Result.Failed
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}

