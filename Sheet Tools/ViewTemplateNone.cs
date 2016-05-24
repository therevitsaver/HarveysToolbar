using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System.Net;
using System.IO;

namespace SheetTools
{
    [Transaction(TransactionMode.Manual)]

    public class ViewTemplateNone : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects


            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                
                View vw = doc.ActiveView;
                Transaction tr = new Transaction(doc, "Clears the View Template Parameter");
                tr.Start();
                Parameter par = vw.LookupParameter("View Template");
                par.Set(new ElementId(-1));
                tr.Commit();


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
