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

    public class ViewTemplatePlumbingDetailView : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects


            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                // see http://wp.me/p2X0gy-2p for more info on how this works
                View viewTemplate = (from v in new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                                     where v.IsTemplate == true && v.Name == "Harvey's Plumbing Detail View"
                                     select v)
                    .First();

                using (Transaction t = new Transaction(doc, "Set View Template"))
                {
                    t.Start();
                    doc.ActiveView.ViewTemplateId = viewTemplate.Id;
                    t.Commit();
                }
         
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
