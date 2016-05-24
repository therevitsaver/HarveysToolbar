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

    public class ViewSectionBoxOn : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects


            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;


                using (Transaction t = new Transaction(doc, "Turn Section Box Off"))
                {
                    t.Start();
                    ShowHideSectionBox(uiDoc, true);
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
        private void ShowHideSectionBox(UIDocument document, bool active)
        {
            if (document.ActiveView is View3D)
            {
                View3D view3d = document.ActiveView as View3D;
                view3d.IsSectionBoxActive = active;
            }
        }
    }
}
