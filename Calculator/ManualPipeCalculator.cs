using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;

namespace Calculator
{
    [Transaction(TransactionMode.Manual)]

    public class ManualPipeCalculator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects

            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;


            try
            {
                ISelectionFilter selFilter = new pipeSelectionFilter();

                List<ElementId> ids = new List<ElementId>();
                Selection sel = uiDoc.Selection;

                ICollection<ElementId> selIds = sel.GetElementIds();

                if (0 < selIds.Count)
                {
                    foreach (ElementId id in selIds)
                    {
                        ids.Add(id);
                    }
                }

                if (0 == selIds.Count)
                {
                    IList<Reference> refs = null;

                    try
                    {
                        refs = sel.PickObjects(ObjectType.Element, selFilter, "Please Select Pipes to Count Total Length.");
                    }
                    catch (Autodesk.Revit.Exceptions
                    .OperationCanceledException)
                    {
                        //  return Result.Cancelled;
                    }
                    ids = new List<ElementId>(
                    refs.Select<Reference, ElementId>(
                    r => r.ElementId));
                }
                double value = 0;

                foreach (ElementId id in ids)
                {
                    Element el = doc.GetElement(id);
                    value = value + el.LookupParameter("Length").AsDouble();

                }
                Units units = doc.GetUnits();

                string asvaluestring = UnitFormatUtils.Format(units, UnitType.UT_Length, value, false, false);

                TaskDialog.Show("The total amount of pipe selected is: ", asvaluestring);
            
                // Return a message window that displays total length to user

            // Assuming that everything went right return Result.Succeeded

            return Result.Succeeded;
            }

            // This is where we "catch" potential error and define how to deal with them
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // If user decided to cancel the operation return Result.Cancelled
                return Result.Cancelled;
            }

            catch (Exception ex)
            {
                // If something went wrong return Result.Failed
                message = ex.Message;
                return Result.Failed;
            }
        }
        //This is how you create a single filter.                
        public class pipeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element el)
            {
                if (el.Category.Name == "Pipes")
                {
                    return true;
                }
                return false;
            }

            public bool AllowReference(Reference refer, XYZ point)
            {
                return false;
            }
        }

        public class pipeFittingSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element el)
            {
                if (el.Category.Name == "Pipe Fittings")
                {
                    return true;
                }
                return false;
            }

            public bool AllowReference(Reference refer, XYZ point)
            {
                return false;
            }
        }
    }
}
