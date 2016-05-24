using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;

namespace RenumberingTools
{
    [Transaction(TransactionMode.Manual)]
    public class PipeAndFittingRenumberingTool : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects


            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;
                Autodesk.Revit.DB.View view = doc.ActiveView;
                Autodesk.Revit.ApplicationServices.Application app = doc.Application;
                ISelectionFilter selFilter = new HarveysSelectionFilter();
                
                List<ElementId> ids = new List<ElementId>();
                Autodesk.Revit.UI.Selection.Selection sel = uiDoc.Selection;
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
                        refs = sel.PickObjects(ObjectType.Element, selFilter, "Please select all elements to renumber.");
                        ids = new List<ElementId>(refs.Select<Reference, ElementId>(r => r.ElementId));
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {

                    }

                }

                foreach (ElementId id in ids)
                {
                    Element el = doc.GetElement(id);
                }
                try
                {
                    using (Transaction tx2 = new Transaction(doc))
                    {
                        tx2.Start("Quick Renumbering");

                        for (int i = 1; i < ids.Count + 1; i++)
                        {
                            ElementId elid = uiDoc.Selection.PickObject(ObjectType.Element, selFilter, "Select Next Element =" + i.ToString()).ElementId;
                            Element eleid = doc.GetElement(elid);
                            eleid.LookupParameter("Harvey's Part Number").Set(i);
                            //Change parameter to your own also remove .tostring() if integer
                        }
                        tx2.Commit();
                    }
                    TaskDialog.Show("Harvey's Renumbering Tool", " Done! Total of pieces = " + ids.Count.ToString());
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {

                }
                catch (Autodesk.Revit.Exceptions.ApplicationException)
                {

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
          public class HarveysSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element el)
            {
                if (el.Category.Name == "Ducts")
                {
                    return true;
                }
                else if (el.Category.Name == "Duct Fittings")
                {
                    return true;
                }
                else if (el.Category.Name == "Fabrication Parts")
                {
                    return true;
                }
                else if (el.Category.Name == "Pipes")
                {
                    return true;
                }
                else if (el.Category.Name == "Pipe Fittings")
                {
                    return true;
                }
                else if (el.Category.Name == "Plumbing Fixtures")
                {
                    return true;
                }
                else if (el.Category.Name == "Pipe Accessories")
                {
                    return true;
                }
                else if (el.Category.Name == "Duct Accessories")
                {
                    return true;
                }
                else if (el.Category.Name == "Mechanical Equipment")
                {
                    return true;
                }
                else if (el.Category.Name == "Air Terminals")
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

            public bool AllowReference(Reference refer, XYZ point)
            {
                return false;
            }
        }
    }

