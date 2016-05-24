using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;

using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;

namespace RenumberingTools
{
    [Transaction(TransactionMode.Manual)]
    public class RenumberTagRight : IExternalCommand
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
                try
                {
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("Lock View");
                        if (doc.ActiveView is View3D)
                        {
                            View3D view3d = doc.ActiveView as View3D;
                            bool isLocked = view3d.IsLocked;
                            if (isLocked == false)
                            {
                                view3d.SaveOrientationAndLock();
                            }
                        }
                        tx.Commit();
                    }
                }
                catch (System.NullReferenceException)
                {

                }
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

                            if (eleid.Category.Name == "Pipe Fittings" || eleid.Category.Name == "Air Terminals" ||
                                eleid.Category.Name == "Mechanical Equipment" || eleid.Category.Name == "Duct Accessories" ||
                                eleid.Category.Name == "Pipe Accessories" || eleid.Category.Name == "Plumbing Fixtures" || eleid.Category.Name == "Duct Fittings")
                            {
                                XYZ centerPnt = new XYZ(0, 0, 0);
                                LocationPoint elemLoc = eleid.Location as LocationPoint;
                                XYZ elemLocPnt = elemLoc.Point;

                                IndependentTag pipeFittingTag = doc.Create.NewTag(view, eleid, true,
                                                                       TagMode.TM_ADDBY_CATEGORY,
                                                                       TagOrientation.Horizontal,
                                                                       elemLocPnt);

                                pipeFittingTag.LeaderEndCondition = LeaderEndCondition.Free;
                                pipeFittingTag.LeaderEnd = elemLocPnt;
                                XYZ headerPnt = elemLocPnt + new XYZ(0.9, 0.0, -0.45);
                                //pipeFittingTag.LeaderElbow = headerPnt;
                                pipeFittingTag.TagHeadPosition = headerPnt;
                            }
                            else if (eleid.Category.Name == "Pipes" || eleid.Category.Name == "Ducts")
                            {
                                LocationCurve pipeLoc = eleid.Location as LocationCurve;
                                XYZ pipeStart = pipeLoc.Curve.GetEndPoint(0);
                                XYZ pipeEnd = pipeLoc.Curve.GetEndPoint(1);
                                XYZ pipeMid = pipeLoc.Curve.Evaluate(0.5, true);

                                IndependentTag pipeTag = doc.Create.NewTag(view, eleid, true,
                                                                               TagMode.TM_ADDBY_CATEGORY,
                                                                               TagOrientation.Horizontal,
                                                                               pipeMid);

                                pipeTag.LeaderEndCondition = LeaderEndCondition.Free;
                                pipeTag.LeaderEnd = pipeMid;
                                XYZ headerPnt = pipeMid + new XYZ(0.9, 0.0, -0.45);
                                //pipeTag.LeaderElbow = headerPnt;
                                pipeTag.TagHeadPosition = headerPnt;
                            }
                            else if (eleid.Category.Name == "Fabrication Parts")
                            {
                                FabricationPart fabPart = eleid as FabricationPart;
                                LocationCurve fabLoc = fabPart.Location as LocationCurve;

                                Connector conn1 = GetPrimaryConnector(fabPart.ConnectorManager);
                                Connector conn2 = GetSecondaryConnector(fabPart.ConnectorManager);
                                XYZ fabStart = conn1.Origin;
                                XYZ fabEnd = conn2.Origin;
                                XYZ fabMid = new XYZ(((fabStart.X + fabEnd.X) / 2), ((fabStart.Y + fabEnd.Y) / 2), ((fabStart.Z + fabEnd.Z) / 2));
                                IndependentTag fabTag = doc.Create.NewTag(view, eleid, true, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, fabMid);

                                fabTag.LeaderEndCondition = LeaderEndCondition.Free;
                                fabTag.LeaderEnd = fabMid;
                                XYZ headerPnt = fabMid + new XYZ(0.9, 0.0, -0.45);
                                //pipeTag.LeaderElbow = headerPnt;
                                fabTag.TagHeadPosition = headerPnt;
                            }

                            doc.Regenerate();
                        }
                        tx2.Commit();
                    }
                    TaskDialog.Show("Harvey's Renumbering Tool", " Done! Total of pieces = " + ids.Count.ToString());
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
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

        Connector GetPrimaryConnector(ConnectorManager cm)
        {
            foreach (Connector cn in cm.Connectors)
            {
                MEPConnectorInfo info = cn.GetMEPConnectorInfo();
                if (info.IsPrimary)
                    return cn;
            }
            return null;
        }
        Connector GetSecondaryConnector(ConnectorManager cm)
        {
            foreach (Connector cn in cm.Connectors)
            {
                MEPConnectorInfo info = cn.GetMEPConnectorInfo();
                if (info.IsSecondary)
                    return cn;
            }
            return null;
        }
    }
}

