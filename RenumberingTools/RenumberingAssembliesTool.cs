using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;

namespace RenumberingTools
{
    [Transaction(TransactionMode.Manual)]

    public class RenumberingAssembliesTool : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects

            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            Autodesk.Revit.DB.View view = doc.ActiveView;
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;


            try
            {

                ISelectionFilter selFilter = new assemblySelectionFilter();
                List<ElementId> ids = new List<ElementId>();
                Selection sel = uiDoc.Selection;
                ICollection<ElementId> selIds = sel.GetElementIds();

                if (selIds.Count > 0)
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
                        refs = sel.PickObjects(ObjectType.Element, selFilter, "Please select all assemblies to renumber.");
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
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("Renumber Assemblies");

                        for (int assNum = 1; assNum < ids.Count + 1; assNum++)
                        {
                            ElementId elid = uiDoc.Selection.PickObject(ObjectType.Element, selFilter, "Select Next Assembly =" + assNum.ToString()).ElementId;
                            AssemblyInstance chosenAssemblyId = (AssemblyInstance)doc.GetElement(elid);
                            string assemblyName = chosenAssemblyId.Name;

                            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Sheets);
                            FilteredElementCollector collector1 = new FilteredElementCollector(doc);
                            collector1.WherePasses(filter);
                            ParameterValueProvider provider = new ParameterValueProvider(new ElementId(BuiltInParameter.SHEET_ASSEMBLY_NAME));
                            FilterStringEquals evaluator = new FilterStringEquals();
                            FilterStringRule rule = new FilterStringRule(provider, evaluator, assemblyName, false);
                            ElementParameterFilter paramFilter = new ElementParameterFilter(rule);
                            collector1.WherePasses(paramFilter);
                            IList<Element> sheets = collector1.ToElements();

                            ElementId assemblySheetId = ((ViewSheet)(sheets[0])).Id;
                            ViewSheet sheet = (ViewSheet)doc.GetElement(assemblySheetId);

                            //------------------------------Continue from Here and at 17 minutes on DevTV - Introduction to Revit 2015 Programming -- Part 2--------------------------------

                            int index = assemblyName.IndexOf("_");
                            string assemblyNumberOld = assemblyName.Substring(0, index);
                            string assemblyNameEdited = assemblyName.Remove(0, index + 1);
                            string numberFull = "AS" + assNum.ToString();
                            string numberFullwUnderscore = numberFull + "_";
                            string newAssemblyName = numberFullwUnderscore + assemblyNameEdited;

                            ElementCategoryFilter sheetFilter = new ElementCategoryFilter(BuiltInCategory.OST_Sheets);
                            FilteredElementCollector sheetCollector = new FilteredElementCollector(doc);
                            sheetCollector.WherePasses(sheetFilter);
                            ParameterValueProvider sheetProvider = new ParameterValueProvider(new ElementId(BuiltInParameter.SHEET_NUMBER));
                            FilterStringEquals sheetEvaluator = new FilterStringEquals();
                            FilterStringRule sheetRule = new FilterStringRule(sheetProvider, sheetEvaluator, numberFull, false);
                            ElementParameterFilter sheetFilter2 = new ElementParameterFilter(sheetRule);
                            sheetCollector.WherePasses(sheetFilter2);
                            IList<Element> sheetsList = sheetCollector.ToElements();

                            if (sheetsList.Count > 0)
                            {
                                ((ViewSheet)(sheetsList[0])).SheetNumber = sheet.SheetNumber;
                                sheet.SheetNumber = numberFull;
                            }
                            else
                            {
                                sheet.SheetNumber = numberFull;
                                sheet.LookupParameter("Sheet Number").SetValueString(numberFull);
                            }

                            //------------------------------------------------------------------------------------------------------------------------
                            ElementCategoryFilter assemblyFilter = new ElementCategoryFilter(BuiltInCategory.OST_Assemblies);
                            FilteredElementCollector assemblyCollector = new FilteredElementCollector(doc);
                            assemblyCollector.WherePasses(assemblyFilter);
                            ParameterValueProvider assemblyProvider = new ParameterValueProvider(new ElementId(BuiltInParameter.ALL_MODEL_TYPE_NAME));
                            FilterStringEquals assemblyEvaluator = new FilterStringEquals();
                            FilterStringRule assemblyRule = new FilterStringRule(assemblyProvider, assemblyEvaluator, newAssemblyName, false);
                            ElementParameterFilter assemblyFilter2 = new ElementParameterFilter(assemblyRule);
                            assemblyCollector.WherePasses(assemblyFilter2);
                            IList<Element> assembliesList = assemblyCollector.ToElements();

                            if (assembliesList.Count > 0)
                            {
                                ((AssemblyInstance)(assembliesList[0])).AssemblyTypeName = "AS" + (assNum + 900).ToString() + "_" + assemblyNameEdited;
                                chosenAssemblyId.AssemblyTypeName = newAssemblyName;
                            }
                            else
                            {
                                chosenAssemblyId.AssemblyTypeName = newAssemblyName;
                            }

                            //------------------------------------------------------------------------------------------------------------------------
                        }

                        tx.Commit();
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {

                }
                catch (Autodesk.Revit.Exceptions.ApplicationException)
                {

                }
           


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
        public class assemblySelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element el)
            {
                if (el.Category.Name == "Assemblies")
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
