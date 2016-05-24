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

    public class RenumberAssembliesTwo : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects

            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;


            try
            {
                ISelectionFilter selFilter = new assemblySelectionFilter();

                List<ElementId> ids = new List<ElementId>();

                try
                {
                    while (true)
                    {
                        Reference reference = uiDoc.Selection.PickObject(ObjectType.Element, selFilter,
                                                                         "Pick Elements in the disired order and hit ESC to complete.");
                        ids.Add(reference.ElementId);


                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {

                }

                //--------------------------------------------------------------------------------------------------------------------------------------------------------------            
                ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Assemblies);

                // Apply the filter to the elements in the active document
                // Use shortcut WhereElementIsNotElementType() to find view instances only
                FilteredElementCollector assemblyCollector = new FilteredElementCollector(doc);
                IList<Element> assemblies =
                assemblyCollector.WherePasses(filter).WhereElementIsNotElementType().ToElements();
                String allAssemblyNames = "The assemblies in the current document are:\n";

                //Obtain all existing assembly names and list them.           

                //------------------------------------------------------------------------------------------------------------------------------

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Rename Assembly and Create Sheets/Views");

                    foreach (ElementId id in ids)
                    {

                        string count = "";
                        foreach (Element e in assemblies)
                        {
                            allAssemblyNames += e.Name + "\n";
                        }
                        //Test to find which assembly number has not been chosen, and use that assembly number if it hasn't been chosen.
                        if (!allAssemblyNames.Contains("AS1_"))
                        {
                            count = "1";
                        }
                        else if (!allAssemblyNames.Contains("AS2_"))
                        {
                            count = "2";
                        }
                        else if (!allAssemblyNames.Contains("AS3_"))
                        {
                            count = "3";
                        }
                        else if (!allAssemblyNames.Contains("AS4_"))
                        {
                            count = "4";
                        }
                        else if (!allAssemblyNames.Contains("AS5_"))
                        {
                            count = "5";
                        }
                        else if (!allAssemblyNames.Contains("AS6_"))
                        {
                            count = "6";
                        }
                        else if (!allAssemblyNames.Contains("AS7_"))
                        {
                            count = "7";
                        }
                        else if (!allAssemblyNames.Contains("AS8_"))
                        {
                            count = "8";
                        }
                        else if (!allAssemblyNames.Contains("AS9_"))
                        {
                            count = "9";
                        }
                        else if (!allAssemblyNames.Contains("AS10_"))
                        {
                            count = "10";
                        }
                        else if (!allAssemblyNames.Contains("AS11_"))
                        {
                            count = "11";
                        }
                        else if (!allAssemblyNames.Contains("AS12_"))
                        {
                            count = "12";
                        }
                        else if (!allAssemblyNames.Contains("AS13_"))
                        {
                            count = "13";
                        }
                        else if (!allAssemblyNames.Contains("AS14_"))
                        {
                            count = "14";
                        }
                        else if (!allAssemblyNames.Contains("AS15_"))
                        {
                            count = "15";
                        }
                        else if (!allAssemblyNames.Contains("AS16_"))
                        {
                            count = "16";
                        }
                        else if (!allAssemblyNames.Contains("AS17_"))
                        {
                            count = "17";
                        }
                        else if (!allAssemblyNames.Contains("AS18_"))
                        {
                            count = "18";
                        }
                        else if (!allAssemblyNames.Contains("AS19_"))
                        {
                            count = "19";
                        }
                        else if (!allAssemblyNames.Contains("AS20_"))
                        {
                            count = "20";
                        }
                        else if (!allAssemblyNames.Contains("AS21_"))
                        {
                            count = "21";
                        }
                        else if (!allAssemblyNames.Contains("AS22_"))
                        {
                            count = "22";
                        }
                        else if (!allAssemblyNames.Contains("AS23_"))
                        {
                            count = "23";
                        }
                        else if (!allAssemblyNames.Contains("AS24_"))
                        {
                            count = "24";
                        }
                        else if (!allAssemblyNames.Contains("AS25_"))
                        {
                            count = "25";
                        }
                        else if (!allAssemblyNames.Contains("AS26_"))
                        {
                            count = "26";
                        }
                        else if (!allAssemblyNames.Contains("AS27_"))
                        {
                            count = "27";
                        }
                        else if (!allAssemblyNames.Contains("AS28_"))
                        {
                            count = "28";
                        }
                        else if (!allAssemblyNames.Contains("AS29_"))
                        {
                            count = "29";
                        }
                        else if (!allAssemblyNames.Contains("AS30_"))
                        {
                            count = "30";
                        }
                        else if (!allAssemblyNames.Contains("AS31_"))
                        {
                            count = "31";
                        }
                        else if (!allAssemblyNames.Contains("AS32_"))
                        {
                            count = "32";
                        }
                        else if (!allAssemblyNames.Contains("AS33_"))
                        {
                            count = "33";
                        }
                        else if (!allAssemblyNames.Contains("AS34_"))
                        {
                            count = "34";
                        }
                        else if (!allAssemblyNames.Contains("AS35_"))
                        {
                            count = "35";
                        }
                        else if (!allAssemblyNames.Contains("AS36_"))
                        {
                            count = "36";
                        }
                        else if (!allAssemblyNames.Contains("AS37_"))
                        {
                            count = "37";
                        }
                        else if (!allAssemblyNames.Contains("AS38_"))
                        {
                            count = "38";
                        }
                        else if (!allAssemblyNames.Contains("AS39_"))
                        {
                            count = "39";
                        }
                        else if (!allAssemblyNames.Contains("AS40_"))
                        {
                            count = "40";
                        }
                        else if (!allAssemblyNames.Contains("AS41_"))
                        {
                            count = "41";
                        }
                        else if (!allAssemblyNames.Contains("AS42_"))
                        {
                            count = "42";
                        }
                        else if (!allAssemblyNames.Contains("AS43_"))
                        {
                            count = "43";
                        }
                        else if (!allAssemblyNames.Contains("AS44_"))
                        {
                            count = "44";
                        }
                        else if (!allAssemblyNames.Contains("AS45_"))
                        {
                            count = "45";
                        }
                        else if (!allAssemblyNames.Contains("AS46_"))
                        {
                            count = "46";
                        }
                        else if (!allAssemblyNames.Contains("AS47_"))
                        {
                            count = "47";
                        }
                        else if (!allAssemblyNames.Contains("AS48_"))
                        {
                            count = "48";
                        }
                        else if (!allAssemblyNames.Contains("AS49_"))
                        {
                            count = "49";
                        }

                        Element el = doc.GetElement(id);
                        string AssemblyName = el.Name;
                        AssemblyInstance AssembliesSelected = el as AssemblyInstance;
                        int index = AssemblyName.IndexOf("_");
                        string AssemblyNameEdited = AssemblyName.Remove(0, index+1);
                        string numberFull = "AS" + count;
                        string numberFullwUnderscore = numberFull + "_";

                        if (AssembliesSelected.AllowsAssemblyViewCreation())
                        {

                            FilteredElementCollector collector = new FilteredElementCollector(doc).
                                OfCategory(BuiltInCategory.OST_TitleBlocks).OfClass(typeof(FamilySymbol));
                            //change .ElementAt(4) to pull a string of "Harveys 8_5 x 11 Assembly Titleblock"
                            ElementId titleblockId = collector.FirstOrDefault(q => q.Name == "Harveys 8_5 x 11 Assembly Titleblock").Id;

                            ViewSheet viewSheet = AssemblyViewUtils.CreateSheet(doc, AssembliesSelected.Id, titleblockId);
                            viewSheet.ViewName = AssemblyNameEdited;
                            viewSheet.SheetNumber = numberFull;

                            //viewSheet.SheetNumber = "This is the Sheet Number, make a ++ command that has a handler to not repeat numbers... Example AS + (some code)";
                            //This will be in the other code

                            //Creating and Applying 3D Detail
                            View3D view3d = AssemblyViewUtils.Create3DOrthographic(doc, AssembliesSelected.Id);
                            view3d.Scale = 12;
                            Viewport.Create(doc, viewSheet.Id, view3d.Id, new XYZ(0.7, 0.35, 0));


                            View view3dTemplate = (from v in new FilteredElementCollector(doc)
                             .OfClass(typeof(View)).Cast<View>()
                                                   where v.IsTemplate == true && v.Name ==
                                                   "Harvey's Plumbing 3D View"
                                                   select v).First();

                            view3d.ViewTemplateId = view3dTemplate.Id;

                            //Creating and Applying Elevation Top
                            ViewSection detailSectionH = AssemblyViewUtils.CreateDetailSection(doc, AssembliesSelected.Id, AssemblyDetailViewOrientation.ElevationTop);
                            detailSectionH.Scale = 16;
                            Viewport.Create(doc, viewSheet.Id, detailSectionH.Id, new XYZ(0.1, 0.25, 0));

                            View viewPlanTemplate = (from v in new FilteredElementCollector(doc)
                             .OfClass(typeof(View)).Cast<View>()
                                                     where v.IsTemplate == true && v.Name ==
                                                     "Harvey's Plumbing Detail View"
                                                     select v).First();

                            detailSectionH.ViewTemplateId = viewPlanTemplate.Id;

                            //Create schedule, add and edit fields
                            ViewSchedule pipeFittingSchedule = AssemblyViewUtils.CreateSingleCategorySchedule(doc, id, new ElementId(BuiltInCategory.OST_PipeFitting));
                            ScheduleSheetInstance.Create(doc, viewSheet.Id, pipeFittingSchedule.Id, new XYZ(0.03, 0.675, 0));

                            View scheduleTemplate = (from v in new FilteredElementCollector(doc)
                             .OfClass(typeof(View)).Cast<View>()
                                                     where v.IsTemplate == true && v.Name ==
                                                     "Schedule"
                                                     select v).First();

                            pipeFittingSchedule.ViewTemplateId = scheduleTemplate.Id;

                            IList<ScheduleFieldId> listOfFieldIds = pipeFittingSchedule.Definition.GetFieldOrder();
                            foreach (SchedulableField schedulableField in pipeFittingSchedule.Definition.GetSchedulableFields())
                            {
                                //Judge if the FieldType is ScheduleFieldType.Instance.
                                if (schedulableField.FieldType == ScheduleFieldType.Instance)
                                {
                                    //Get ParameterId of SchedulableField.
                                    ElementId parameterId = schedulableField.ParameterId;

                                    //Add a new schedule field to the view schedule by using the SchedulableField as argument of AddField method of Autodesk.Revit.DB.ScheduleDefinition class.
                                    string fieldName = schedulableField.GetName(doc);
                                    string fieldNameRemove1 = schedulableField.GetName(doc);

                                    if (fieldName == "Assembly Name") //Lots of ORs
                                    {
                                        ScheduleField fieldAssemblyName = pipeFittingSchedule.Definition.AddField(schedulableField);
                                        fieldAssemblyName.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                                        fieldAssemblyName.IsHidden = true;

                                        //Group and sort the view schedule by type
                                        if (fieldAssemblyName.ParameterId == new ElementId(BuiltInParameter.ASSEMBLY_NAME))
                                        {

                                            ScheduleSortGroupField sortGroupField2 = new ScheduleSortGroupField(fieldAssemblyName.FieldId);
                                            sortGroupField2.ShowHeader = true;
                                            pipeFittingSchedule.Definition.AddSortGroupField(sortGroupField2);
                                        }
                                    }

                                    else if (fieldName == "Comments")
                                    {
                                        ScheduleField fieldComments = pipeFittingSchedule.Definition.AddField(schedulableField);
                                        fieldComments.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                                        ScheduleFilter filterScheduleAssembly = new ScheduleFilter(fieldComments.FieldId, ScheduleFilterType.NotContains, "Omit from Schedule");
                                        pipeFittingSchedule.Definition.AddFilter(filterScheduleAssembly);
                                        fieldComments.IsHidden = true;
                                    }

                                    else if (fieldName == "Harvey's Part Number")
                                    {
                                        ScheduleField fieldPieceNumber = pipeFittingSchedule.Definition.InsertField(schedulableField, 0);
                                        fieldPieceNumber.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                                        ScheduleSortGroupField sortGroupField2 = new ScheduleSortGroupField(fieldPieceNumber.FieldId);
                                        sortGroupField2.ShowHeader = false;
                                        pipeFittingSchedule.Definition.AddSortGroupField(sortGroupField2);
                                        fieldPieceNumber.ColumnHeading = "Part #";

                                        fieldPieceNumber.GridColumnWidth = 0.75 * fieldPieceNumber.GridColumnWidth;
                                    }

                                    else if (fieldName == "Basic Part Name")
                                    {
                                        ScheduleField fieldBasicPartName = pipeFittingSchedule.Definition.AddField(schedulableField);
                                        fieldBasicPartName.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                                        fieldBasicPartName.GridColumnWidth = 4 * fieldBasicPartName.GridColumnWidth;
                                        fieldBasicPartName.ColumnHeading = "Part Name";
                                    }

                                    foreach (ScheduleFieldId idRemove1 in listOfFieldIds)
                                    {
                                        if (fieldNameRemove1 == "Family and Type" || fieldNameRemove1 == "Count")
                                        {
                                            pipeFittingSchedule.Definition.RemoveField(idRemove1);
                                        }
                                    }
                                }
                            }

                            //Create schedule, add and edit fields
                            ViewSchedule pipeSchedule = AssemblyViewUtils.CreateSingleCategorySchedule(doc, id, new ElementId(BuiltInCategory.OST_PipeCurves));
                            ScheduleSheetInstance.Create(doc, viewSheet.Id, pipeSchedule.Id, new XYZ(0.03, 0.35, 0));
                            ScheduleDefinition sd = pipeSchedule.Definition;

                            pipeSchedule.ViewTemplateId = scheduleTemplate.Id;

                            foreach (SchedulableField schedulableField1 in pipeSchedule.Definition.GetSchedulableFields())
                            {
                                //Judge if the FieldType is ScheduleFieldType.Instance.
                                if (schedulableField1.FieldType == ScheduleFieldType.Instance)
                                {
                                    //Get ParameterId of SchedulableField.
                                    ElementId parameterId = schedulableField1.ParameterId;

                                    //Add a new schedule field to the view schedule by using the SchedulableField as argument of AddField method of Autodesk.Revit.DB.ScheduleDefinition class.
                                    string fieldName = schedulableField1.GetName(doc);
                                    string fieldNameRemove1 = schedulableField1.GetName(doc);

                                    if (fieldName == "Assembly Name")
                                    {
                                        ScheduleField fieldAssemblyName = pipeSchedule.Definition.AddField(schedulableField1);
                                        fieldAssemblyName.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                                        fieldAssemblyName.IsHidden = true;

                                        //Group and sort the view schedule by type
                                        if (fieldAssemblyName.ParameterId == new ElementId(BuiltInParameter.ASSEMBLY_NAME))
                                        {

                                            ScheduleSortGroupField sortGroupField2 = new ScheduleSortGroupField(fieldAssemblyName.FieldId);
                                            sortGroupField2.ShowHeader = true;
                                            pipeSchedule.Definition.AddSortGroupField(sortGroupField2);
                                        }
                                    }

                                    else if (fieldName == "Comments")
                                    {
                                        ScheduleField fieldComments = pipeSchedule.Definition.AddField(schedulableField1);
                                        fieldComments.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                                        ScheduleFilter filterScheduleAssembly = new ScheduleFilter(fieldComments.FieldId, ScheduleFilterType.NotContains, "Omit from Schedule");
                                        sd.AddFilter(filterScheduleAssembly);
                                        fieldComments.IsHidden = true;
                                    }

                                    else if (fieldName == "Harvey's Part Number")
                                    {
                                        ScheduleField fieldPieceNumber = pipeSchedule.Definition.AddField(schedulableField1);
                                        fieldPieceNumber.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                                        ScheduleSortGroupField sortGroupField2 = new ScheduleSortGroupField(fieldPieceNumber.FieldId);
                                        sortGroupField2.ShowHeader = false;
                                        pipeSchedule.Definition.AddSortGroupField(sortGroupField2);
                                        fieldPieceNumber.ColumnHeading = "Part #";

                                        fieldPieceNumber.GridColumnWidth = fieldPieceNumber.GridColumnWidth;
                                    }
                                    foreach (ScheduleFieldId idRemove1 in listOfFieldIds)
                                    {
                                        if (fieldNameRemove1 == "Family and Type" || fieldNameRemove1 == "Count")
                                        {
                                            pipeSchedule.Definition.RemoveField(idRemove1);
                                        }
                                    }

                                }
                                if (schedulableField1.FieldType == ScheduleFieldType.ElementType)
                                {
                                    ElementId parameterId = schedulableField1.ParameterId;

                                    //Add a new schedule field to the view schedule by using the SchedulableField 
                                    //as argument of AddField method of Autodesk.Revit.DB.ScheduleDefinition class.
                                    string fieldName = schedulableField1.GetName(doc);
                                    string fieldNameRemove1 = schedulableField1.GetName(doc);

                                    if (fieldName == "Basic Part Name") //Lots of ORs
                                    {
                                        ScheduleField fieldBasicPartName = pipeSchedule.Definition.AddField(schedulableField1);
                                        fieldBasicPartName.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                                        fieldBasicPartName.GridColumnWidth = 2.5 * fieldBasicPartName.GridColumnWidth;
                                        fieldBasicPartName.ColumnHeading = "Part Name";
                                    }

                                }


                            }
                            ScheduleField fieldLength = sd.AddField(ScheduleFieldType.Instance, new ElementId(BuiltInParameter.CURVE_ELEM_LENGTH));
                            fieldLength.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                            fieldLength.GridColumnWidth = 1.25 * fieldLength.GridColumnWidth;
                        }


                        AssembliesSelected.AssemblyTypeName = numberFullwUnderscore + AssemblyNameEdited;

                    }

                    tx.Commit();


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
            private bool ShouldSkip(ElementId parameterId)
            {
                foreach (BuiltInParameter bip in s_skipParameters)
                {
                    if (new ElementId(bip) == parameterId)
                        return true;
                }
                return false;
            }
        private static BuiltInParameter[] s_skipParameters = new BuiltInParameter[] { BuiltInParameter.ALL_MODEL_MARK };
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
