using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;

namespace UngroupAssemble
{
    [Transaction(TransactionMode.Manual)]

    public class UngroupAssemble : IExternalCommand
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

                //-----------------------------     
                ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Assemblies);

                // Apply the filter to the elements in the active document
                // Use shortcut WhereElementIsNotElementType() to find view instances only
                FilteredElementCollector assemblyCollector = new FilteredElementCollector(doc);
                IList<Element> assemblies =
                assemblyCollector.WherePasses(filter).WhereElementIsNotElementType().ToElements();
                String allAssemblyNames = "The assemblies in the current document are:\n";
                String assemblyNumber = "";

                //Obtain all existing assembly names and list them.
                foreach (Element e in assemblies)
                {
                    allAssemblyNames += e.Name + "\n";
                }
                //-------------------------------------------------------------------------------------------------------------------------------            
                //Test to find which assembly number has not been chosen, and use that assembly number if it hasn't been chosen.
                if (!allAssemblyNames.Contains("AS0A_"))
                {
                    assemblyNumber = "AS0A";
                }
                else if (!allAssemblyNames.Contains("AS0B_"))
                {
                    assemblyNumber = "AS0B";
                }
                else if (!allAssemblyNames.Contains("AS0C_"))
                {
                    assemblyNumber = "AS0C";
                }
                else if (!allAssemblyNames.Contains("AS0D_"))
                {
                    assemblyNumber = "AS0D";
                }
                else if (!allAssemblyNames.Contains("AS0E_"))
                {
                    assemblyNumber = "AS0E";
                }
                else if (!allAssemblyNames.Contains("AS0F_"))
                {
                    assemblyNumber = "AS0F";
                }
                else if (!allAssemblyNames.Contains("AS0G_"))
                {
                    assemblyNumber = "AS0G";
                }
                else if (!allAssemblyNames.Contains("AS0H_"))
                {
                    assemblyNumber = "AS0H";
                }
                else if (!allAssemblyNames.Contains("AS0I_"))
                {
                    assemblyNumber = "AS0I";
                }
                else if (!allAssemblyNames.Contains("AS0J_"))
                {
                    assemblyNumber = "AS0J";
                }
                else if (!allAssemblyNames.Contains("AS0K_"))
                {
                    assemblyNumber = "AS0K";
                }
                else if (!allAssemblyNames.Contains("AS0L_"))
                {
                    assemblyNumber = "AS0L";
                }
                else if (!allAssemblyNames.Contains("AS0M_"))
                {
                    assemblyNumber = "AS0M";
                }
                else if (!allAssemblyNames.Contains("AS0N_"))
                {
                    assemblyNumber = "AS0N";
                }
                else if (!allAssemblyNames.Contains("AS0O_"))
                {
                    assemblyNumber = "AS0O";
                }
                else if (!allAssemblyNames.Contains("AS0P_"))
                {
                    assemblyNumber = "AS0P";
                }
                else if (!allAssemblyNames.Contains("AS0Q_"))
                {
                    assemblyNumber = "AS0Q";
                }
                else if (!allAssemblyNames.Contains("AS0R_"))
                {
                    assemblyNumber = "AS0R";
                }
                else if (!allAssemblyNames.Contains("AS0S_"))
                {
                    assemblyNumber = "AS0S";
                }
                else if (!allAssemblyNames.Contains("AS0T_"))
                {
                    assemblyNumber = "AS0T";
                }
                else if (!allAssemblyNames.Contains("AS0U_"))
                {
                    assemblyNumber = "AS0U";
                }
                else if (!allAssemblyNames.Contains("AS0V_"))
                {
                    assemblyNumber = "AS0V";
                }
                else if (!allAssemblyNames.Contains("AS0W_"))
                {
                    assemblyNumber = "AS0W";
                }
                else if (!allAssemblyNames.Contains("AS0X_"))
                {
                    assemblyNumber = "AS0X";
                }
                else if (!allAssemblyNames.Contains("AS0Y_"))
                {
                    assemblyNumber = "AS0Y";
                }
                else if (!allAssemblyNames.Contains("AS0Z_"))
                {
                    assemblyNumber = "AS0Z";
                }
                else if (!allAssemblyNames.Contains("AS1A_"))
                {
                    assemblyNumber = "AS1A";
                }
                else if (!allAssemblyNames.Contains("AS1B_"))
                {
                    assemblyNumber = "AS1B";
                }
                else if (!allAssemblyNames.Contains("AS1C_"))
                {
                    assemblyNumber = "AS1C";
                }
                else if (!allAssemblyNames.Contains("AS1D_"))
                {
                    assemblyNumber = "AS1D";
                }
                else if (!allAssemblyNames.Contains("AS1E_"))
                {
                    assemblyNumber = "AS1E";
                }
                else if (!allAssemblyNames.Contains("AS1F_"))
                {
                    assemblyNumber = "AS1F";
                }
                else if (!allAssemblyNames.Contains("AS1G_"))
                {
                    assemblyNumber = "AS1G";
                }
                else if (!allAssemblyNames.Contains("AS1H_"))
                {
                    assemblyNumber = "AS1H";
                }
                else if (!allAssemblyNames.Contains("AS1I_"))
                {
                    assemblyNumber = "AS1I";
                }
                else if (!allAssemblyNames.Contains("AS1J_"))
                {
                    assemblyNumber = "AS1J";
                }
                else if (!allAssemblyNames.Contains("AS1K_"))
                {
                    assemblyNumber = "AS1K";
                }
                else if (!allAssemblyNames.Contains("AS1L_"))
                {
                    assemblyNumber = "AS1L";
                }
                else if (!allAssemblyNames.Contains("AS1M_"))
                {
                    assemblyNumber = "AS1M";
                }
                else if (!allAssemblyNames.Contains("AS1N_"))
                {
                    assemblyNumber = "AS1N";
                }
                else if (!allAssemblyNames.Contains("AS1O_"))
                {
                    assemblyNumber = "AS1O";
                }
                else if (!allAssemblyNames.Contains("AS1P_"))
                {
                    assemblyNumber = "AS1P";
                }
                else if (!allAssemblyNames.Contains("AS1Q_"))
                {
                    assemblyNumber = "AS1Q";
                }
                else if (!allAssemblyNames.Contains("AS1R_"))
                {
                    assemblyNumber = "AS1R";
                }
                else if (!allAssemblyNames.Contains("AS1S_"))
                {
                    assemblyNumber = "AS1S";
                }
                else if (!allAssemblyNames.Contains("AS1T_"))
                {
                    assemblyNumber = "AS1T";
                }
                else if (!allAssemblyNames.Contains("AS1U_"))
                {
                    assemblyNumber = "AS1U";
                }
                else if (!allAssemblyNames.Contains("AS1V_"))
                {
                    assemblyNumber = "AS1V";
                }
                else if (!allAssemblyNames.Contains("AS1W_"))
                {
                    assemblyNumber = "AS1W";
                }
                else if (!allAssemblyNames.Contains("AS1X_"))
                {
                    assemblyNumber = "AS1X";
                }
                else if (!allAssemblyNames.Contains("AS1Y_"))
                {
                    assemblyNumber = "AS1Y";
                }
                else if (!allAssemblyNames.Contains("AS1Z_"))
                {
                    assemblyNumber = "AS1Z";
                }
                //----------------------------

                List<ElementId> ids = new List<ElementId>();
                List<Element> ele = new List<Element>();
                List<ElementId> idsFiltered = new List<ElementId>();

                ElementId categoryId = null;

                Group GroupSelected = doc.GetElement(
                      uiDoc.Selection.PickObject(ObjectType.Element,
                        "Select a Group")) as Group;

                string AssemblyName = GroupSelected.Name;

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Ungrouping");

                    foreach (ElementId gids in GroupSelected.UngroupMembers())
                    {

                        Element gel = doc.GetElement(gids);
                        if (gel.GetValidTypes().FirstOrDefault() != null)
                        {
                            ids.Add(gids);
                            if (gel.Category.Name == "Pipes" || gel.Category.Name == "Ducts")
                            {
                                categoryId = gel.Category.Id;
                            }
                            ele.Add(gel);
                        }
                    }

                    tx.Commit();


                    tx.Start("Creating Assembly");
                    //----------------------------------------------------------------------------------------------
                    //Cleans out unassembl-able ids.
                    foreach (Element pipeOrFitting in ele)
                    {
                        ElementId idFiltered = pipeOrFitting.Id;

                        if (pipeOrFitting.Category.Name == "Pipes")
                        {
                            idsFiltered.Add(idFiltered);
                        }
                        else if (pipeOrFitting.Category.Name == "Pipe Accessories")
                        {
                            idsFiltered.Add(idFiltered);
                        }
                        else if (pipeOrFitting.Category.Name == "Pipe Fittings")
                        {
                            idsFiltered.Add(idFiltered);
                        }
                        else if (pipeOrFitting.Category.Name == "Plumbing Fixtures")
                        {
                            idsFiltered.Add(idFiltered);
                        }
                        else if (pipeOrFitting.Category.Name == "Ducts")
                        {
                            idsFiltered.Add(idFiltered);
                        }
                        else if (pipeOrFitting.Category.Name == "Duct Accessories")
                        {
                            idsFiltered.Add(idFiltered);
                        }
                        else if (pipeOrFitting.Category.Name == "Duct Fittings")
                        {
                            idsFiltered.Add(idFiltered);
                        }
                        else if (pipeOrFitting.Category.Name == "Mechanical Equipment")
                        {
                            idsFiltered.Add(idFiltered);
                        }
                        else if (pipeOrFitting.Category.Name == "Fabrication Parts")
                        {
                            idsFiltered.Add(idFiltered);
                        }
                    }

                    //----------------------------------------------------------------------------------------------

                    AssemblyInstance assemblyInstance = AssemblyInstance.Create(doc, idsFiltered, categoryId);

                    tx.Commit();

                    tx.Start("Renaming Assembly / Creating Views");
                    assemblyInstance.AssemblyTypeName = assemblyNumber + "_" + AssemblyName;

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
    }
}
