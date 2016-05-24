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

    public class QRCodeGenerator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects
            

            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                List<ElementId> ids = new List<ElementId>();
                Selection sel = uiApp.ActiveUIDocument.Selection;

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
                        refs = sel.PickObjects(ObjectType.Element,
                        "Please Select Element(s) to Generate QrCode.");
                    }
                    catch (Autodesk.Revit.Exceptions
                    .OperationCanceledException)
                    {
                        //                      return Result.Cancelled;
                    }
                    ids = new List<ElementId>(
                    refs.Select<Reference, ElementId>(
                    r => r.ElementId));
                }

                foreach (ElementId id in ids)
                {
                    Element el = doc.GetElement(id);

                    string data = el.Name;
                    string folderpath = @"C:\Users\Brian\Desktop\";
                    string Filename = el.Name;

                    GenerateQRCode(data, folderpath, Filename);
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
        private static void GenerateQRCode(string Data, string folderpath, string fileName)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://api.qrserver.com/v1/create-qr-code/?size=900x900&data=" + Data);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                if ((response.StatusCode == HttpStatusCode.OK ||
                    response.StatusCode == HttpStatusCode.Moved ||
                    response.StatusCode == HttpStatusCode.Redirect) &&
                    response.ContentType.StartsWith("image", StringComparison.OrdinalIgnoreCase))
                {
                    using (Stream inputStream = response.GetResponseStream())
                    using (Stream outputStream = File.OpenWrite(folderpath + fileName + ".png"))
                    {
                        byte[] buffer = new byte[4096];
                        int bytesRead;
                        do
                        {
                            bytesRead = inputStream.Read(buffer, 0, buffer.Length);
                            outputStream.Write(buffer, 0, bytesRead);
                        } while (bytesRead != 0);
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }

        }
    }
}
