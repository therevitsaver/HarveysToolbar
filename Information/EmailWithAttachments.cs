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

namespace Email
{
    [Transaction(TransactionMode.Manual)]

    public class EmailWithAttachments : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects

            try
            {
                //Opens Email
                OutlookApp outlookApp = new OutlookApp();
                MailItem mailItem = outlookApp.CreateItem(OlItemType.olMailItem);
                
                //Assigns contacts and all other important information to the email.
                mailItem.To = "ed@harveysph.com";
                mailItem.CC = "bob@harveysph.com";
                mailItem.Subject = "Pipe and Pipe Assembly Order Form";
                mailItem.HTMLBody = "Ed, here are the pipe order forms as requested per Bob Harvey.";

                
                string fileName1 = @"C:\Users\Brian\Desktop\Schedules\Plumbing - Pipe Order Sheet.csv";
                string fileName2 = @"C:\Users\Brian\Desktop\Schedules\Plumbing - Pipe Fitting Order Sheet.csv";
                mailItem.Attachments.Add(fileName1, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                mailItem.Attachments.Add(fileName2, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
               
                


                //Set a high priority to the message
                mailItem.Importance = OlImportance.olImportanceHigh;

                mailItem.Display(false);



                // Measure their total length

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

            catch (System.Exception ex)
            {
                // If something went wrong return Result.Failed
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
