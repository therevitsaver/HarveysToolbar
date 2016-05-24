using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using winform = System.Windows.Forms;


using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using adWin = Autodesk.Windows;




namespace HarveysToolbar
{
    public class App : IExternalApplication
    {
        //define a method that will create a tab and button
        static void AddRibbonPanel(UIControlledApplication application)
        {
            ////////////// █████╗ ███████╗███████╗███████╗███╗   ███╗██████╗ ██╗  ██╗   ██╗    ████████╗ ██████╗  ██████╗ ██╗     ███████╗
            //////////////██╔══██╗██╔════╝██╔════╝██╔════╝████╗ ████║██╔══██╗██║  ╚██╗ ██╔╝    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██╔════╝
            //////////////███████║███████╗███████╗█████╗  ██╔████╔██║██████╔╝██║   ╚████╔╝        ██║   ██║   ██║██║   ██║██║     ███████╗
            //////////////██╔══██║╚════██║╚════██║██╔══╝  ██║╚██╔╝██║██╔══██╗██║    ╚██╔╝         ██║   ██║   ██║██║   ██║██║     ╚════██║
            //////////////██║  ██║███████║███████║███████╗██║ ╚═╝ ██║██████╔╝███████╗██║          ██║   ╚██████╔╝╚██████╔╝███████╗███████║
            //////////////╚═╝  ╚═╝╚══════╝╚══════╝╚══════╝╚═╝     ╚═╝╚═════╝ ╚══════╝╚═╝          ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚══════╝

            //create a custom ribbon tab
            String tabName = "Harvey's Spool o' Matic 9000";
            application.CreateRibbonTab(tabName);
            
            // Add a new ribbon panel

            RibbonPanel ribbonPanel1AT = application.CreateRibbonPanel(tabName, "Assembly Tools");


             //Get dll assembly path
             string thisAssemblyPath1AT = Assembly.GetExecutingAssembly().Location;

            //create push button for RenumberingTool
            PushButtonData b1DataAT = new PushButtonData(
                "cmdUngroupAssemble",
                "Ungroup" + System.Environment.NewLine + " Assemble ",
                thisAssemblyPath1AT, "UngroupAssemble.UngroupAssemble");

            PushButton pb1AT = ribbonPanel1AT.AddItem(b1DataAT) as PushButton;
            pb1AT.ToolTip = "Ungroup and Assemble all existing groups in the project!";
            BitmapImage pb1ATImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/ungroupassemble.png"));
            pb1AT.LargeImage = pb1ATImage;

            //////////////██████╗ ███████╗███╗   ██╗██╗   ██╗███╗   ███╗██████╗ ███████╗██████╗ ██╗███╗   ██╗ ██████╗     ████████╗ ██████╗  ██████╗ ██╗     ███████╗
            //////////////██╔══██╗██╔════╝████╗  ██║██║   ██║████╗ ████║██╔══██╗██╔════╝██╔══██╗██║████╗  ██║██╔════╝     ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██╔════╝
            //////////////██████╔╝█████╗  ██╔██╗ ██║██║   ██║██╔████╔██║██████╔╝█████╗  ██████╔╝██║██╔██╗ ██║██║  ███╗       ██║   ██║   ██║██║   ██║██║     ███████╗
            //////////////██╔══██╗██╔══╝  ██║╚██╗██║██║   ██║██║╚██╔╝██║██╔══██╗██╔══╝  ██╔══██╗██║██║╚██╗██║██║   ██║       ██║   ██║   ██║██║   ██║██║     ╚════██║
            //////////////██║  ██║███████╗██║ ╚████║╚██████╔╝██║ ╚═╝ ██║██████╔╝███████╗██║  ██║██║██║ ╚████║╚██████╔╝       ██║   ╚██████╔╝╚██████╔╝███████╗███████║
            //////////////╚═╝  ╚═╝╚══════╝╚═╝  ╚═══╝ ╚═════╝ ╚═╝     ╚═╝╚═════╝ ╚══════╝╚═╝  ╚═╝╚═╝╚═╝  ╚═══╝ ╚═════╝        ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚══════╝

           

            // Add a new ribbon panel
            RibbonPanel ribbonPanel1RT = application.CreateRibbonPanel(tabName, "Renumbering Tools");

            //Get dll assembly path
            string thisAssemblyPath1RT = Assembly.GetExecutingAssembly().Location;

            //create push button for RenumberingTool
            PushButtonData b1Data = new PushButtonData(
                "cmdPipeAndFittingRenumberingTool",
                "Renumber" + System.Environment.NewLine + " MEP Tool ",
                thisAssemblyPath1RT, "RenumberingTools.PipeAndFittingRenumberingTool");

            PushButton pb1 = ribbonPanel1RT.AddItem(b1Data) as PushButton;
            pb1.ToolTip = "Renumber Pipe and Pipe Fittings";
            BitmapImage pb1Image = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/renumberingtoolpipeandfittings.png"));
            pb1.LargeImage = pb1Image;

            //Get dll assembly path
            string thisAssemblyPath4RT = Assembly.GetExecutingAssembly().Location;

            //create push button for RenumberingTool
            PushButtonData b4Data = new PushButtonData(
                "cmdRenumberTagLeft",
                "Renumber" + System.Environment.NewLine + " Tag Left ",
                thisAssemblyPath4RT, "RenumberingTools.RenumberTagLeft");

            PushButton pb4 = ribbonPanel1RT.AddItem(b4Data) as PushButton;
            pb4.ToolTip = "Renumber Pipe and Pipe Fittings";
            BitmapImage pb4Image = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/renumbertagleft.png"));
            pb4.LargeImage = pb4Image;

            //Get dll assembly path
            string thisAssemblyPath5RT = Assembly.GetExecutingAssembly().Location;

            //create push button for RenumberingTool
            PushButtonData b5Data = new PushButtonData(
                "cmdRenumberTagRight",
                "Renumber" + System.Environment.NewLine + " Tag Right ",
                thisAssemblyPath5RT, "RenumberingTools.RenumberTagRight");

            PushButton pb5 = ribbonPanel1RT.AddItem(b5Data) as PushButton;
            pb5.ToolTip = "Renumbers all assemblies appropriately";
            BitmapImage pb5Image = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/renumbertagright.png"));
            pb5.LargeImage = pb5Image;

            //Get dll assembly path
            string thisAssemblyPath2RT = Assembly.GetExecutingAssembly().Location;

            //create push button for RenumberAssembliesAndBuildingSheets
            PushButtonData b2Data = new PushButtonData(
                "cmdRenumberAssembliesOne",
                "Create Initial" + System.Environment.NewLine + " Assemblies ",
                thisAssemblyPath2RT, "RenumberingTools.RenumberAssembliesOne");

            PushButton pb2 = ribbonPanel1RT.AddItem(b2Data) as PushButton;
            pb2.ToolTip = "Renumbers all assemblies appropriately";
            BitmapImage pb2Image = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/initialassemblies.png"));
            pb2.LargeImage = pb2Image;

            //Get dll assembly path
            string thisAssemblyPath3RT = Assembly.GetExecutingAssembly().Location;

            //create push button for RenumberAssembliesAndBuildingSheetsFromEndpoint
            PushButtonData b3Data = new PushButtonData(
                "cmdRenumberAssembliesTwo",
                "Create Post" + System.Environment.NewLine + " Assemblies ",
                thisAssemblyPath3RT, "RenumberingTools.RenumberAssembliesTwo");

            PushButton pb3 = ribbonPanel1RT.AddItem(b3Data) as PushButton;
            pb3.ToolTip = "Renumber Assemblies from an End Point";
            BitmapImage pb3Image = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/postassemblies.png"));
            pb3.LargeImage = pb3Image;

            //Get dll assembly path
            string thisAssemblyPath6RT = Assembly.GetExecutingAssembly().Location;

            //create push button for RenumberAssemblies
            PushButtonData b6Data = new PushButtonData(
                "cmdRenumberingAssembliesTool",
                "Renumber" + System.Environment.NewLine + " Assemblies ",
                thisAssemblyPath6RT, "RenumberingTools.RenumberingAssembliesTool");

            PushButton pb6 = ribbonPanel1RT.AddItem(b6Data) as PushButton;
            pb6.ToolTip = "Renumber Assemblies from an End Point";
            BitmapImage pb6Image = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/renumberassembliestool.png"));
            pb6.LargeImage = pb6Image;

            //////////////███████╗██╗  ██╗███████╗███████╗████████╗    ████████╗ ██████╗  ██████╗ ██╗     ███████╗
            //////////////██╔════╝██║  ██║██╔════╝██╔════╝╚══██╔══╝    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██╔════╝
            //////////////███████╗███████║█████╗  █████╗     ██║          ██║   ██║   ██║██║   ██║██║     ███████╗
            //////////////╚════██║██╔══██║██╔══╝  ██╔══╝     ██║          ██║   ██║   ██║██║   ██║██║     ╚════██║
            //////////////███████║██║  ██║███████╗███████╗   ██║          ██║   ╚██████╔╝╚██████╔╝███████╗███████║
            //////////////╚══════╝╚═╝  ╚═╝╚══════╝╚══════╝   ╚═╝          ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚══════╝

            // Add a new ribbon panel
            RibbonPanel ribbonPanel1ST = application.CreateRibbonPanel(tabName, "Sheet Tools");

            //Get dll assembly path
            string thisAssemblyPath1ST = Assembly.GetExecutingAssembly().Location;

            //create push button for QR Code Generator
            PushButtonData b1DataST = new PushButtonData(
                "cmdQRCodeGenerator",
                "QR Code" + System.Environment.NewLine + " Generator ",
                thisAssemblyPath1ST, "SheetTools.QRCodeGenerator");

            PushButton pb1ST = ribbonPanel1ST.AddItem(b1DataST) as PushButton;
            pb1ST.ToolTip = "Generates QR Codes for all assemblies created in the project";
            BitmapImage pb1STImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/qrcode.png"));
            pb1ST.LargeImage = pb1STImage;

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //This is the pull down button for all of Harvey's View Template selections.

            string viewTemplatesPD = Assembly.GetExecutingAssembly().Location;

            // create push buttons for split button drop down

            PushButtonData bOne = new PushButtonData("ButtonNameA", "Plumbing Detail View",
            viewTemplatesPD, "SheetTools.ViewTemplatePlumbingDetailView");
            bOne.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/viewtemplates.png"));

            PushButtonData bTwo = new PushButtonData("ButtonNameB", "3D View Template",
            viewTemplatesPD, "SheetTools.ViewTemplate3D");
            bTwo.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/viewtemplates.png"));

            PushButtonData bThree = new PushButtonData("ButtonNameC", "Plumbing Plan View Template",
            viewTemplatesPD, "SheetTools.ViewTemplatePlumbingPlan");
            bThree.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/viewtemplates.png"));

            PushButtonData bFour = new PushButtonData("ButtonNameD", "Pipe and Fittings View Template",
            viewTemplatesPD, "SheetTools.ViewTemplatePipeAndFittings");
            bFour.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/viewtemplates.png"));

            

            PulldownButtonData sb1 = new PulldownButtonData("pullDownButtonA", "Harvey's" + System.Environment.NewLine + "View Templates");
            sb1.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/viewtemplates.png"));
            PulldownButton pd = ribbonPanel1ST.AddItem(sb1) as PulldownButton;
            pd.AddPushButton(bOne);
            pd.AddPushButton(bTwo);
            pd.AddPushButton(bThree);
            pd.AddPushButton(bFour);

            //Get dll assembly path
            string thisAssemblyPath4ST = Assembly.GetExecutingAssembly().Location;

            //create push button for removing View Templates
            PushButtonData b4DataST = new PushButtonData(
                "cmdViewTemplateNone",
                "Remove" + System.Environment.NewLine + "View Template",
                thisAssemblyPath4ST, "SheetTools.ViewTemplateNone");

            PushButton pb4ST = ribbonPanel1ST.AddItem(b4DataST) as PushButton;
            pb4ST.ToolTip = "Remove the view template from current view.";
            BitmapImage pb4STImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/viewtemplatenone.png"));
            pb4ST.LargeImage = pb4STImage;

            //Get dll assembly path
            string sheettools = Assembly.GetExecutingAssembly().Location;

            // create push buttons for split button drop down
            PushButtonData stb1 = new PushButtonData("ButtonNameA", "Add Section Box",
            sheettools, "SheetTools.ViewSectionBoxOn");
            stb1.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/viewtemplatenone.png"));

            PushButtonData stb2 = new PushButtonData("ButtonNameB", "Remove Secton Box",
            sheettools, "SheetTools.ViewSectionBoxOff");
            stb2.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/viewtemplatenone.png"));

            //Creates a pull down button set for the Section Box Tools
            PulldownButtonData st1 = new PulldownButtonData("pullDownButtonB", "Harvey's" + System.Environment.NewLine + "Section Box Tools");
            st1.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/viewtemplatenone.png"));
            PulldownButton st = ribbonPanel1ST.AddItem(st1) as PulldownButton;
            st.AddPushButton(stb1);
            st.AddPushButton(stb2);

            //Get dll assembly path
            string thisAssemblyPath2ST = Assembly.GetExecutingAssembly().Location;

            //create push button for Bill of Materials for Pipe 
            PushButtonData b2DataST = new PushButtonData(
                "cmdSheetToolTwo",
                "Pipe" + System.Environment.NewLine + " Order Form ",
                thisAssemblyPath2ST, "SheetTools.PipeOrderForm");

            PushButton pb2ST = ribbonPanel1ST.AddItem(b2DataST) as PushButton;
            pb2ST.ToolTip = "Prints a bill of materials, for pipe, to excel";
            BitmapImage pb2STImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/pipeorderform.png"));
            pb2ST.LargeImage = pb2STImage;

            //Get dll assembly path
            string thisAssemblyPath3ST = Assembly.GetExecutingAssembly().Location;

            //create push button for Bill of Materials for Pipe Fittings
            PushButtonData b3DataST = new PushButtonData(
                "cmdSheetToolThree",
                "Pipe Fitting" + System.Environment.NewLine + " Order Form ",
                thisAssemblyPath3ST, "SheetTools.PipeFittingOrderForm");

            PushButton pb3ST = ribbonPanel1ST.AddItem(b3DataST) as PushButton;
            pb3ST.ToolTip = "Prints a bill of materials, for pipe fittings, to excel";
            BitmapImage pb3STImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/pipeorderform.png"));
            pb3ST.LargeImage = pb3STImage;
            

            ////////////// ██████╗ █████╗ ██╗      ██████╗██╗   ██╗██╗      █████╗ ████████╗ ██████╗ ██████╗ 
            //////////////██╔════╝██╔══██╗██║     ██╔════╝██║   ██║██║     ██╔══██╗╚══██╔══╝██╔═══██╗██╔══██╗
            //////////////██║     ███████║██║     ██║     ██║   ██║██║     ███████║   ██║   ██║   ██║██████╔╝
            //////////////██║     ██╔══██║██║     ██║     ██║   ██║██║     ██╔══██║   ██║   ██║   ██║██╔══██╗
            //////////////╚██████╗██║  ██║███████╗╚██████╗╚██████╔╝███████╗██║  ██║   ██║   ╚██████╔╝██║  ██║
            ////////////// ╚═════╝╚═╝  ╚═╝╚══════╝ ╚═════╝ ╚═════╝ ╚══════╝╚═╝  ╚═╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝

            // Add a new ribbon panel
            RibbonPanel ribbonPanel1C = application.CreateRibbonPanel(tabName, "Calculator");

            //Get dll assembly path
            string thisAssemblyPath1C = Assembly.GetExecutingAssembly().Location;

            //create push button for RenumberingTool
            PushButtonData b1DataC = new PushButtonData(
                "cmdManualPipeCalculator",
                "Manual" + System.Environment.NewLine + " Pipe Calculator ",
                thisAssemblyPath1C, "Calculator.ManualPipeCalculator");

            PushButton pb1C = ribbonPanel1C.AddItem(b1DataC) as PushButton;
            pb1C.ToolTip = "Manual Pipe Calculator for calculating total pipe length based on clicks";
            BitmapImage pb1CImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/pipecalculator.png"));
            pb1C.LargeImage = pb1CImage;


            //////////////██╗███╗   ██╗███████╗ ██████╗ ██████╗ ███╗   ███╗ █████╗ ████████╗██╗ ██████╗ ███╗   ██╗
            //////////////██║████╗  ██║██╔════╝██╔═══██╗██╔══██╗████╗ ████║██╔══██╗╚══██╔══╝██║██╔═══██╗████╗  ██║
            //////////////██║██╔██╗ ██║█████╗  ██║   ██║██████╔╝██╔████╔██║███████║   ██║   ██║██║   ██║██╔██╗ ██║
            //////////////██║██║╚██╗██║██╔══╝  ██║   ██║██╔══██╗██║╚██╔╝██║██╔══██║   ██║   ██║██║   ██║██║╚██╗██║
            //////////////██║██║ ╚████║██║     ╚██████╔╝██║  ██║██║ ╚═╝ ██║██║  ██║   ██║   ██║╚██████╔╝██║ ╚████║
            //////////////╚═╝╚═╝  ╚═══╝╚═╝      ╚═════╝ ╚═╝  ╚═╝╚═╝     ╚═╝╚═╝  ╚═╝   ╚═╝   ╚═╝ ╚═════╝ ╚═╝  ╚═══╝


            // Add a new ribbon panel
            RibbonPanel ribbonPanel1Email = application.CreateRibbonPanel(tabName, "Information");

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //This is the pull down button for all of Harvey's View Template selections.

            

            ////Get dll assembly path
            //string thisAssemblyPath2Email= Assembly.GetExecutingAssembly().Location;

           
            ////create push button for A360
            //PushButtonData b2DataEmail = new PushButtonData(
            //    "cmdA360",
            //    "A360",
            //    thisAssemblyPath2Email, "Email.A360");

            //PushButton pb2Email = ribbonPanel1Email.AddItem(b2DataEmail) as PushButton;
            //pb2Email.ToolTip = "Open A360!";
            //BitmapImage pb2EmailImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/a360.png"));
            //pb2Email.LargeImage = pb2EmailImage;

            ////Get dll assembly path
            //string thisAssemblyPath6Email = Assembly.GetExecutingAssembly().Location;

            ////create push button for buildingops
            //PushButtonData b6DataEmail = new PushButtonData(
            //    "cmdBuildingOps",
            //    "Autodesk" + System.Environment.NewLine + " Building Ops ",
            //    thisAssemblyPath6Email, "Email.BuildingOps");

            //PushButton pb6Email = ribbonPanel1Email.AddItem(b6DataEmail) as PushButton;
            //pb6Email.ToolTip = "Open Harvey's Building Ops!";
            //BitmapImage pb6EmailImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/buildingops.png"));
            //pb6Email.LargeImage = pb6EmailImage;

            ////Get dll assembly path
            //string thisAssemblyPath3Email = Assembly.GetExecutingAssembly().Location;

            ////create push button for google calendar
            //PushButtonData b3DataEmail = new PushButtonData(
            //    "cmdGoogleCalendar",
            //    "Google" + System.Environment.NewLine + " Calendar ",
            //    thisAssemblyPath3Email, "Email.GoogleCalendar");

            //PushButton pb3Email = ribbonPanel1Email.AddItem(b3DataEmail) as PushButton;
            //pb3Email.ToolTip = "Open Google Calendar!";
            //BitmapImage pb3EmailImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/google calendar.png"));
            //pb3Email.LargeImage = pb3EmailImage;

            ////Get dll assembly path
            //string thisAssemblyPath7Email = Assembly.GetExecutingAssembly().Location;

            ////create push button for Gmail
            //PushButtonData b7DataEmail = new PushButtonData(
            //    "cmdGmail",
            //    "Gmail",
            //    thisAssemblyPath7Email, "Email.Gmail");

            //PushButton pb7Email = ribbonPanel1Email.AddItem(b7DataEmail) as PushButton;
            //pb7Email.ToolTip = "Open your Gmail Account!";
            //BitmapImage pb7EmailImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/gmail.png"));
            //pb7Email.LargeImage = pb7EmailImage;
            

            ////Get dll assembly path
            //string thisAssemblyPath4Email = Assembly.GetExecutingAssembly().Location;

            ////create push button for slack
            //PushButtonData b4DataEmail = new PushButtonData(
            //    "cmdSlack",
            //    "Slack",
            //    thisAssemblyPath3Email, "Email.Slack");

            //PushButton pb4Email = ribbonPanel1Email.AddItem(b4DataEmail) as PushButton;
            //pb4Email.ToolTip = "Open Slack!";
            //BitmapImage pb4EmailImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/Slack.png"));
            //pb4Email.LargeImage = pb4EmailImage;

            ////Get dll assembly path
            //string thisAssemblyPath5Email = Assembly.GetExecutingAssembly().Location;

            ////create push button for googlemymaps
            //PushButtonData b5DataEmail = new PushButtonData(
            //    "cmdGoogleMyMaps",
            //    "Harvey's" + System.Environment.NewLine + " Job Map ",
            //    thisAssemblyPath3Email, "Email.GoogleMyMaps");

            //PushButton pb5Email = ribbonPanel1Email.AddItem(b5DataEmail) as PushButton;
            //pb5Email.ToolTip = "Open Harvey's Job Map!";
            //BitmapImage pb5EmailImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/Googlemymaps.png"));
            //pb5Email.LargeImage = pb5EmailImage;


            //Get dll assembly path
            string thisAssemblyPath1Email = Assembly.GetExecutingAssembly().Location;

            //create push button for email ordering forms
            PushButtonData b1DataEmail = new PushButtonData(
                "cmdEmail",
                "Email" + System.Environment.NewLine + " Order Forms ",
                thisAssemblyPath1Email, "Email.PipeOrderForm");

            PushButton pb1Email = ribbonPanel1Email.AddItem(b1DataEmail) as PushButton;
            pb1Email.ToolTip = "Email the Pipe Order forms to Ed!";
            BitmapImage pb1EmailImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/email.png"));
            pb1Email.LargeImage = pb1EmailImage;

            //create push button for email ordering forms
            PushButtonData b2DataEmail = new PushButtonData(
                "cmdEmailWithAttachments",
                "Email" + System.Environment.NewLine + " With Attachments ",
                thisAssemblyPath1Email, "Email.EmailWithAttachments");

            PushButton pb2Email = ribbonPanel1Email.AddItem(b2DataEmail) as PushButton;
            pb2Email.ToolTip = "Email the Pipe Order forms to Ed (with attachements)!";
            BitmapImage pb2EmailImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/email.png"));
            pb2Email.LargeImage = pb1EmailImage;
            
            //create push button for calculator
            PushButtonData b3DataEmail = new PushButtonData(
                "cmdCalculator",
                "Calculator",
                thisAssemblyPath1Email, "Email.Calculator");

            PushButton pb3Email = ribbonPanel1Email.AddItem(b3DataEmail) as PushButton;
            pb3Email.ToolTip = "Opens Windows Calculator!";
            BitmapImage pb3EmailImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/calculator.png"));
            pb3Email.LargeImage = pb3EmailImage;

            string emails = Assembly.GetExecutingAssembly().Location;

            // create push buttons for split button drop down
            PushButtonData emb1 = new PushButtonData("ButtonNameA", "A360",
            emails, "Email.A360");
            emb1.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/A360.png"));

            PushButtonData emb2 = new PushButtonData("ButtonNameB", "Slack",
            emails, "Email.Slack");
            emb2.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/slack.png"));

            PushButtonData emb3 = new PushButtonData("ButtonNameC", "Google Calendar",
            emails, "Email.GoogleCalendar");
            emb3.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/google calendar.png"));

            PushButtonData emb4 = new PushButtonData("ButtonNameD", "Gmail",
            emails, "Email.Gmail");
            emb4.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/gmail.png"));

            PushButtonData emb5 = new PushButtonData("ButtonNameE", "Autodesk Building Ops",
            emails, "Email.BuildingOps");
            emb5.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/buildingops.png"));

            PushButtonData emb6 = new PushButtonData("ButtonNameF", "Harvey's Job Map",
            emails, "Email.GoogleMyMaps");
            emb6.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/Googlemymaps.png"));

            PulldownButtonData em1 = new PulldownButtonData("pullDownButtonB", "Harvey's" + System.Environment.NewLine + "Resources");
            em1.LargeImage = new BitmapImage(new Uri("pack://application:,,,/HarveysToolbar;component/Resources/database.png"));
            PulldownButton em = ribbonPanel1Email.AddItem(em1) as PulldownButton;
            em.AddPushButton(emb1);
            em.AddPushButton(emb2);
            em.AddPushButton(emb3);
            em.AddPushButton(emb4);
            em.AddPushButton(emb5);
            em.AddPushButton(emb6);

            // Defines Color To Bottom of Ribbon

            adWin.RibbonControl ribbon = adWin.ComponentManager.Ribbon;

            LinearGradientBrush gradientBrush
              = new LinearGradientBrush();

            gradientBrush.GradientStops.Add(
              new GradientStop(Colors.IndianRed, 1));

            foreach (adWin.RibbonTab tab in ribbon.Tabs)
            {
                if (tab.Id.Contains("Harvey's"))
                {
                    foreach (adWin.RibbonPanel panel in tab.Panels)
                    {
                        panel.CustomPanelTitleBarBackground
                          = gradientBrush;

                    }
                }
            }

        }

        public Result OnShutdown(UIControlledApplication application)
        {
            //do nothing 
            return Result.Succeeded;
        }
        public Result OnStartup(UIControlledApplication application)
        {
            //call our method that will load up our toolbar
            AddRibbonPanel(application);
            return Result.Succeeded;
        }
    }
}