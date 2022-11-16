#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Forms = System.Windows.Forms;

#endregion

namespace RAB_Session_02_Skills
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // setup open file dialog
            Forms.OpenFileDialog selectFile = new Forms.OpenFileDialog();
            selectFile.InitialDirectory = "C:\\";
            selectFile.Filter = "CSV files|*.csv|All files|*.*";
            //selectFile.Filter = "Revit file|*.rvt;*.rfa";
            //selectFile.Filter = "All files|*.*";
            selectFile.Multiselect = false;

            // open file dialog
            string fileName = "";
            if (selectFile.ShowDialog() == Forms.DialogResult.OK)
            {
                fileName = selectFile.FileName;
            }

            if(fileName != "")
            {
                // do something with the file
            }

            myStruct struct1 = new myStruct();
            struct1.Name = "test name";
            struct1.Description = "this is a description";
            struct1.Distance = 100;

            myStruct struct2 = new myStruct("test name2", "description 2", 200);

            List<myStruct> myList = new List<myStruct>();
            myList.Add(struct1);
            myList.Add(struct2);

            foreach(myStruct currentStruct in myList)
            {
                Debug.Print(currentStruct.Name);
            }

            FilteredElementCollector vftCollector = new FilteredElementCollector(doc);
            vftCollector.OfClass(typeof(ViewFamilyType));

            Element tblock = GetTitleBlockByName(doc, "E1 30x42 Horizontal");
            ElementId tBlockId = tblock.Id;

            ViewFamilyType planVFT = null;
            ViewFamilyType rcpVFT = null;
            foreach (ViewFamilyType vft in vftCollector)
            {
                if (vft.ViewFamily == ViewFamily.FloorPlan)
                    planVFT = vft;

                if (vft.ViewFamily == ViewFamily.CeilingPlan)
                    rcpVFT = vft;
            }

            Transaction t = new Transaction(doc);
            t.Start("Create some stuff");

            Level newLevel = Level.Create(doc, 20);

            ViewPlan newPlanView = ViewPlan.Create(doc, planVFT.Id, newLevel.Id);
            ViewPlan newRCPView = ViewPlan.Create(doc, rcpVFT.Id, newLevel.Id);

            ViewSheet newSheet = ViewSheet.Create(doc, tBlockId);
            ViewSheet newCeilingSheet = ViewSheet.Create(doc, tBlockId);

            XYZ insertPoint = new XYZ(2, 1, 0);
            XYZ secondInsertPoint = new XYZ(0, 1, 0);

            Viewport newViewPort = Viewport.Create(doc, newSheet.Id, newPlanView.Id, insertPoint);
            Viewport newCeilingViewPort = Viewport.Create(doc, newSheet.Id, newRCPView.Id, secondInsertPoint);

            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }

        internal double ConvertMetersToFeet(double meters)
        {
            double feet = meters * 3.28084;

            return feet;
        }

        internal Element GetTitleBlockByName(Document doc, string name)
        {
            FilteredElementCollector tblockCollector = new FilteredElementCollector(doc);
            tblockCollector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            foreach(Element currentTblock in tblockCollector)
            {
                if (currentTblock.Name == name)
                    return currentTblock;
            }

            return null;
        }

        struct myStruct
        {
            public string Name;
            public string Description;
            public double Distance;
            private int Number;

            public myStruct(string name, string description, double dist)
            {
                Name = name;
                Description = description;
                Distance = dist;
                Number = 0;
            }
        }
    }
}
