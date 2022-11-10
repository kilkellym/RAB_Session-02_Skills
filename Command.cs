#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

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

            string[] myStringArray = new string[4];
            List<string> myStringList = new List<string>();

            myStringArray[0] = "This is the first item";
            myStringArray[1] = "12312";
            myStringArray[2] = "This is the third item";
            myStringArray[3] = "This is the last item";
            //myStringArray[4] = "adfasdfasdfasf";

            myStringList.Add("this is the first item");
            myStringList.Add("this is the second item");

            myStringList.RemoveAt(0);

            List<string[]> myArrayList = new List<string[]>();
            myArrayList.Add(myStringArray);
            
            foreach(string myString in myStringArray)
            {
                //TaskDialog.Show("Test", myString);
            }

            foreach(string myString2 in myStringList)
            {
                //TaskDialog.Show("Test", myString2);
            }

            foreach (string[] myArray in myArrayList)
            {
                string comboString = "";

                foreach(string myString3 in myArray)
                {
                    comboString = comboString + myString3;
                }

                Debug.Print("******" + comboString);
            }

            // read text file data
            string filepath = @"C:\Users\micha\OneDrive\Documents\Revit Add-in Bootcamp\Session 02_Room List.csv";
            string fileText = System.IO.File.ReadAllText(filepath);

            TaskDialog.Show("Test", fileText);

            string[] fileArray = System.IO.File.ReadAllLines(filepath);

            foreach(string rowString in fileArray)
            {
                string[] cellString = rowString.Split(',');

                string roomNumber = cellString[0];
                string roomName = cellString[1];
                string roomArea = cellString[2];

                //double roomAreaDouble = double.Parse(roomArea);

                double roomAreaDouble2 = 0;
                bool didItParse = double.TryParse(roomArea, out roomAreaDouble2);
            }

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            ElementId tblockId = collector.FirstElementId();

            Transaction t = new Transaction(doc);
            t.Start("Create level and sheet");

            // create level
            double levelHeight = ConvertMetersToFeet(10);
            Level myLevel = Level.Create(doc, levelHeight);
            myLevel.Name = "My Level";

            // create sheet
            ViewSheet mySheet = ViewSheet.Create(doc, tblockId);
            mySheet.Name = "My new sheet";
            mySheet.SheetNumber = "A101";

            t.Commit();
            t.Dispose();


            return Result.Succeeded;
        }

        internal double ConvertMetersToFeet(double meters)
        {
            double feet = meters * 3.28084;

            return feet;
        }
    }
}
