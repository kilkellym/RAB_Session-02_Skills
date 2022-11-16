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
    public class CmdProjectSetup : IExternalCommand
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

            // declare variables
            string levelPath = "C:\\Users\\micha\\OneDrive\\Documents\\Revit Add-in Bootcamp\\RAB_Session_02_Challenge_Levels.csv";
            string sheetPath = "C:\\Users\\micha\\OneDrive\\Documents\\Revit Add-in Bootcamp\\RAB_Session_02_Challenge_Sheets.csv";

            List<string[]> levelData = new List<string[]>();
            List<string[]> sheetData = new List<string[]>();

            // read text file data
            string[] levelArray = System.IO.File.ReadAllLines(levelPath);
            string[] sheetArray = System.IO.File.ReadAllLines(sheetPath);

            // loop through file data and put into list
            foreach(string levelString in levelArray)
            {
                string[] cellString = levelString.Split(',');
                string[] levelDataArray = new string[3];

                levelDataArray[0] = cellString[0];
                levelDataArray[1] = cellString[1]; 
                levelDataArray[2] = cellString[2];

                levelData.Add(levelDataArray);
            }

            foreach (string sheetString in sheetArray)
            {
                string[] cellString = sheetString.Split(',');
                string[] sheetDataArray = new string[2];

                sheetDataArray[0] = cellString[0];
                sheetDataArray[1] = cellString[1];

                sheetData.Add(sheetDataArray);
            }

            // remove header rows
            levelData.RemoveAt(0);
            sheetData.RemoveAt(0);

            // create levels
            Transaction t1 = new Transaction(doc);
            t1.Start("Create levels");

            foreach (string[] curLevelData in levelData)
            {
                double heightFeet = 0;
                double heightMeters = 0;

                bool convertFeet = double.TryParse(curLevelData[1], out heightFeet);
                bool convertMeters = double.TryParse(curLevelData[2], out heightMeters);

                // show option to use metric w/ conversion to decimal feet
                ConvertMetersToFeet(heightMeters);

                Level myLevel = Level.Create(doc, heightFeet);
                myLevel.Name = curLevelData[0];
            }

            t1.Commit();
            t1.Dispose();

            // create sheets
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            ElementId tblockId = collector.FirstElementId();

            Transaction t2 = new Transaction(doc);
            t2.Start("Create sheets");

            foreach (string[] curSheetData in sheetData)
            {
                ViewSheet mySheet = ViewSheet.Create(doc, tblockId);
                mySheet.Name = curSheetData[1];
                mySheet.SheetNumber = curSheetData[0];
            }

            t2.Commit();
            t2.Dispose();


            return Result.Succeeded;
        }

        internal double ConvertMetersToFeet(double meters)
        {
            double feet = meters * 3.28084;

            double test = UnitUtils.ConvertToInternalUnits(meters, UnitTypeId.Meters);
            return feet;
        }
    }
}
