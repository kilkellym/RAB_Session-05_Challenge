#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Forms = System.Windows.Forms;
using Excel = OfficeOpenXml;
using OfficeOpenXml;

#endregion

namespace RAB_Session_05_Challenge
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

            // NOTE: Remember to add the EPPlus library using the NuGet Package Manager

            // part 0: prompt user to select Excel file
            Forms.OpenFileDialog selectFile = new Forms.OpenFileDialog();
            selectFile.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";
            selectFile.InitialDirectory = "C:\\";
            selectFile.Multiselect = false;

            string excelFile = "";

            if (selectFile.ShowDialog() == Forms.DialogResult.OK)
                excelFile = selectFile.FileName;

            if (excelFile == "")
            {
                TaskDialog.Show("Error", "Please select an Excel file.");
                return Result.Failed;
            }

            // set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // open Excel file
            ExcelPackage excel = new ExcelPackage(excelFile);
            ExcelWorkbook workbook = excel.Workbook;
            ExcelWorksheet setWS = workbook.Worksheets[0];
            ExcelWorksheet typeWS = workbook.Worksheets[1];

            // get row and column count
            int setRows = setWS.Dimension.Rows;
            int typeRows = typeWS.Dimension.Rows;

            // part 1: get data into list of classes
            List<FurnitureType> furnitureTypeList = new List<FurnitureType>();
            List<FurnitureSet> furnitureSetList = new List<FurnitureSet>();

            for (int i = 1; i <= setRows; i++)
            {
                string setName = setWS.Cells[i, 1].Value.ToString();
                string setRoom = setWS.Cells[i, 2].Value.ToString();
                string setFurn = setWS.Cells[i, 3].Value.ToString();

                FurnitureSet curSet = new FurnitureSet(setName, setRoom, setFurn);
                furnitureSetList.Add(curSet);
            }

            for (int j = 1; j <= typeRows; j++)
            {
                string typeName = typeWS.Cells[j, 1].Value.ToString();
                string typeFamily = typeWS.Cells[j, 2].Value.ToString();
                string typeType = typeWS.Cells[j, 3].Value.ToString();

                FurnitureType curType = new FurnitureType(typeName, typeFamily, typeType);
                furnitureTypeList.Add(curType);
            }

            furnitureTypeList.RemoveAt(0);
            furnitureSetList.RemoveAt(0);

            int overallCounter = 0;

            // part 2. - get rooms, loop through them and insert correct furniture
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Insert families");

                foreach (SpatialElement room in collector)
                {
                    int counter = 0;

                    string furnSet = GetParameterValueByName(room, "Furniture Set");
                    Debug.Print(furnSet);

                    LocationPoint roomLocation = room.Location as LocationPoint;
                    XYZ insPoint = roomLocation.Point;

                    foreach(FurnitureSet curSet in furnitureSetList)
                    {
                        if(curSet.Set == furnSet)
                        {
                            foreach(string furnPiece in curSet.Furniture)
                            {
                                foreach(FurnitureType curType in furnitureTypeList)
                                {
                                    if(curType.Name == furnPiece.Trim())
                                    {
                                        FamilySymbol curFS = GetFamilySymbolByName(doc, curType.Family, curType.Type);
                                        
                                        if(curFS.IsActive == false)
                                            curFS.Activate();

                                        FamilyInstance instance = doc.Create.NewFamilyInstance(insPoint, curFS, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                        counter++;
                                        overallCounter++;
                                    }
                                }
                            }
                        }
                    }

                    SetParameterByName(room, "Furniture Count", counter);
                }

                t.Commit();
            }

            TaskDialog.Show("complete", "Inserted " + overallCounter.ToString() + " pieces of furniture.");
                
            return Result.Succeeded;
        }

        private void SetParameterByName(Element element, string paramName, int value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
            {
                Parameter param = paramList[0];

                param.Set(value);
            }
        }

        private FamilySymbol GetFamilySymbolByName(Document doc, string familyName, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));

            foreach(FamilySymbol currentFS in collector)
            {
                if (currentFS.Name == typeName && currentFS.FamilyName == familyName)
                    return currentFS;
            }

            return null;
        }

        private string GetParameterValueByName(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if(paramList != null)
            {
                Parameter param = paramList[0];
                string paramValue = param.AsString();
                return paramValue;
            }

            return "";
            
        }

        private List<string[]> GetFurnitureTypes()
        {
            List<string[]> returnList = new List<string[]>();
            returnList.Add(new string[] { "Furniture Name", "Revit Family Name", "Revit Family Type" });
            returnList.Add(new string[] { "desk", "Desk", "60in x 30in" });
            returnList.Add(new string[] { "task chair", "Chair-Task", "Chair-Task" });
            returnList.Add(new string[] { "side chair", "Chair-Breuer", "Chair-Breuer" });
            returnList.Add(new string[] { "bookcase", "Shelving", "96in x 12in x 84in" });
            returnList.Add(new string[] { "loveseat", "Sofa", "54in" });
            returnList.Add(new string[] { "teacher desk", "Table-Rectangular", "48in x 30in" });
            returnList.Add(new string[] { "student desk", "Desk", "60in x 30in Student" });
            returnList.Add(new string[] { "computer desk", "Table-Rectangular", "48in x 30in" });
            returnList.Add(new string[] { "lab desk", "Table-Rectangular", "72in x 30in" });
            returnList.Add(new string[] { "lounge chair", "Chair-Corbu", "Chair-Corbu" });
            returnList.Add(new string[] { "coffee table", "Table-Coffee", "30in x 60in x 18in" });
            returnList.Add(new string[] { "sofa", "Sofa-Corbu", "Sofa-Corbu" });
            returnList.Add(new string[] { "dining table", "Table-Dining", "30in x 84in x 22in" });
            returnList.Add(new string[] { "dining chair", "Chair-Breuer", "Chair-Breuer" });
            returnList.Add(new string[] { "stool", "Chair-Task", "Chair-Task" });

            return returnList;
        }

        private List<string[]> GetFurnitureSets()
        {
            List<string[]> returnList = new List<string[]>();
            returnList.Add(new string[] { "Furniture Set", "Room Type", "Included Furniture" });
            returnList.Add(new string[] { "A", "Office", "desk, task chair, side chair, bookcase" });
            returnList.Add(new string[] { "A2", "Office", "desk, task chair, side chair, bookcase, loveseat" });
            returnList.Add(new string[] { "B", "Classroom - Large", "teacher desk, task chair, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk" });
            returnList.Add(new string[] { "B2", "Classroom - Medium", "teacher desk, task chair, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk" });
            returnList.Add(new string[] { "C", "Computer Lab", "computer desk, computer desk, computer desk, computer desk, computer desk, computer desk, task chair, task chair, task chair, task chair, task chair, task chair" });
            returnList.Add(new string[] { "D", "Lab", "teacher desk, task chair, lab desk, lab desk, lab desk, lab desk, lab desk, lab desk, lab desk, stool, stool, stool, stool, stool, stool, stool" });
            returnList.Add(new string[] { "E", "Student Lounge", "lounge chair, lounge chair, lounge chair, sofa, coffee table, bookcase" });
            returnList.Add(new string[] { "F", "Teacher's Lounge", "lounge chair, lounge chair, sofa, coffee table, dining table, dining chair, dining chair, dining chair, dining chair, bookcase" });
            returnList.Add(new string[] { "G", "Waiting Room", "lounge chair, lounge chair, sofa, coffee table" });

            return returnList;
        }
    }
}
