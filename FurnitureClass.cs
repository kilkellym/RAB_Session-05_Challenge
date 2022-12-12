using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RAB_Session_05_Challenge
{
    internal class FurnitureSet
    {
        public string Set { get; set; }
        public string Room { get; set; }
        public string[] Furniture { get; set; }

        public FurnitureSet(string set, string room, string furnitureList)
        {
            Set = set;
            Room = room;
            Furniture = furnitureList.Split(',');
        }

    }

    internal class FurnitureType
    {
        public string Name { get; set; }
        public string Family { get; set; }
        public string Type { get; set; }
        public FamilySymbol RevitType {get; set;} 

        public FurnitureType(string name, string family, string type)
        {
            Name = name;
            Family = family;
            Type = type;

        }
        private FamilySymbol GetFamilySymbolByName(Document doc, string familyName, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));

            foreach (FamilySymbol currentFS in collector)
            {
                if (currentFS.Name == typeName && currentFS.FamilyName == familyName)
                    return currentFS;
            }

            return null;
        }
    }
}
