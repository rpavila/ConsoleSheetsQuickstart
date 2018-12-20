using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Reflection;

namespace ConsoleSheetsQuickstart
{
    class Program
    {        
        static void Main(string[] args)
        {
            Assembly currentAssembly = Assembly.GetExecutingAssembly();
            string absolutePath = Path.GetDirectoryName(currentAssembly.Location);
            Assembly testAssembly = Assembly.LoadFile(absolutePath + "/GSpreadSheet.dll");
            Type googleSheetsType = testAssembly.GetType("GSpreadSheet.GoogleSheets");

            object[] constructorParams = new object[] { "credentials.json" };
            object gsInstance = Activator.CreateInstance(googleSheetsType, constructorParams);
            
            //  Spreadsheets DocID
            String spreadsheetId = "1mfFHzoYsz9Rfypme3bRZGHYqXLVFJNLVz4hNiSS9tfk";
            
            //  Creating a Session
            MethodInfo methodOpenSession = googleSheetsType.GetMethod("OpenSession");
            object[] parametersArray = new object[] { spreadsheetId };
            object session = methodOpenSession.Invoke(gsInstance, parametersArray);

            //  Creating a CellAddress instance
            Type cellAddressType = testAssembly.GetType("GSpreadSheet.CellAddress");
            Type cellAddressValueType = testAssembly.GetType("GSpreadSheet.CellAddressWithValue");
            List<object> listCellAddress = new List<object>();

            object instanceCellAddress = Activator.CreateInstance(cellAddressType);
            FieldInfo propertyAddress = cellAddressType.GetField("Address");
            propertyAddress.SetValue(instanceCellAddress, "A2");
            listCellAddress.Add(instanceCellAddress);

            //  Calling a ReadCellValues function
            //MethodInfo methodReadCellValues = googleSheetsType.GetMethod("ReadCellValues");
            //parametersArray = new object[] { session, listCellAddress };
            //IList<object> values = (IList<object>)methodReadCellValues.Invoke(gsInstance, parametersArray);
            //int i = 1;
            //foreach (var row in values)
            //{
            //    propertyAddress = cellAddressValueType.GetField("Address");
            //    var address = propertyAddress.GetValue(row);
            //    propertyAddress = cellAddressValueType.GetField("SheetName");
            //    var sheetName = propertyAddress.GetValue(row);
            //    propertyAddress = cellAddressValueType.GetField("Value");
            //    var val = propertyAddress.GetValue(row);
            //    Console.WriteLine("{0}) - {1}, {2}, {3}", i++, sheetName, address, val);
            //}

            //  Creating a CellAddressWithValue instance
            listCellAddress = new List<object>();
            cellAddressValueType = testAssembly.GetType("GSpreadSheet.CellAddressWithValue");
            object instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType);
            propertyAddress = cellAddressValueType.GetField("Address");
            propertyAddress.SetValue(instanceCellAddressValue, "H5");
            propertyAddress = cellAddressValueType.GetField("Value");
            propertyAddress.SetValue(instanceCellAddressValue, "Andrew");
            listCellAddress.Add(instanceCellAddressValue);
            instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType);
            propertyAddress = cellAddressValueType.GetField("Address");
            propertyAddress.SetValue(instanceCellAddressValue, "H7");
            propertyAddress = cellAddressValueType.GetField("Value");
            propertyAddress.SetValue(instanceCellAddressValue, "Constantine");
            listCellAddress.Add(instanceCellAddressValue);
            instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType);
            propertyAddress = cellAddressValueType.GetField("Address");
            propertyAddress.SetValue(instanceCellAddressValue, "H9");
            propertyAddress = cellAddressValueType.GetField("Value");
            propertyAddress.SetValue(instanceCellAddressValue, "Ricardo");
            listCellAddress.Add(instanceCellAddressValue);

            //  Calling a WriteCellValues function
            //MethodInfo methodWriteCellValues = googleSheetsType.GetMethod("WriteCellValues");
            //parametersArray = new object[] { session, listCellAddress };
            //methodWriteCellValues.Invoke(gsInstance, parametersArray);

            Console.ReadLine();
        }
    }
}
