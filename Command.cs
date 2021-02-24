#region Namespaces

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

#endregion Namespaces

namespace BillofQuantities
{
    #region Data holding Classes
    internal static class InputData
    {
        public static string folderPath { get; set; }
        public static bool instancesSheet { get; set; }
        public static bool elementTypesSheet { get; set; }
        public static bool billofQuantitiesSheet { get; set; }
    }
    public class ListEI
    {
        public int ID { get; set; }
        public int IsType { get; set; }
        public string CategoryName { get; set; }
        public string TypeName { get; set; }
        public int TypeNameId { get; set; } // ListaEI sorted by TypeNameId
        public string FamilyName { get; set; }
        public string Volume { get; set; }
        public string Area { get; set; }
        public string Width { get; set; }
        public string Length { get; set; }
    }

    public class ListET
    {
        public int ID { get; set; } // ListaET sorted by ID
        public int IsType { get; set; }
        public string CategoryName { get; set; }
        public int CategoryId { get; set; }
        public string TypeName { get; set; }
        public string FamilyName { get; set; }
        public int Quantity { get; set; }
        public string TotalVolume { get; set; }
        public string TotalArea { get; set; }
        public string TotalLength { get; set; }
        public string Costunit { get; set; }
        public string AssemblyCode { get; set; }
        public string AssemblyDesc { get; set; }
        public string KeyValue { get; set; }
        public string KeyText { get; set; }
    }

    public class ListMQ
    {
        public string AssemblyCode { get; set; }
        public string AssemblyDesc { get; set; }
        public string KeyValue { get; set; } // ListaMQ sorted by KeyValue
        public string KeyText { get; set; }
        public string Unit { get; set; }
        public string Quant { get; set; }
        public string PrUnit { get; set; }
        public string Partial { get; set; }
    }

    #endregion Data holding Classes

    [Transaction(TransactionMode.ReadOnly)]
    public class Command : IExternalCommand
    {
        public string IncomingValue { get; set; }

        Result IExternalCommand.Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Calls Modeless Form
            try
            {
                Application.thisApp.ShowForm(commandData.Application);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}