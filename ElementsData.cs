﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace BillofQuantities
{
    internal static class InputData
    {
        public static string folderPath { get; set; }
        public static bool instancesSheet { get; set; }
        public static bool elementTypesSheet { get; set; }
        public static bool billofQuantitiesSheet { get; set; }
    }

    public class EI
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

        public EI(Element eI)
        {
            ID = eI.Id.IntegerValue;
            IsType = 0;
            CategoryName = eI.Category.Name;
            TypeName = eI.Name;
            TypeNameId = eI.GetTypeId().IntegerValue;
            FamilyName = eI as FamilyInstance != null ?
                (eI as FamilyInstance).Symbol.Family.Name : "*NA*";
        }
    }

    public class ET
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
        public string Cost { get; set; }
        public string Unit { get; set; }
        public string AssemblyCode { get; set; }
        public string AssemblyDesc { get; set; }
        public string KeyValue { get; set; }
        public string KeyText { get; set; }

        public ET(Element eT)
        {
            ID = eT.Id != null ? eT.Id.IntegerValue : -1;
            IsType = 1;
            CategoryName = eT.Category != null ? eT.Category.Name : "*NA*";
            CategoryId = eT.Category!= null ? eT.Category.Id.IntegerValue : -1;
            TypeName = eT.Name != null ? eT.Name : "*NA*";
            FamilyName = eT as FamilySymbol != null ?
                (eT as FamilySymbol).FamilyName : "*NA*";
        }
    }

    public class BQ
    {
        public string AssemblyCode { get; set; }
        public string AssemblyDesc { get; set; }
        public string KeyValue { get; set; } // ListaBQ sorted by KeyValue
        public string KeyText { get; set; }
        public string Unit { get; set; }
        public string Quant { get; set; }
        public string PrUnit { get; set; }
        public string Partial { get; set; }
    }
}