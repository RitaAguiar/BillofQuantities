#region namespaces

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Diagnostics;

#endregion //namespaces

namespace BillofQuantities
{
    public class RevitUtils
    {
        // Stopwatches - measures time spent on tasks
        Stopwatch sw = Stopwatch.StartNew();
        //Stopwatch sw_EI_ET = Stopwatch.StartNew();
        Stopwatch sw_EI = Stopwatch.StartNew();
        Stopwatch sw_ET = Stopwatch.StartNew();
        Stopwatch sw_MQ = Stopwatch.StartNew();
        Stopwatch sw_Data = Stopwatch.StartNew();
        static Stopwatch sw_descunid = Stopwatch.StartNew();
        static Stopwatch sw_Quant = Stopwatch.StartNew();

        // CreateBillOfQuantities
        static List<Element> collectorEI = null;
        IEnumerable<ElementId> collectorET = null;
        List<Element> eTs = new List<Element>();
        string docTitle = null;
        static StreamWriter writetxt = null;

        static List<ElementType> ETypes = new List<ElementType>();

        //Additional Instance Parameters
        public static List<string> paramNamesEI = new List<string>(new string[] { "Volume", "Area", "Length" });

        internal static List<ListET> ETS = null;
        internal static List<ListEI> EIS = null;

        static bool IsMissing = true;

        //MAIN METHOD
        public void CreateBillOfQuantities(UIApplication uiapp)
        {
            //Active document in Revit application
            UIDocument ActiveUIDoc = uiapp.ActiveUIDocument;
            Document doc = uiapp.ActiveUIDocument.Document;

            #region Filters

            collectorEI = new FilteredElementCollector(doc) // returns all Instances
               .WhereElementIsNotElementType()
                .Where(e => e.IsPhysicalElement())
                .ToList();

            List<BuiltInCategory> cats = new List<BuiltInCategory>();

            foreach (Category cat in doc.Settings.Categories)
            {
                if (cat.Id.IntegerValue == (int)BuiltInCategory.OST_HVAC_Zones) continue;
                if (cat.CategoryType != CategoryType.Model) continue;
                if (!cat.CanAddSubcategory) continue;
                cats.Add((BuiltInCategory)cat.Id.IntegerValue);
            }

            ElementMulticategoryFilter multicatsFilter = new ElementMulticategoryFilter(cats, false);

            collectorET = new FilteredElementCollector(doc) // Returns all Element Types Ids
                .WhereElementIsNotElementType()
                .WherePasses(multicatsFilter)
                .Where(e => e.GetTypeId() != null && e.GetTypeId() != ElementId.InvalidElementId)
                .Where(e => doc.GetElement(e.GetTypeId()).Category != null)
                .Where(e => doc.GetElement(e.GetTypeId()).Category != null && doc.GetElement(e.GetTypeId()).Category.Name != "Piping Systems")
                .Where(e => doc.GetElement(e.GetTypeId()).Category != null && doc.GetElement(e.GetTypeId()).Category.Name != "Detail Items")
                .Where(e => e.Category != null)
                .Where(e => e.Category.CategoryType == CategoryType.Model)
                .Select(e => e.GetTypeId()) // selects and retrives the Element Type Ids
                .Distinct();

            #endregion // Filters

            #region Distinct TypeIds

            foreach (ElementId eTId in collectorET)
            {
                ElementType eT = doc.GetElement(eTId) as ElementType;
                ETypes.Add(eT);
            }

            #endregion Distinct TypeIds

            #region Report docTitle

            string docTitle = doc.Title;

            writetxt = new StreamWriter(InputData.folderPath + "//ClassReport_" + docTitle + ".txt", append: true);

            // writes report's issue date
            writetxt.WriteLine($"\r\n");
            writetxt.WriteLine($"Report " + docTitle + ".rvt created on " + DateTime.Now + ":\r\n");

            #endregion // Report docTitle

            List<ListEI> EIData = retrieveDataEI();

            List<ListET> ETData = retrieveDataET(uiapp);

            //Lauch Excel
            ExcelUtils.LauchExcel();
            ExcelUtils.PreventInteraction();

            if (InputData.instancesSheet == true)
            {
                sw_EI.Restart();
                ExcelUtils.CreateInstancesSpreadsheet(paramNamesEI, EIData);
                sw_EI.Stop();
            }
            if (InputData.elementTypesSheet == true)
            {
                sw_ET.Restart();
                ExcelUtils.CreateElementTypesSpreadsheet(ETData);
                sw_ET.Stop();
            }
            if (InputData.billofQuantitiesSheet == true)
            {
                sw_MQ.Restart();
                ExcelUtils.CreateBillofQuantitiesSpreadsheet(uiapp, ETS, docTitle);
                sw_MQ.Stop();
            }

            ExcelUtils.EnableInteraction();

            Evaluation(InputData.folderPath);
        }

        #region Evaluation Method
        public void Evaluation(string folderPath)
        {
            sw.Stop(); // stops measuring the time taken for the task

            #region Evaluation

            //Dialog Box to inform the user that there are missing elements' classifications
            if (IsMissing == false)
                TaskDialog.Show("Bill of Quantities Export",
                    string.Format("Type elements with one or more Assembly Code or Keynote parameter values are missing. " +
                    "Before running make sure to fill in all parameter values for all element types.\n" +
                    "For more info check the report: " + folderPath + @"\ClassReport_" + docTitle + ".txt\n" +
                    "The export is finished. Time: " + sw.Elapsed.TotalSeconds + " seconds.\n" +
                    //"Tempo Tabelas EI e ET " + sw_EI_ET.Elapsed.TotalSeconds +
                    "\nTime table MQ " + sw_MQ.Elapsed.TotalSeconds + ":\nTime Data " + sw_Data.Elapsed.TotalSeconds +
                    ":\nTime desc unid " + sw_descunid.Elapsed.TotalSeconds + ", Time Quant " + sw_Quant.Elapsed.TotalSeconds));

            //Dialog Box to inform the user that all elements' classifications have been found
            else
            {
                TaskDialog.Show("Bill of Quantities Export",
                    string.Format("All Element Type's Assembly Code's and Keynote's were found.\n" +
                    "For more info check the report: " + folderPath + @"\ClassReport_" + docTitle + ".txt\n" +
                    "The export is finished. Time: " + sw.Elapsed.TotalSeconds + " seconds.\n" +
                    //"Time tables EI and ET " + sw_EI_ET.Elapsed.TotalSeconds +
                    "\nTime table MQ " + sw_MQ.Elapsed.TotalSeconds + ":\nTime Data " + sw_Data.Elapsed.TotalSeconds +
                    ":\nTime desc unid " + sw_descunid.Elapsed.TotalSeconds + ", Time Quant " + sw_Quant.Elapsed.TotalSeconds));

                // writes on the .txt report that there are no parameters to classify
                writetxt.WriteLine($"No parameter to classify\n");
            }

            writetxt.WriteLine($"\r\n");
            writetxt.Close(); // closes the report

            #endregion // Evaluation
        }

        #endregion //Evaluation Method

        #region Data EI

        public static List<ListEI> retrieveDataEI()
        {
            List<ElementId> ElementIdEI = new List<ElementId>();

            EIS = new List<ListEI>();

            foreach (Element eI in collectorEI)
            {
                var EI = new ListEI
                {
                    ID = eI.Id.IntegerValue,
                    IsType = eI is ElementType ? 1 : 0,
                    CategoryName = eI.Category.Name,
                    TypeName = eI.Name,
                    TypeNameId = eI.GetTypeId().IntegerValue
                };

                try // eI is a Family Instance - eI as FamilySymbol to get its FamilyName
                {
                    FamilyInstance eIFamilyInstance = eI as FamilyInstance;
                    FamilySymbol eIFamilySymbol = eIFamilyInstance.Symbol;
                    Family eIFamily = eIFamilySymbol.Family;
                    string eIFamilyName = eIFamily.Name;
                    EI.FamilyName = eIFamilyName;
                }
                catch // eI is not a Family Instance
                {
                    EI.FamilyName = "*NA*";
                }

                foreach (string paramName in paramNamesEI)
                {
                    string paramValue = "*NA*";

                    Parameter p = eI.LookupParameter(paramName);

                    if (p != null) paramValue = GetParameterValue(p);

                    EI.GetType().GetProperty(paramName).SetValue(EI, paramValue);
                }
                ElementId eTId = eI.GetTypeId();
                ElementIdEI.Add(eTId); //ãdds eIId to the list elementIdEI  - this list will be later used to create the Element Types table

                EIS.Add(EI);
            }

            //Sorts EIS by TypeNameId
            var EISSorted = EIS.AsQueryable().OrderBy(eI => eI.TypeNameId).ToList();

            return EISSorted;
        }

        #endregion // Data EI

        #region Data ET

        public static List<ListET> retrieveDataET(UIApplication uiapp)
        {

            ETS = new List<ListET>();

            foreach (Element eT in ETypes)
            {
                var ET = new ListET();

                try
                {
                    ET.ID = eT.Id.IntegerValue;
                    ET.IsType = eT is ElementType ? 1 : 0;
                    ET.CategoryName = eT.Category.Name;
                    ET.CategoryId = eT.Category.Id.IntegerValue;
                    ET.TypeName = eT.Name;
                }
                catch
                {
                    ET.ID = -1;
                    ET.IsType = eT is ElementType ? 1 : 0;
                    ET.CategoryName = "*NA*";
                    ET.CategoryId = -1;
                    ET.TypeName = "*NA*";
                }

                try
                {
                    FamilySymbol eTFamilySymbol = eT as FamilySymbol; // FamilySymbol of eT
                    string eTFamilyName = eTFamilySymbol.FamilyName; //  FamilyName of eTFamilySymbol
                    ET.FamilyName = eTFamilyName;
                }
                catch
                {
                    ET.FamilyName = "*NA*"; // eT does not have a FamilyName
                }

                // new ListEI with all instances of Type Id
                List<Element> ListEI = new List<Element>();
                try
                {
                    ListEI = collectorEI.Where(q => q.GetTypeId() == eT.Id).ToList();
                }
                catch
                {
                    // When the Id is -1 it means the elements belong to the Parts Category
                    ListEI = collectorEI.Where(q => q.GetTypeId().IntegerValue == -1).ToList();
                }

                ET.Quantity = ListEI.Count();

                foreach (string paramName in paramNamesEI)
                {
                    foreach (Element eI in ListEI)
                    {
                        if (eI.LookupParameter(paramName) != null)
                        {
                            Parameter p = eI.LookupParameter(paramName);

                            double paramValue = Convert.ToDouble(GetParameterValue(p));

                            double value = Convert.ToDouble(ET.GetType().GetProperty("Total" + paramName).GetValue(ET)); // gets the value of the ET property
                            ET.GetType().GetProperty("Total" + paramName).SetValue(ET, (value + paramValue).ToString()); // defines the value of the ET property
                        }
                    }
                }

                // Element Type's cost per unit
                try
                {
                    ET.Costunit = GetParameterValue(eT.LookupParameter("Cost"));
                }
                catch
                {
                    ET.Costunit = "*NA*";
                }

                #region Classification

                ET.AssemblyCode = GetBuiltInParamValue(eT, BuiltInParameter.UNIFORMAT_CODE);
                ET.AssemblyDesc = GetBuiltInParamValue(eT, BuiltInParameter.UNIFORMAT_DESCRIPTION);
                ET.KeyValue = GetBuiltInParamValue(eT, BuiltInParameter.KEYNOTE_PARAM);

                ET.KeyText = GetKeynoteText(ET.KeyValue, uiapp);
                if (ET.KeyText == null || ET.KeyText == "")
                {
                    ET.KeyText = "MISSING";
                }

                #endregion // Classification

                ETS.Add(ET);
            }

            //Sorts ETS lists by ID
            var ETSSorted = ETS.AsQueryable().OrderBy(eT => eT.ID).ToList();

            return ETSSorted;
        }
        #endregion Data ET

        #region Data MQ

        public static List<ListMQ> RetrieveMQData(UIApplication uiapp, List<ListET> ETS)
        {
            List<ListMQ> MQS = new List<ListMQ>();

            sw_descunid.Reset();
            sw_Quant.Reset();

            int index = 0;

            foreach (Element eT in ETypes)
            {
                List<Element> ListEI = collectorEI.Where(q => q.GetTypeId() == eT.Id).ToList();  // new ListEI with all instances of Type Id

                var MQ = new ListMQ();

                if (eT != null)
                {
                    if (eT.Category.Name == "Mass") // Mass Category
                    {
                        MQ.AssemblyCode = "*NA*";
                        MQ.AssemblyDesc = "MISCELLANEOUS VOLUMETRIES";
                        MQ.KeyValue = "Mass";
                        MQ.KeyText = "Miscellaneous.";
                    }

                    MQ.AssemblyCode = GetBuiltInParamValue(eT, BuiltInParameter.UNIFORMAT_CODE);
                    MQ.AssemblyDesc = GetBuiltInParamValue(eT, BuiltInParameter.UNIFORMAT_DESCRIPTION);
                    MQ.KeyValue = GetBuiltInParamValue(eT, BuiltInParameter.KEYNOTE_PARAM);

                    MQ.KeyText = GetKeynoteText(MQ.KeyValue, uiapp);
                    if (MQ.KeyText == null || MQ.KeyText == "")
                    {
                        MQ.KeyText = "MISSING";
                    }

                    // eT parameters values were not found created
                    else if (MQ.AssemblyCode == "MISSING" && MQ.AssemblyDesc == "MISSING" || MQ.KeyValue == "MISSING" && MQ.KeyText == "MISSING")
                    {
                        writetxt.WriteLine($"The Element Type " + eT.Name + " with Id " + eT.Id.IntegerValue + " does not have a value for the Assembly Code or the Keynote\n");
                    }

                    //Units and Costs
                    sw_Quant.Start();

                    if (MQ.Unit == null || MQ.Unit.ToString() == "") MQ.Unit = "vg";

                    string instances = null;
                    string paramValue = null;

                    switch (MQ.Unit.ToString())
                    {
                        // converts the parameter value
                        case "m3":
                            string totalVolume = ETS[index].GetType().GetProperty("TotalVolume").GetValue(ETS[index]).ToString();
                            paramValue = totalVolume;
                            break;
                        case "m2":
                            string totalArea = ETS[index].GetType().GetProperty("TotalArea").GetValue(ETS[index]).ToString();
                            paramValue = totalArea;
                            break;
                        case "m":
                            string totalLength = ETS[index].GetType().GetProperty("TotalLength").GetValue(ETS[index]).ToString();
                            paramValue = totalLength;
                            break;
                        case "vg":
                            instances = Convert.ToInt32(ETS[index].GetType().GetProperty("Quantity").GetValue(ETS[index])).ToString();
                            paramValue = instances;
                            break;
                        default:
                            paramValue = instances;
                            break;
                    }

                    MQ.GetType().GetProperty("Quant").SetValue(MQ, paramValue);

                    index++;

                    sw_Quant.Stop();

                    // Price per unit
                    MQ.PrUnit = GetParameterValue(eT.LookupParameter("Cost"));

                    // Partial costs
                    double PartialCost = Convert.ToDouble(MQ.Quant) * Convert.ToDouble(MQ.PrUnit);
                    MQ.Partial = PartialCost.ToString();

                    MQS.Add(MQ);
                }
            }

            return MQS;
        }

        #endregion Data MQ

        #region GetKeynoteTable Method
        public static KeyBasedTreeEntries GetKeynoteEntries(UIApplication uiapp)
        {

            Document doc = uiapp.ActiveUIDocument.Document;

            KeynoteTable Kt = KeynoteTable.GetKeynoteTable(doc);

            KeyBasedTreeEntries kbte = Kt.GetKeyBasedTreeEntries();

            return kbte;
        }

        #endregion //GetKeynoteTable Method

        #region GetBuiltInParamValue Method

        public static string GetBuiltInParamValue(Element eT, BuiltInParameter bip)
        {
            Parameter p = eT.get_Parameter(bip);
            string pValue = null;

            if (p != null)
            {
                pValue = GetParameterValue(p);
                if(pValue == "" || pValue == null)
                {
                    IsMissing = false;
                    return "MISSING";
                }
                return pValue;
            }
            return "*NA*";
        }

        #endregion //GetBuiltInParamValue Method

        #region GetKeynote Method

        public static string GetKeynoteText(string keyValue, UIApplication uiapp)
        {
            KeyBasedTreeEntries kbte = GetKeynoteEntries(uiapp);

            IEnumerable<KeyBasedTreeEntry> keyValues;

            string keynoteText = null;

            keyValues = kbte.Where(k => k.Key.Equals(keyValue));

            foreach (KeynoteEntry k in keyValues)
            {
                keynoteText = k.KeynoteText;
            }
            if (keynoteText != "")
            {
                return keynoteText;
            }
            else
            {
                return "MISSING";
            }
        }

        #endregion //GetKeynote

        #region GetParameterValue and RealString Method

        public static string GetParameterValue(Parameter param)
        {
            string s;
            switch (param.StorageType)
            {
                case StorageType.Double:
                    //
                    // the internal database unit for all lengths is feet.
                    // for instance, if a given room perimeter is returned as
                    // 102.36 as a double and the display unit is millimeters,
                    // then the length will be displayed as
                    // peri = 102.36220472440
                    // peri * 12 * 25.4
                    // 31200 mm
                    //
                    //s = param.AsValueString(); // value seen by user, in display units
                    //s = param.AsDouble().ToString(); // if not using not using LabUtils.RealString()
                    s = RealString(param.AsDouble()); // raw database value in internal units, e.g. feet
                    break;

                case StorageType.Integer:
                    s = param.AsInteger().ToString();
                    break;

                case StorageType.String:
                    s = param.AsString();
                    break;

                case StorageType.ElementId:
                    s = param.AsElementId().IntegerValue.ToString();
                    break;

                case StorageType.None:
                    s = "?NONE?";
                    break;

                default:
                    s = "?ELSE?";
                    break;
            }
            return s;
        }

        public static string RealString(double a)
        {
            return a.ToString("0.##");
        }

        #endregion SetValue, GetParameterValue and RealString Method
    }
    #region Extensions class

    public static class Extensions
    {
        //Method to check for Element Instances
        public static bool IsPhysicalElement(this Element e) // definiçao do metodo estatico IsPhysicalElement com um argumento e
        {
            if (e.Category == null) return false; // if the element category is null, the element is not an instance
            if (e.Category.Name.ToString() == "HVAC Zones") return false; // if the category name is "HVAC Zones", the element is not an instance
            if (e.ViewSpecific) return false; // if the element is view specific, the element is not an instance
            return e.Category.CategoryType == CategoryType.Model && e.Category.CanAddSubcategory; //  returns all instance elements that have category and subcategoria
        }
    }

    #endregion Extensions class
}
