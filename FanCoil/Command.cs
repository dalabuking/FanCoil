#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.CSharp;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace FanCoil
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

            List<Object> oVentilo = new List<Object>();


            Form1 form1 = new Form1(commandData);
            form1.ShowDialog();
            

            if (form1.isCancle == true){

                return Result.Cancelled;
            }

            string filepath = form1.fileName;

           FilteredElementCollector Spaces = GetSpaces(uiapp);
           List<Parameter> oCoolingLoadList = GetParameterList(Spaces, "Design Cooling Load");
           List<Parameter> oHeatingLoadList = GetParameterList(Spaces, "Design Heating Load");
            try
            {
                 oVentilo = CreateAllVentilokonvektors(filepath);
            }
            catch
            {
                TaskDialog.Show("Not Fount", "Please insert correct path to excel file!");
                return Result.Cancelled;
            }
           
           List<Object> chosenVentilos = PickVentilo(oVentilo, oCoolingLoadList, oHeatingLoadList);
            //writeResultsToExcel(Spaces ,chosenVentilos);
            List<FamilyInstance> oVentiloFamily = placeFamilyToRoom(Spaces, chosenVentilos, uiapp);
            changeParameters(oVentiloFamily, chosenVentilos, uiapp);


            return Result.Succeeded;
        }
        public FilteredElementCollector GetSpaces(UIApplication oApp)
        {
           
            Document doc = oApp.ActiveUIDocument.Document;
            FilteredElementCollector oSpaces = new FilteredElementCollector(doc).OfClass(typeof(SpatialElement));
            return oSpaces;
        }

        public List<Parameter> GetParameterList(FilteredElementCollector oSpaces, string parameterString)
        {
            List<Parameter> oParameterList = new List<Parameter>();
            foreach (Element element in oSpaces)
            {
                Parameter para = element.LookupParameter(parameterString);

                oParameterList.Add(para);
            }
            return oParameterList;
        }
        public List<Object> CreateAllVentilokonvektors(string fileName)
        {
            List<Object> ventilokonvektorList = new List<Object>();
            string path = fileName;
            var excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            Excel.Worksheet techdata = (Excel.Worksheet)wb.Worksheets[1];


            excelApp.Visible = false;

            var modelValue = techdata.Range["B2", "B19"];
            var headload = techdata.Range["O3", "O19"];
            var coolingload = techdata.Range["P3", "P19"];
            var depths = techdata.Range["Q3", "Q19"];
            var widths = techdata.Range["R3", "R19"];
            var heights = techdata.Range["S3", "S19"];

            int i = 0;
            foreach (Excel.Range dataa in modelValue)
            {

                Excel.Range dep = (Excel.Range)depths[i];
                Excel.Range heig = (Excel.Range)heights[i];
                Excel.Range wid = (Excel.Range)widths[i];
                Excel.Range cool = (Excel.Range)coolingload[i];
                Excel.Range heat = (Excel.Range)headload[i];

                var ventilo = new Ventilokonvektor(dataa.Value.ToString());
                ventilo.Heatingload = Convert.ToDouble(heat.Value.ToString());
                ventilo.Coolingload = Convert.ToDouble(cool.Value.ToString());
                ventilo.Height = Convert.ToDouble(heig.Value.ToString());
                ventilo.Depth = Convert.ToDouble(dep.Value.ToString());
                ventilo.Width = Convert.ToDouble(wid.Value.ToString());

                ventilokonvektorList.Add(ventilo);
                i = i + 1;
            }
            wb.Close();
            excelApp.Quit();
            return ventilokonvektorList;
        }

        public List<Object> PickVentilo(List<Object> oVentilos, List<Parameter> oCoolingLoadLis, List<Parameter> oHeatingLoadList)
        {
            int i = 0;

            List<Object> chosenVentilos = new List<Object>();
            foreach (Parameter coolPara in oCoolingLoadLis)
            {
                string stringCoolValue = coolPara.AsValueString();
                char[] removeChar = { '-', 'W', ' ' };
                string finalCoolValue = stringCoolValue.Trim(removeChar);
                double doubleCoolValue = Convert.ToDouble(Convert.ToInt32(finalCoolValue));
                Parameter heatPara = (Parameter)oHeatingLoadList[i];
                string stringHeatValue = heatPara.AsValueString();
                string finalHeatValue = stringHeatValue.Trim(removeChar);
                double doubleHeatValue = Convert.ToDouble(Convert.ToInt32(finalHeatValue));

                Object choosenVentilo = GetCorrectVentilo(oVentilos, doubleHeatValue, doubleCoolValue);
                chosenVentilos.Add(choosenVentilo);


                i = i + 1;
            }
            return chosenVentilos;
        }

        public Object GetCorrectVentilo(List<Object> oVentilos, double heat, double cool)
        {
            Ventilokonvektor chosenVentilo = new Ventilokonvektor("Cant Find Solution!");

            foreach (Ventilokonvektor ventilo in oVentilos)
            {
                if ((ventilo.Heatingload > heat) && (ventilo.Coolingload > cool))
                {
                    chosenVentilo = ventilo;
                    break;
                }
            }
            double multi = 1;
            while (chosenVentilo.Model == "Cant Find Solution!")
            {
                multi = multi + 1;
                foreach (Ventilokonvektor oVentilo in oVentilos)
                {

                    if (((multi * oVentilo.Heatingload > heat)) && ((multi * oVentilo.Coolingload) > cool))
                    {


                        string newModel = oVentilo.Model;
                        string Combine = newModel + " x" + multi.ToString();
                        chosenVentilo.Heatingload = oVentilo.Heatingload;
                        chosenVentilo.Coolingload = oVentilo.Coolingload;
                        chosenVentilo.Width = oVentilo.Width;
                        chosenVentilo.Depth = oVentilo.Depth;
                        chosenVentilo.Height = oVentilo.Height;
                        chosenVentilo.Model = Combine;
                        break;
                    }
                }

            }


            return chosenVentilo;
        }

        public void writeResultsToExcel(FilteredElementCollector Spaces, List<Object> choosenVentilos)
        {
            List<Object> ventilokonvektorList = new List<Object>();
            string path = @"D:\Revit Projects\REVIT\Results.xlsx";
            var excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(path, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            Excel.Worksheet results = (Excel.Worksheet)wb.Worksheets[1];

            List<Parameter> oCoolingLoadList = GetParameterList(Spaces, "Design Cooling Load");
            List<Parameter> oHeatingLoadList = GetParameterList(Spaces, "Design Heating Load");
            List<Parameter> oRoomName = GetParameterList(Spaces, "SC_Roomdescription");
            int i = 2;
            int index = 0;

            foreach (Ventilokonvektor ventilo in choosenVentilos)
            {
                TaskDialog.Show("123", "1");
                Parameter oCool = (Parameter)oCoolingLoadList[index];
                Parameter oHeat = (Parameter)oHeatingLoadList[index];
                Parameter oRoom = (Parameter)oRoomName[index];
                TaskDialog.Show("123", "2");
                results.Cells[i, 1] = oRoom.AsValueString().ToString();
                results.Cells[i, 2] = oHeat.AsValueString().ToString();
                results.Cells[i, 3] = oCool.AsValueString().ToString();
                results.Cells[i, 4] = ventilo.Model.ToString();
                results.Cells[i, 5] = ventilo.Heatingload.ToString();
                results.Cells[i, 6] = ventilo.Coolingload.ToString();

                

                i = i + 1;
                index = index + 1;
            }
           
            wb.Save();
            wb.Close();
            excelApp.Quit();
        }

        public List<FamilyInstance> placeFamilyToRoom(FilteredElementCollector Spaces, List<Object> chosenVentilos, UIApplication oApp)
        {
          
            Document doc = oApp.ActiveUIDocument.Document;
            Element AirTerm = getFamilySymbol("Air Terminal", oApp);
            FamilySymbol AirTerminalSymbol = AirTerm as FamilySymbol;
            List<FamilyInstance> ventiloList = new List<FamilyInstance>();
            foreach (Element oSpace in Spaces)
            {

                LocationPoint roomCenter = oSpace.Location as LocationPoint;
                
                XYZ xyzPos = roomCenter.Point;

                Level lvl = getLevelById(oSpace.LevelId, doc);

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Placing Air Terminals");
                   
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyzPos, AirTerminalSymbol, lvl, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    ventiloList.Add(instance);
                    t.Commit();
                }

            }
            return ventiloList;
        }

        public Level getLevelById(ElementId lvlid, Document doc)
        {
            Level lvl = null;
            FilteredElementCollector lvlCol = new FilteredElementCollector(doc).OfClass(typeof(Level));
            foreach(Element el in lvlCol)
            {
                lvl = el as Level;
               if (lvl.Id == lvlid)
                {
                    return lvl;
                }
            }
            return lvl;

        }

        public Element getFamilySymbol(string FName, UIApplication oApp)
        {
           
            Document doc = oApp.ActiveUIDocument.Document;

            FilteredElementCollector oSymbols = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol));
            Element AirTerminal = oSymbols.FirstElement();
            foreach (Element oSy in oSymbols)
            {
                if (oSy.Name == FName)
                {
                    AirTerminal = oSy;
                }
            }
            return AirTerminal;
        }

        public void changeParameters(List<FamilyInstance> oVentiloFamily, List<Object> chosenVentilos, UIApplication oApp)
        {
            int i = 0;
            
            Document doc = oApp.ActiveUIDocument.Document;
            foreach (Ventilokonvektor oVentilo in chosenVentilos)
            {
                FamilyInstance oEl = oVentiloFamily[i];


                Parameter w = oEl.LookupParameter("Width");
                Parameter d = oEl.LookupParameter("Depth");
                Parameter h = oEl.LookupParameter("Height");
                Parameter CL = oEl.LookupParameter("Cooling Load");
                Parameter HL = oEl.LookupParameter("Heating Load");
                Parameter M = oEl.LookupParameter("_Model");
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Changing Parameters");

                    w.Set(oVentilo.Width / 304.8);
                    h.Set(oVentilo.Height / 304.8);
                    d.Set(oVentilo.Depth / 304.8);
                    CL.Set(oVentilo.Heatingload * 10);
                    HL.Set(oVentilo.Coolingload * 10);
                    M.Set(oVentilo.Model.ToString());


                    t.Commit();
                }




                i = i + 1;

            }

        }





    }
	public class Ventilokonvektor
    {
        public double Coolingload;
        public double Heatingload;
        public double Height;
        public double Width;
        public double Depth;
        public string Model;

        public Ventilokonvektor(string _model)
        {
            Model = _model;

        }

    }
}
