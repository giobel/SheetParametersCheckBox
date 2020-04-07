
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Linq;
using System.Text;
using System;



namespace SheetParametersCheckBox
{
    [Transaction(TransactionMode.Manual)]
    public class SheetCopier : IExternalCommand
    {
        public ElementId TitlebId { get; set; }
        public FamilySymbol TitleSym { get; set; }
        public List<String> Sheets = new List<string>();
        public List<String> ProjectSheetNo = new List<string>();

        public string CheckedByText { get; set; }
        public string DrawnByText { get; set; }
        public string PassedByText { get; set; }
        public string Zone { get; set; }
        public string Level { get; set; }
        public string Role { get; set; }
        public string ProjectCode { get; set; }
        public string Originator { get; set; }
        public string Status { get; set; }
        public string Type { get; set; }
        public string Date { get; set; }
        public string Scale1 { get; set; }
        public string SheetName { get; set; }
        public string TITLE2 { get; set; }

        public StringBuilder sb = new StringBuilder();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UIDocument and Document
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            FilteredElementCollector fecSheets = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType();

            foreach (ViewSheet sheet in fecSheets)
            {
                String a = sheet.SheetNumber;
                ProjectSheetNo.Add(a);
            }


            IList<Element> alltitleblocks = new List<Element>();
            IList<Element> ElementsOnSheet = new List<Element>();
            FamilySymbol Titleblock;

            uidoc.RefreshActiveView();
            View currentView = uidoc.ActiveView;

            ViewSheet Vfam = (ViewSheet)doc.GetElement(currentView.Id);

            //NOTE: Should these be extracted only if the user tick the "Copy Sheet Parameters" checkbox?

            //CheckedByText = GetParameterValue(Vfam, "Checked By");
            //DrawnByText = GetParameterValue(Vfam, "Drawn By");
            //PassedByText = GetParameterValue(Vfam, "Passed By");
            //Zone = GetParameterValue(Vfam, "ZONE");
            //Level = GetParameterValue(Vfam, "LEVEL");
            //Role = GetParameterValue(Vfam, "ROLE");
            //ProjectCode = GetParameterValue(Vfam, "Project Code");
            //Originator = GetParameterValue(Vfam, "Originator");
            //Status = GetParameterValue(Vfam, "Drawing Status");
            //Date = GetParameterValue(Vfam, "Date Drawn");
            //Scale1 = GetParameterValue(Vfam, "Scale - Manual");
            //SheetName = GetParameterValue(Vfam, "Sheet Name");
            //TITLE2 = GetParameterValue(Vfam, "TITLE 2");

            
            


            // get elements on viewsheet copuld use current view instead
            foreach (Element e in new FilteredElementCollector(doc).OwnedByView(Vfam.Id))
            {
                ElementsOnSheet.Add(e);
            }

            // get all titleblocks 

            FilteredElementCollector Collector = new FilteredElementCollector(doc);
            Collector.OfClass(typeof(FamilySymbol));
            Collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            alltitleblocks = Collector.ToElements();


            //get family symbol of sheet

            foreach (Element el in ElementsOnSheet)
            {
                foreach (FamilySymbol Fs in alltitleblocks)
                {
                    if (el.GetTypeId().IntegerValue == Fs.Id.IntegerValue)
                    {

                        Titleblock = Fs;
                        TitleSym = Fs;

                        //TaskDialog.Show("Type  Name", a.ToString() + TitlebId.ToString());
                    }
                }

            }

            Dictionary<string, XYZ> Schedules = new Dictionary<string, XYZ>();
            List<ViewSchedule> ViewSchedules = new List<ViewSchedule>();

            //Empty Dictionaries to store the schedules,General Annotations & legends
            Dictionary<ViewSchedule, XYZ> SchedulesAndPoints = new Dictionary<ViewSchedule, XYZ>();
            Dictionary<Element, XYZ> GeneralAnnotationOnSheet = new Dictionary<Element, XYZ>();
            Dictionary<View, XYZ> SheetLegends = new Dictionary<View, XYZ>();

            GetSheetSchedulesAndPoint(Vfam, doc, SchedulesAndPoints);
            GetSheetAnnotationsAndPoint(doc, GeneralAnnotationOnSheet);
            GetSheetLegendAndPoint(Vfam, doc, SheetLegends);

            string sheetNumberAndName = $"{Vfam.SheetNumber} - {currentView.Name}";

            using (var form = new Form8(sheetNumberAndName))


            {
                
                form.ShowDialog();

                if (form.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    if (form.NewSheetNoText.Contains('-'))
                    {

                        string[] Ranges = null;

                        Char mychar = '-';

                        Ranges = form.NewSheetNoText.Split(mychar);

                        string b = Ranges[0];
                        string c = Ranges[1];

                        int startindex = int.Parse(b);
                        int endindex = int.Parse(c);

                        int count = 0;

                        string fmt = "0000.##";

                        for (count = startindex; count < endindex + 1; count++)
                        {
                            Sheets.Add(count.ToString(fmt));
                        }

                    }
                    else if (form.NewSheetNoText.Contains(','))
                    {
                        string[] Range1 = null;
                        char mychar1 = ',';
                        Range1 = form.NewSheetNoText.Split(mychar1);



                        for (int i = 0; i < Range1.Length; i++)
                        {
                            if (Range1[i].Length != 4)
                            {
                                Sheets.Add("0" + Range1[i].ToString());

                            }
                            else
                            {
                                Sheets.Add(Range1[i].ToString());
                            }


                        }
                    }

                    else if (form.NewSheetNoText.Length == 4)
                    {

                        Sheets.Add(form.NewSheetNoText);

                    }
                    else
                    {
                        TaskDialog.Show("Warning", "Sheet numbers entered not valid");

                    }

                }

                using (Transaction t = new Transaction(doc, "Sheet Copier"))
                {
                    t.Start();

                    try
                    {


                        if (form.Sheetparams == false)
                        {
                            foreach (string s in Sheets)

                                if (ProjectSheetNo.Contains(s))

                                {
                                    sb.AppendLine(s);
                                }
                                else
                                {
                                    ViewSheet sheet = ViewSheet.Create(doc, TitleSym.Id);
                                    sheet.SheetNumber = s;


                                    //method for adding schedules
                                    InsertSchedulesOnSheet(doc, sheet, SchedulesAndPoints);
                                    
                                    //method for legends
                                    InsertLegendOnSheet(doc, sheet, SheetLegends);

                                    //method for annotation
                                    InsertAnnotationOnSheet(doc, sheet, GeneralAnnotationOnSheet);







                                }
                        }
                        else
                        {
                            foreach (string s in Sheets)

                                if (ProjectSheetNo.Contains(s))

                                {
                                    sb.AppendLine(s);
                                }
                                else
                                {
                                    ViewSheet sheet = ViewSheet.Create(doc, TitleSym.Id);
                                    sheet.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(s);
                                    sheet.LookupParameter("Checked By").Set(CheckedByText);
                                    sheet.LookupParameter("Checked By").Set(CheckedByText);
                                    sheet.LookupParameter("Drawn By").Set(DrawnByText);
                                    sheet.LookupParameter("Passed By").Set(PassedByText);
                                    sheet.LookupParameter("ZONE").Set(Zone);
                                    sheet.LookupParameter("LEVEL").Set(Level);
                                    sheet.LookupParameter("ROLE").Set(Role);
                                    sheet.LookupParameter("Project Code").Set(ProjectCode);
                                    sheet.LookupParameter("Originator").Set(Originator);
                                    sheet.LookupParameter("Drawing Status").Set(Status);
                                    sheet.LookupParameter("Type").Set(Type);
                                    sheet.LookupParameter("Date Drawn").Set(Date);
                                    sheet.LookupParameter("Scale - Manual").Set(Scale1);
                                    sheet.LookupParameter("Sheet Name").Set(SheetName);
                                    sheet.LookupParameter("TITLE 2").Set(TITLE2);

                                    //method for adding schedules
                                    InsertSchedulesOnSheet(doc, sheet, SchedulesAndPoints);

                                    //method for legends
                                    InsertLegendOnSheet(doc, sheet, SheetLegends);

                                    //method for annotation
                                    InsertAnnotationOnSheet(doc, sheet, GeneralAnnotationOnSheet);

                                }
                        }

                        

                        t.Commit();
                        uidoc.RefreshActiveView();
                        
                       
                    }

                    catch (Exception ex)
                    {
                        TaskDialog.Show("Error", ex.Message);


                    }

                    if (sb.Length > 0)
                    {
                        TaskDialog.Show("Warning", "cant add sheet No" + sb + "to project as is already in use");

                        
                    }


                   
                }

                return Result.Succeeded;
                
            }
        }



        public string GetParameterValue(Element elems, string ParamName)

                {
                    string paramValue = "";


                    Parameter p = elems.LookupParameter(ParamName);
                    StorageType parameterType = p.StorageType;

                    if (StorageType.Double == parameterType)

                    {
                        return paramValue = UnitUtils.ConvertFromInternalUnits(p.AsDouble(), DisplayUnitType.DUT_MILLIMETERS).ToString();
                        /*askDialog.Show("parameter value = {0}", paramValue);*/
                    }

                    else if (StorageType.String == parameterType)
                    {
                        return paramValue = p.AsString();
                    }

                    else if (StorageType.Integer == parameterType)
                    {
                        return paramValue = p.AsInteger().ToString();
                    }

                    else if (StorageType.ElementId == parameterType)
                    {
                        return paramValue = p.AsValueString();
                    }

                    else
                    {
                        return "";

                    }

                  

                }




        
        public void GetSheetSchedulesAndPoint(ViewSheet Vfam, Document doc, Dictionary<ViewSchedule, XYZ> Schedules)
        {

            var shdoc = Vfam.Document;
            FilteredElementCollector collector = new FilteredElementCollector(shdoc, Vfam.Id);
            var scheduleSheetInstances = collector.OfClass(typeof(ScheduleSheetInstance)).ToElements().OfType<ScheduleSheetInstance>();


            foreach (var scheduleSheetInstance in scheduleSheetInstances)
            {
               
                XYZ Point = scheduleSheetInstance.Point;
                var scheduleId = scheduleSheetInstance.ScheduleId;

                if (scheduleId == ElementId.InvalidElementId)
                    continue;

                var viewSch = doc.GetElement(scheduleId) as ViewSchedule;
                if (viewSch != null && !viewSch.Name.Contains( "<Revision Schedule>") )
                {
                    XYZ a = Point;
                    Schedules.Add(viewSch, a);
                  
                }
                else
                {
                    //do nothing
                }
            }
        }
 

        public void GetSheetAnnotationsAndPoint( Document doc, Dictionary<Element, XYZ>AnnotationOnSheet)

        {
            FilteredElementCollector allElementsInView = new FilteredElementCollector(doc, doc.ActiveView.Id);
            IList<Element> ElementsinView = allElementsInView.ToElements();

            
            

            foreach (Element ele in allElementsInView)
            {
                if (ele.Category.Name == "Generic Annotations")
                {
                    LocationPoint a = ele.Location as LocationPoint;

                    XYZ c = a.Point;

                    AnnotationOnSheet.Add(ele, c);


                }

            }
        }

        public void GetSheetLegendAndPoint(ViewSheet Vfam, Document doc, Dictionary<View, XYZ> Legend)
        {
            ICollection<ElementId> viewports = Vfam.GetAllViewports();
            {
                foreach (ElementId v in viewports)
                {
                    Viewport a = doc.GetElement(v) as Viewport;
                    XYZ center = a.GetBoxCenter();
                    ElementId b = a.ViewId;
                    View c = doc.GetElement(b) as View;

                    if (c.ViewType == ViewType.Legend)
                    {
                        //String d = c.ViewType.ToString();
                        //String e = c.Name;

                        Legend.Add(c, center);

                    }

                    else
                    {
                        //do nothing
                    }

                }

            }
        }

       
        //places legends from a dictonary on to a sheet

        public void InsertLegendOnSheet(Document doc, ViewSheet sheet, Dictionary<View, XYZ> Legend)


        {
            foreach (KeyValuePair<View,XYZ> leg in Legend)
            {
                View Lview = leg.Key;
                XYZ point = leg.Value;

                Viewport.Create(doc, sheet.Id, Lview.Id, point);

            }

          
        }

        
        //place annotation from a dictionary on to a sheet

        public void InsertAnnotationOnSheet(Document doc, ViewSheet sheet, Dictionary<Element, XYZ> Annotation)
        {

            foreach (KeyValuePair<Element, XYZ> Anno in Annotation)
            {
               
                    Element e = Anno.Key;
                    FamilyInstance f = e as FamilyInstance;
                    FamilySymbol FamilySymb = f.Symbol;
                    XYZ Point = Anno.Value;

                    doc.Create.NewFamilyInstance(Point, FamilySymb, sheet);

                
            }
        }

        //  place schedules from dictionary on to a sheet

        public void InsertSchedulesOnSheet(Document doc, ViewSheet sheet, Dictionary<ViewSchedule, XYZ> Schedules)
        {


            foreach (KeyValuePair<ViewSchedule, XYZ> schedule in Schedules)
            {
                View sched = schedule.Key;
                XYZ point = schedule.Value;

                ScheduleSheetInstance.Create(doc, sheet.Id, sched.Id, point);
            }



                
        }


    }



}










                                   





