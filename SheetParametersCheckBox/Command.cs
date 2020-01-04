#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace SheetParametersCheckBox
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

            List<string> sheetNumberList = new List<string>();

            FilteredElementCollector fecSheets = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType();

            foreach (ViewSheet viewSheet in fecSheets)
            {
                sheetNumberList.Add(viewSheet.SheetNumber);
            }

            using (var form = new Form1())
            {
                //if the user hits cancel just drop out
                if (form.DialogResult == System.Windows.Forms.DialogResult.Cancel)
                {
                    return Result.Cancelled;
                }

                //assing the sheet number list to the check list source
                form.checkedListSource = sheetNumberList;

                //use ShowDialog to show the form as a modal dialog box. 
                form.ShowDialog();

            }

            return Result.Succeeded;
        }
    }
}
