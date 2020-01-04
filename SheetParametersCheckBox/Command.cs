#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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

            Dictionary<string, Element> sheetElements = new Dictionary<string, Element>();

            FilteredElementCollector fecSheets = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType();

            foreach (ViewSheet viewSheet in fecSheets)
            {
                sheetNumberList.Add(viewSheet.SheetNumber);
                sheetElements.Add(viewSheet.SheetNumber, viewSheet);
            }

            using (var form = new Form1())
            {

                //assing the sheet number list to the check list source
                form.checkedListSource = sheetElements.Keys.ToList();

                //use ShowDialog to show the form as a modal dialog box. 
                form.ShowDialog();

                //if the user hits cancel just drop out
                if (form.DialogResult == System.Windows.Forms.DialogResult.Cancel)
                {
                    return Result.Cancelled;
                }

                using (Transaction t  = new Transaction(doc, "Set sheet numbers"))
                {
                    t.Start();
                    try
                    {
                        foreach (string sheetNumber in form.checkedItems)
                        {
                            if (form.CheckedByText != null && form.CheckedByText.Length > 0)
                            {
                                sheetElements[sheetNumber].LookupParameter("Checked By").Set(form.CheckedByText);
                            }
                        }
                    }
                    catch(Exception ex)
                    {
                        TaskDialog.Show("Error", ex.Message);
                    }
                    finally
                    {
                        t.Commit();
                    }

                }
            }
            return Result.Succeeded;
        }
    }
}
