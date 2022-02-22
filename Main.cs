using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPI4._2
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //string pipeInfo = string.Empty;

            var pipes = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            var saveDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "pipes.xlsx",
                DefaultExt = ".xlsx"
            };

            //string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "walls.xlsx");

            string selectedFilePath = string.Empty;
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveDialog.FileName;
            }

            if (string.IsNullOrEmpty(selectedFilePath))
                return Result.Cancelled;

            //File.WriteAllText(selectedFilePath, wallInfo);

            using (FileStream stream = new FileStream(selectedFilePath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист1");

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    Parameter pipeInnerDiam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                    Parameter pipeOuterDiam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    Parameter pipeLength = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    double innerDiameter = UnitUtils.ConvertFromInternalUnits(pipeInnerDiam.AsDouble(), UnitTypeId.Millimeters);
                    double outerDiameter = UnitUtils.ConvertFromInternalUnits(pipeOuterDiam.AsDouble(), UnitTypeId.Millimeters);
                    double length = UnitUtils.ConvertFromInternalUnits(pipeLength.AsDouble(), UnitTypeId.Millimeters);
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.Name);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, innerDiameter);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, outerDiameter);
                    sheet.SetCellValue(rowIndex, columnIndex: 3, length);
                    rowIndex++;
                }

                workbook.Write(stream);
                workbook.Close();
            }

            TaskDialog.Show("Selection", "Данные успешно записаны");
            System.Diagnostics.Process.Start(selectedFilePath);


            return Result.Succeeded;
        }
    }
}
