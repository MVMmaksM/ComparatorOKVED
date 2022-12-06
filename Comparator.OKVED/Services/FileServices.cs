using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Comparator.OKVED.Model;
using System.IO;
using System.Windows.Forms;

namespace Comparator.OKVED.Services
{
    class FileServices
    {
        private static double? ValidationNullDataColumn(object value)
        {
            if (value != null)
            {
                return (double)value;
            }
            else
            {
                return null;
            }
        }
        private static int GetNumberColumn(ExcelWorksheet worksheet, string nameColumn)
        {

            for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
            {
                if (worksheet.Cells[1, i].Value.ToString().Trim() == nameColumn)
                {
                    return i;
                }
            }

            return default(int);
        }
        private static void SetPropertyColumnExcel(ExcelWorksheet excelWorksheet, int countRows, int countColumn)
        {
            excelWorksheet.View.FreezePanes(2, 1);
            excelWorksheet.Cells[1, 1, 1, countColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            excelWorksheet.Cells[1, 1, 1, countColumn].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 217, 217));
            excelWorksheet.Columns[1, countColumn].Width = 30;
            excelWorksheet.Cells[1, 1, countRows, countColumn].Style.Font.Name = "Arial";
            excelWorksheet.Cells[1, 1, countRows, countColumn].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            excelWorksheet.Columns[1, countColumn].Style.Font.Size = 10;
            excelWorksheet.Cells[1, 1, 1, countColumn].Style.Font.Bold = true;
            excelWorksheet.Cells[1, 1, 1, countColumn].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            excelWorksheet.Rows[1].Height = 30;
            excelWorksheet.Cells[1, 1, countRows, countColumn].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            excelWorksheet.Cells[1, 1, countRows, countColumn].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            excelWorksheet.Cells[1, 1, countRows, countColumn].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            excelWorksheet.Cells[1, 1, countRows, countColumn].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }
        public static IEnumerable<ModelPBD> LoadExcelPBD(string pathFile)
        {      

            try
            {
                ExcelPackage package = new ExcelPackage(pathFile);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                List<ModelPBD> listDataPBD = new List<ModelPBD>();

                int numColPeriod = GetNumberColumn(worksheet, "До 15 человек");
                int numColOKATO = GetNumberColumn(worksheet, "ОКАТО");
                int numColOKPO = GetNumberColumn(worksheet, "ОКПО");
                int numColName = GetNumberColumn(worksheet, "Наименование предприятия");
                int numColKodPokaz = GetNumberColumn(worksheet, "Код показателя");
                int numColOKVEDHoz = GetNumberColumn(worksheet, "ОКВЭД Хозяйственный");
                int numColOKVEDChist = GetNumberColumn(worksheet, "ОКВЭД Чистый");
                int numColOtchMes = GetNumberColumn(worksheet, "За отчётный месяц");
                int numColPredMes = GetNumberColumn(worksheet, "За предыдущий месяц");
                int numColSovMesPredGod = GetNumberColumn(worksheet, "За соответствующий месяц прошлого года");
                int numColPerSnachOtchGod = GetNumberColumn(worksheet, "За период с начала отчетного года");
                int numColOtchKvart = GetNumberColumn(worksheet, "За отчетный квартал");
                int numColPredKvart = GetNumberColumn(worksheet, "За предыдущий квартал");
                int numColSovPerPredGod = GetNumberColumn(worksheet, "За соответствующий период предыдущего года");
                int numColSovKvartPredGod = GetNumberColumn(worksheet, "За соответствующий квартал предыдущего года");

                for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                {
                    ModelPBD newRow = new ModelPBD();
                    newRow.Period = worksheet.Cells[i, numColPeriod].Value.ToString();
                    newRow.OKATO = worksheet.Cells[i, numColOKATO].Value.ToString();
                    newRow.OKPO = worksheet.Cells[i, numColOKPO].Value.ToString();
                    newRow.Name = worksheet.Cells[i, numColName].Value.ToString();
                    newRow.KodPokaz = worksheet.Cells[i, numColKodPokaz].Value.ToString();
                    newRow.OKVEDHoz = worksheet.Cells[i, numColOKVEDHoz].Value.ToString();
                    newRow.OKVEDChist = worksheet.Cells[i, numColOKVEDChist].Value.ToString();

                    newRow.OtchMes = ValidationNullDataColumn(worksheet.Cells[i, numColOtchMes].Value);
                    newRow.PredMes = ValidationNullDataColumn(worksheet.Cells[i, numColPredMes].Value);
                    newRow.OtchKvart = ValidationNullDataColumn(worksheet.Cells[i, numColOtchKvart].Value);
                    newRow.PredKvart = ValidationNullDataColumn(worksheet.Cells[i, numColPredKvart].Value);
                    newRow.SovMesPredGod = ValidationNullDataColumn(worksheet.Cells[i, numColSovMesPredGod].Value);
                    newRow.PerSnachOtchGod = ValidationNullDataColumn(worksheet.Cells[i, numColPerSnachOtchGod].Value);
                    newRow.SovPerPredGod = ValidationNullDataColumn(worksheet.Cells[i, numColSovPerPredGod].Value);
                    newRow.SovKvartPredGod = ValidationNullDataColumn(worksheet.Cells[i, numColSovKvartPredGod].Value);

                    listDataPBD.Add(newRow);
                }

                return listDataPBD;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace + "итерация" );
                return null;
            }
        }
        public static byte[] CreateExcelResultCompare(IEnumerable<ModelResultHozChist> resultHozCompareChist, IEnumerable<ModelResultChist> resultOtchComparePredChist)
        {
            var package = new ExcelPackage();
            var sheetResultHozCompareChist = package.Workbook.Worksheets.Add("Сравнение хозяйственного и чистого ОКВЭД");
            var sheetResultOtchComparePredChist = package.Workbook.Worksheets.Add("Сравнение чистого ОКВЭД");

            int countRowsResHozCompChist = resultHozCompareChist.Count() + 1;
            int countRowsResOtchCompPredChist = resultOtchComparePredChist.Count() + 1;

            sheetResultHozCompareChist.Cells["A1"].Value = "До 15 человек";
            sheetResultHozCompareChist.Cells["B1"].Value = "ОКПО";
            sheetResultHozCompareChist.Cells["C1"].Value = "Наименование предприятия";
            sheetResultHozCompareChist.Cells["D1"].Value = "ОКАТО";
            sheetResultHozCompareChist.Cells["E1"].Value = "Код показателя";
            sheetResultHozCompareChist.Cells["F1"].Value = "ОКВЭД хозяйственный";
            sheetResultHozCompareChist.Cells["G1"].Value = "ОКВЭД чистый";

            sheetResultOtchComparePredChist.Cells["A1"].Value = "До 15 человек";
            sheetResultOtchComparePredChist.Cells["B1"].Value = "ОКПО";
            sheetResultOtchComparePredChist.Cells["C1"].Value = "Наименование предприятия";
            sheetResultOtchComparePredChist.Cells["D1"].Value = "ОКАТО";
            sheetResultOtchComparePredChist.Cells["E1"].Value = "Код показателя";
            sheetResultOtchComparePredChist.Cells["F1"].Value = "Чистый ОКВЭД текущего периода";
            sheetResultOtchComparePredChist.Cells["G1"].Value = "Чистый ОКВЭД предыдущего периода";
            //sheetResultOtchComparePredChist.Cells["H1"].Value = "Чистый ОКВЭД отчетного квартала";
            //sheetResultOtchComparePredChist.Cells["I1"].Value = "Чистый ОКВЭД предыдущего квартала";

            SetPropertyColumnExcel(sheetResultHozCompareChist, countRowsResHozCompChist, typeof(ModelResultHozChist).GetProperties().Length);
            SetPropertyColumnExcel(sheetResultOtchComparePredChist, countRowsResOtchCompPredChist, typeof(ModelResultChist).GetProperties().Length);

            sheetResultHozCompareChist.Cells[2, 1].LoadFromCollection<ModelResultHozChist>(resultHozCompareChist);
            sheetResultOtchComparePredChist.Cells[2, 1].LoadFromCollection<ModelResultChist>(resultOtchComparePredChist);


            return package.GetAsByteArray();
        }
        public static void SaveFile(byte[] dataSave, string pathSave)
        {
            File.WriteAllBytes(pathSave, dataSave);
        }
    }
}
