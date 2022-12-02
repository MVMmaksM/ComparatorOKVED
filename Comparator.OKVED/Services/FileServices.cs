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


        public static IEnumerable<ModelPBD> LoadExcelPBD(string pathFile)
        {
            int a = 0;

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
                //int numColSovMesPredGod = GetNumberColumn(worksheet, "");
                int numColOtchKvart = GetNumberColumn(worksheet, "За отчетный квартал");
                int numColPredKvart = GetNumberColumn(worksheet, "За предыдущий квартал");
                //int numColSovKvartPredGod = GetNumberColumn(worksheet, "");

               

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

                    listDataPBD.Add(newRow);

                    a += i;
                }

                return listDataPBD;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace + "итерация" + a);
                return null;
            }
        }

        public static byte[] CreateExcelResultHozChist(IEnumerable<ModelResultHozChist> collectionResult)
        {

            var package = new ExcelPackage();
            var sheetResultCompareOKVED = package.Workbook.Worksheets.Add("Сравнение хозяйственного и чистого ОКВЭД");

            int rowsCountListCompareOKVED = collectionResult.Count() + 1;

            sheetResultCompareOKVED.Cells["A1"].Value = "До 15 человек";
            sheetResultCompareOKVED.Cells["B1"].Value = "ОКПО";
            sheetResultCompareOKVED.Cells["C1"].Value = "Наименование предприятия";
            sheetResultCompareOKVED.Cells["D1"].Value = "ОКАТО";
            sheetResultCompareOKVED.Cells["E1"].Value = "ОКВЭД хозяйственный";
            sheetResultCompareOKVED.Cells["F1"].Value = "ОКВЭД чистый";

            sheetResultCompareOKVED.Cells[2, 1].LoadFromCollection<ModelResultHozChist>(collectionResult);

            sheetResultCompareOKVED.View.FreezePanes(2, 1);
            sheetResultCompareOKVED.Cells[1, 1, 1, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheetResultCompareOKVED.Cells[1, 1, 1, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 217, 217));
            sheetResultCompareOKVED.Columns[1].Width = 20;
            sheetResultCompareOKVED.Columns[2].Width = 32;
            sheetResultCompareOKVED.Columns[3].Width = 32;
            sheetResultCompareOKVED.Columns[4].Width = 32;
            sheetResultCompareOKVED.Columns[5].Width = 32;
            sheetResultCompareOKVED.Columns[6].Width = 32;
            sheetResultCompareOKVED.Columns[1, 6].Style.Font.Name = "Arial";
            sheetResultCompareOKVED.Columns[1, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            sheetResultCompareOKVED.Columns[1, 6].Style.Font.Size = 10;
            sheetResultCompareOKVED.Cells[1, 1, 1, 6].Style.Font.Bold = true;
            sheetResultCompareOKVED.Cells[1, 1, 1, 6].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            sheetResultCompareOKVED.Rows[1].Height = 30;
            sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 6].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 6].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            return package.GetAsByteArray();
        }

        public static void SaveFile(byte[] dataSave, string pathSave)
        {
            File.WriteAllBytes(pathSave, dataSave);
        }
    }
}
