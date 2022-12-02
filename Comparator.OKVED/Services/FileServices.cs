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
        private static int GetNumberColumn(ExcelWorksheet worksheet, string nameColumn)
        {

            for (int i = 1; i < worksheet.Dimension.End.Column; i++)
            {
                if (worksheet.Cells[1, i].Value.ToString().Trim() == nameColumn)
                {
                    return i;
                }
            }

            return default(int);
        }

        public static IEnumerable<ModelPBD> ReadExcel(string pathFile)
        {
            ExcelPackage package = new ExcelPackage(pathFile);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            List<ModelPBD> listDataPBD = new List<ModelPBD>();

            int numColOKATO = GetNumberColumn(worksheet, "ОКАТО");
            int numColOKPO = GetNumberColumn(worksheet, "ОКПО");
            int numColName = GetNumberColumn(worksheet, "Наименование предприятия");
            int numColOKVEDHoz = GetNumberColumn(worksheet, "ОКВЭД Хозяйственный");
            int numColOKVEDChist = GetNumberColumn(worksheet, "ОКВЭД Чистый");
            int numColOtchMes = GetNumberColumn(worksheet, "За отчётный месяц");
            int numColPerNachGod = GetNumberColumn(worksheet, "За период с начала отчетного года");
            int numColOtchKvart = GetNumberColumn(worksheet, "За отчетный квартал");

            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                ModelPBD newRow = new ModelPBD();
                newRow.OKATO = worksheet.Cells[i, numColOKATO].Value.ToString();
                newRow.OKPO = worksheet.Cells[i, numColOKPO].Value.ToString();
                newRow.Name = worksheet.Cells[i, numColName].Value.ToString();
                newRow.OKVEDHoz = worksheet.Cells[i, numColOKVEDHoz].Value.ToString();
                newRow.OKVEDChist = worksheet.Cells[i, numColOKVEDChist].Value.ToString();

                if (worksheet.Cells[i, numColOtchMes].Value != null)
                {
                    newRow.OtchMes= (double)worksheet.Cells[i, numColOtchMes].Value;
                }
                else
                {
                    newRow.PerNachOtchGod = null;
                }

                if (worksheet.Cells[i, numColPerNachGod].Value != null)
                {
                    newRow.PerNachOtchGod = (double)worksheet.Cells[i, numColPerNachGod].Value;
                }
                else
                {
                    newRow.PerNachOtchGod = null;
                }

                if (worksheet.Cells[i, numColOtchKvart].Value != null)
                {
                    newRow.OtchKvart = (double)worksheet.Cells[i, numColOtchKvart].Value;
                }
                else
                {
                    newRow.OtchKvart = null;
                }

                listDataPBD.Add(newRow);
            }

            return listDataPBD;
        }

        public static byte[] CreateExcelResultHozChist(IEnumerable<ModelResultHozChist> collectionResult)
        {

            var package = new ExcelPackage();
            var sheetResultCompareOKVED = package.Workbook.Worksheets.Add("Сравнение хозяйственного и чистого ОКВЭД");
           
            int rowsCountListCompareOKVED = collectionResult.Count() + 1;

            sheetResultCompareOKVED.Cells["A1"].Value = "ОКПО";
            sheetResultCompareOKVED.Cells["B1"].Value = "Наименование предприятия";
            sheetResultCompareOKVED.Cells["C1"].Value = "ОКАТО";
            sheetResultCompareOKVED.Cells["D1"].Value = "ОКВЭД хозяйственный";
            sheetResultCompareOKVED.Cells["E1"].Value = "ОКВЭД чистый";

            sheetResultCompareOKVED.Cells[2, 1].LoadFromCollection<ModelResultHozChist>(collectionResult);

            sheetResultCompareOKVED.View.FreezePanes(2, 1);
            sheetResultCompareOKVED.Cells[1, 1, 1, 5].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheetResultCompareOKVED.Cells[1, 1, 1, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 217, 217));
            sheetResultCompareOKVED.Columns[1].Width = 32;
            sheetResultCompareOKVED.Columns[2].Width = 32;
            sheetResultCompareOKVED.Columns[3].Width = 32;
            sheetResultCompareOKVED.Columns[4].Width = 32;
            sheetResultCompareOKVED.Columns[5].Width = 32;
            sheetResultCompareOKVED.Columns[1, 5].Style.Font.Name = "Arial";
            sheetResultCompareOKVED.Columns[1, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            sheetResultCompareOKVED.Columns[1, 5].Style.Font.Size = 10;
            sheetResultCompareOKVED.Cells[1, 1, 1, 5].Style.Font.Bold = true;
            sheetResultCompareOKVED.Cells[1, 1, 1, 5].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            sheetResultCompareOKVED.Rows[1].Height = 30;
            sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 5].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            return package.GetAsByteArray();
        }

        public static void SaveFile(byte[] dataSave, string pathSave)
        {
            File.WriteAllBytes(pathSave, dataSave);
        }
    }
}
