using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EpplusPOC.Models;
using OfficeOpenXml;
using System.IO;

namespace EpplusPOC
{
    public class ExcelUtility
    {
        public List<EpplusModel> loadData(string excelPath)
        {
            var modelData = new List<EpplusModel>();
            FileInfo filePath = new FileInfo(excelPath);
            {
                using (ExcelPackage package = new ExcelPackage(filePath))
                {
                    ExcelWorksheet workSheet = package.Workbook.Worksheets["Sheet1"]; //excel sheet name//
                    int totalRows = workSheet.Dimension.Rows;
                    EpplusModel model = new EpplusModel();
                    for (int i = 2; i <= totalRows; i++)
                    {
                        modelData.Add(new EpplusModel
                        {

                            DealerName = (workSheet.Cells[i, 2].Value != null) ? model.ConvertToPascalCase(workSheet.Cells[i, 2].Value.ToString().Trim()) : "",
                            Address = (workSheet.Cells[i, 3].Value != null) ? workSheet.Cells[i, 3].Value.ToString().Trim() : "",
                            City = (workSheet.Cells[i, 4].Value != null) ? workSheet.Cells[i, 4].Value.ToString().Trim() : "",
                            State = (workSheet.Cells[i, 5].Value != null) ? workSheet.Cells[i, 5].Value.ToString().Trim() : "",
                            ZipCode = (workSheet.Cells[i, 6].Value != null) ? workSheet.Cells[i, 6].Value.ToString().Trim() : "",
                            PhoneNumber = (workSheet.Cells[i, 7].Value != null) ? model.GetPhoneNumber(workSheet.Cells[i, 7].Value.ToString().Trim()) : "",
                            PhoneNumber2 = (workSheet.Cells[i, 8].Value != null) ? model.GetPhoneNumber(workSheet.Cells[i, 8].Value.ToString().Trim()) : "",

                        });
                    }


                }

            }
            return modelData;
        }
    }
}