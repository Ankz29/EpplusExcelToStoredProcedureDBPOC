using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace EpplusPOC.Models
{
    public class EpplusModel
    {
        //public string _Description;
        public string DealerName
        {
            get;
            set;
            //get
            //{
            //    if(!String.IsNullOrEmpty(_Description))
            //        return ConvertToPascalCase(_Description);
            //    return null;
            //}
            //set
            //{
            //    _Description = value;
            //}
        }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }

        //public string _phoneNumber;
        public string PhoneNumber
        {
            get;
            set;
            //get { return GetPhoneNumber(_phoneNumber); }
            //set { _phoneNumber = PhoneNumber; }
        }

        //public string _phoneNumber2;
        public string PhoneNumber2
        {
            get;
            set;
            //get { return GetPhoneNumber(PhoneNumber2); }
            //set{ _phoneNumber2 = PhoneNumber2; }
        }

        //method to convert DealerName to Pascal Case & format DealerName as per requirement//
        public string ConvertToPascalCase(string DealerName)
        {
            string dealerNameLowerCase;
            string[] exceptionalCases = new string[] { "PTAC", "SVC", "SVCS", "LLC", "INC", "A/C", "ACR", "CO" };

            // Make DealerName string all lowercase, because ToTitleCase does not change all uppercase correctly //
            dealerNameLowerCase = DealerName.ToLower();

            // Creates a TextInfo based on the "en-US" culture//           
            TextInfo myTextInfo = new CultureInfo("en-US", false).TextInfo;
            dealerNameLowerCase = myTextInfo.ToTitleCase(dealerNameLowerCase);

            var data = GetFormattedDealerName(exceptionalCases, dealerNameLowerCase);

            return data;
        }

        public string GetFormattedDealerName(string[] exceptionalCases, string dealerNameLowerCase)
        {
            string formattedString = dealerNameLowerCase;
            var dealerNames = dealerNameLowerCase.Split(' ');
            foreach (var str in dealerNames)
            {
                bool exists = exceptionalCases.Any(s => s.ToLower().Equals(str.ToLower()));
                if (exists)
                {
                    var tempStr = str.ToUpper();
                    formattedString = formattedString.Replace(str, tempStr);
                }
            }
            return formattedString;
        }

        //method to format Phonenumber to required format//
        public string GetPhoneNumber(string number)
        {
            string formattedNumber;
            string result = Regex.Replace(number, "[^0-9a-zA-Z]+", "");
             if (string.IsNullOrEmpty(result))
            {
                return number = "";
            }
             else if (result != "NULL" && result.Length >= 10 )
            {

                formattedNumber = "(" + result.Substring(0, 3) + ")" + " " + result.Substring(3, 3) + " " + "-" + " " + result.Substring(6, 4);
            }

            
            else
            {
                return number = "";
            }

            return formattedNumber;
        }

        
        //public List<EpplusModel> loadData(string excelPath)
        //{
        //   var modelData = new List<EpplusModel>();
        //   FileInfo filePath = new FileInfo(excelPath);
        //    {
        //        using (ExcelPackage package = new ExcelPackage(filePath))
        //        {
        //            ExcelWorksheet workSheet = package.Workbook.Worksheets["Sheet1"]; //excel sheet name//
        //            int totalRows = workSheet.Dimension.Rows;
                   
        //            for (int i = 2; i <= totalRows; i++)
        //            {
        //                modelData.Add(new EpplusModel
        //                {
        //                    //addressid = !string.IsNullOrEmpty(i.addressid) ? i.addressid : ""
        //                    Description = ConvertToPascalCase(workSheet.Cells[i, 2].Value.ToString().Trim()),
        //                    Address = workSheet.Cells[i, 3].Value.ToString().Trim(),
        //                    City = workSheet.Cells[i, 4].Value.ToString().Trim(),
        //                    State = workSheet.Cells[i, 5].Value.ToString().Trim(),
        //                    ZipCode = workSheet.Cells[i, 6].Value.ToString().Trim(),
        //                    PhoneNumber = GetPhoneNumber(workSheet.Cells[i, 7].Value.ToString().Trim()), //call the variable in place of calling method, make use of getter & setter methods to server the functionality//
        //                    PhoneNumber2 =GetPhoneNumber( workSheet.Cells[i, 8].Value.ToString().Trim())
        //                });
        //            }


        //        }

        //    }
        //    return modelData;
        //}
    }
}