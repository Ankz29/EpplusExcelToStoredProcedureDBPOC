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
using System.Xml;
using System.Xml.Serialization;

namespace EpplusPOC.Models
{
    public class EpplusModel
    {     
        public string DealerName { get; set; }        
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public string PhoneNumber { get; set; }
        public string PhoneNumber2 { get; set; }
        
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
            else if (result != "NULL" && result.Length >= 10)
            {

                formattedNumber = "(" + result.Substring(0, 3) + ")" + " " + result.Substring(3, 3) + " " + "-" + " " + result.Substring(6, 4);
            }
            else
            {
                return number = "";
            }
            return formattedNumber;
        }

        public string serializeXml(List<EpplusModel> data)
        {
            //Represents an XML document,
            XmlDocument xmlDoc = new XmlDocument();
            // Initializes a new instance of the XmlDocument class.          
            XmlSerializer xmlSerializer = new XmlSerializer(data.GetType());
            // Creates a stream whose backing store is memory. 
            using (MemoryStream xmlStream = new MemoryStream())
            {
                xmlSerializer.Serialize(xmlStream, data);
                //gets or sets position of stream//
                xmlStream.Position = 0;
                //Loads the XML document from the specified string.
                xmlDoc.Load(xmlStream);
                return xmlDoc.InnerXml;
            }
        }
    }
}