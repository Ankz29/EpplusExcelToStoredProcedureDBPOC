using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using EpplusPOC.Models;
using OfficeOpenXml; //epplus extension//
using System.Configuration;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using EpplusPOC;
using System.Xml.Serialization;
using System.Xml;
using EpplusPOC;

namespace EpplusPOC.Controllers
{
    public class EpplusController : Controller
    {
        // GET: /Epplus/
        public ActionResult Index()
        {

            string fileName = "D:\\Practice\\PtacNewDealers.xlsx";

            //ExcelUtility excelutility = new ExcelUtility();

            List<EpplusModel> resultantData = ExcelUtility.retrieveExcelData(fileName);
    
            var paramsData = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("@dealersXml",resultantData.toXmlExtension())
                };
            var storedProcedureName = System.Configuration.ConfigurationManager.AppSettings["storedProcedureName"];

            SqlUtility.exeuteStorProc(storedProcedureName, paramsData);

            return View(resultantData);
        }


    }
}


