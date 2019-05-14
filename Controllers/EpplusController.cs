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


namespace EpplusPOC.Controllers
{
    public class EpplusController : Controller
    {
        //
        // GET: /Epplus/
        public  ActionResult Index()
        {
           
            string fileName = "D:\\Practice\\PtacNewDealers.xlsx";
            ExcelUtility excelutility = new ExcelUtility();

            List<EpplusModel> resulatantData = excelutility.loadData(fileName);
            String strConnString = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;

            //command to truncate the previous data from Database//
            SqlConnection con = new SqlConnection(strConnString);
            //String TruncateQuery = "DELETE FROM PtacDealer_TB";
            //SqlCommand command = new SqlCommand(TruncateQuery, con);

            con.Open();
            //command.ExecuteNonQuery();
            //con.Close();
            //con.Open();

            try
            {
                //method to serialize List of data to XML//
                var model = new EpplusModel();
                var xmlData = model.serializeXml(resulatantData);                   
                   SqlCommand cmd = new SqlCommand();
                   cmd.CommandType = CommandType.StoredProcedure;
                   cmd.CommandText = "spInsertPtacDealersXmlInfo";
                   cmd.Parameters.Add("@dealersXml", SqlDbType.Xml).Value = xmlData;
                   cmd.Connection = con;
                   cmd.ExecuteNonQuery();
            }

            catch (Exception ex)
            {
                throw ex;
            }

            finally
            {
                con.Close();
            }
            return View(resulatantData);
        }


       
    }
}


