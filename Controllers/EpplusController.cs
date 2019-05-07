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


namespace EpplusPOC.Controllers
{
    public class EpplusController : Controller
    {
        //
        // GET: /Epplus/
        public ActionResult Index()
        {
            string fileName = "D:\\Practice\\PtacNewDealers.xlsx";
            //string conString = string.Empty;
            //conString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
            ExcelUtility excelutility = new ExcelUtility();

            List<EpplusModel> resulatantData = excelutility.loadData(fileName);
            String strConnString = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;

            SqlConnection con = new SqlConnection(strConnString);
                            
                                String TruncateQuery = "DELETE FROM PtacDealer_TB";
                                SqlCommand command = new SqlCommand(TruncateQuery, con);

                                con.Open();
                                command.ExecuteNonQuery();
                                con.Close();

            //SqlConnection con = new SqlConnection(strConnString);

           
            con.Open();
            
            try
            {              
               
                foreach (var item in resulatantData)
                {
                    SqlCommand cmd = new SqlCommand();

                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.CommandText = "spInsertPtacDealersInfo";

                    cmd.Parameters.Add("@DealerName", SqlDbType.VarChar).Value = item.DealerName;

                    cmd.Parameters.Add("@Address", SqlDbType.VarChar).Value = item.Address;

                    cmd.Parameters.Add("@City", SqlDbType.VarChar).Value = item.City;

                    cmd.Parameters.Add("@State", SqlDbType.VarChar).Value = item.State;

                    cmd.Parameters.Add("@ZipCode", SqlDbType.VarChar).Value = item.ZipCode;

                    cmd.Parameters.Add("@PhoneNumber", SqlDbType.VarChar).Value = item.PhoneNumber;

                    cmd.Parameters.Add("@PhoneNumber2", SqlDbType.VarChar).Value = item.PhoneNumber2;

                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    //con.Close();

                    //try
                    //{
                    //    con.Open();
                    //    cmd.ExecuteNonQuery();
                    //}
                    //finally
                    //{
                    //    con.Close();
                    //}
                }
            }
            catch (Exception)
            {
                
                throw;
            }

            finally
            {

                con.Close();                

            }
                return View(resulatantData);
            }
        }
    }


