using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using EpplusPOC.Models;
using System.Data;
using EpplusPOC.Controllers;
using EpplusPOC.Models;

namespace EpplusPOC
{
    public static class SqlUtility
    {
        public static void exeuteStorProc(string storedProcedureName, List<KeyValuePair<string, string>> paramsData)
        {
            try
            {
                //resultantData.ToXML();
                String strConnString = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;

                using (SqlConnection connection = new SqlConnection(strConnString))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand();
                    
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = storedProcedureName;
                    foreach(var item in paramsData)
                    {
                        cmd.Parameters.Add(item.Key, SqlDbType.Xml).Value = item.Value;
                    }
                  
                    cmd.Connection = connection;
                    cmd.ExecuteNonQuery();
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}