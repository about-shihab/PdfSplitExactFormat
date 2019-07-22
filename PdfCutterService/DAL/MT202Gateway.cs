using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OracleClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfCutterService.DAL
{
    public class MT202Gateway
    {
        string connectionString = ConfigurationManager.ConnectionStrings["ULTIMUS"].ConnectionString;

        public void InsertSEBL_M202_MASTER_INFO(string plc_no, string pbill_ref, string pbill_date, int? pamount, string pemail_id, string pswfilelocation, string pvofilelocation)
        {

            OracleDataAdapter adp = new OracleDataAdapter();
            using (OracleConnection connection = new OracleConnection())
            {

                connection.ConnectionString = connectionString;
                try
                {

                    connection.Open();
                    OracleCommand command = new OracleCommand();
                    command.Connection = connection;
                    command.CommandText = "rsp_set_m202_master_info";
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.Add("plc_no", OracleType.NVarChar,50).Value = plc_no;
                    command.Parameters.Add("pbill_ref", OracleType.NVarChar,25).Value = pbill_ref;
                    command.Parameters.Add("pbill_date", OracleType.VarChar).Value = pbill_date;
                    command.Parameters.Add("pamount", OracleType.Number).Value = (pamount==null)?0: pamount;
                    command.Parameters.Add("pemail_id", OracleType.NVarChar,50).Value = pemail_id;
                    command.Parameters.Add("psent_status", OracleType.NVarChar,10).Value = "System";
                    command.Parameters.Add("psent_date", OracleType.VarChar).Value = DateTime.Now.ToString("dd/MM/yyyy");
                    command.Parameters.Add("premarks", OracleType.NVarChar).Value = "Sent Successfully";
                    command.Parameters.Add("pswfilelocation", OracleType.VarChar,500).Value = pswfilelocation==null?"": pswfilelocation;
                    command.Parameters.Add("pvofilelocation", OracleType.VarChar,500).Value = pvofilelocation;
                    command.Parameters.Add("perrorcode", OracleType.Int32, 5).Direction = ParameterDirection.Output;
                    command.Parameters.Add("perrormsg", OracleType.VarChar, 2000).Direction = ParameterDirection.Output;

                    command.ExecuteNonQuery();
                    string perrormsg = command.Parameters["perrormsg"].Value.ToString();
                    command.Parameters.Clear();
                    //return perrormsg;
                }

                catch (OracleException ex)
                {
                    throw ex;
                }
                finally
                {
                    connection.Close();
                }
            }

        }
    }
}
