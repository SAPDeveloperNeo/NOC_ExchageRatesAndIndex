
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Net;
using System.Data.Common;
using System.Web;
using System.Net.Mail;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Drawing.Drawing2D;
using SAPbobsCOM;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using Sap.Data.Hana;
using RestSharp;
using System.Net.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExchageRatesAndIndex
{
    public class Currency
    {
        public string iso3 { get; set; }
        public string name { get; set; }
        public int unit { get; set; }
    }
    public class Rate
    {
        public Currency currency { get; set; }
        public string buy { get; set; }
        public string sell { get; set; }
    }

    public class ExchangeRateData
    {
        public string date { get; set; }
        public List<Rate> rates { get; set; }
    }
    public class Root
    {
        public ExchangeRateData data { get; set; }
    }

    class clsMain
    {
        #region Variables

        public static SAPbobsCOM.Company oCompany = null;
        public static string CompanyDB, ServerName, UserName, Password, API, sPath, DocEntry = null, DocNum = "", Series = "";

        string[] supportedFormats = new string[] { "dd/MM/yy", "dd/MM/yyyy", "dd-MM-yy", "dd-MM-yyyy", "MM/dd/yyyy", "MM/dd/yy", "ddMMMyyyy", "ddMMyyyy", "yyyyMMdd" };


        private static SAPbobsCOM.Recordset oRec, oRecConfig, oRec1, oRecdata, oRecordset, oRecM;
        private static SAPbobsCOM.SBObob oSBObob;
        private System.Data.DataTable oRecordSet;


        //public static string DBName, ServerName, UserName;
        public static string ProjectName = "Exchange Rates And Index";
        public static string ProjectCode = "ERI";
        private int IntCode;
        private SAPbobsCOM.UserTablesMD oUserTablesMD = null;
        private SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
        private SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
        public static bool CloseFlg, BtnFlag = false, DBSetupFlg = false;
        public static bool Success = false;

        //System.Data.DataTable SynHeaderDT = new System.Data.DataTable();
        //System.Data.DataTable SynRowDT = new System.Data.DataTable();
        //System.Data.DataTable SynUpdateDT = new System.Data.DataTable();

        HanaConnection objConnection = new HanaConnection();
        HanaCommand objCommand = new HanaCommand();
        HanaDataAdapter objAdpter = new HanaDataAdapter();

        DataTable objDataTable = null;



        #endregion

        #region Main
        public clsMain()
        {
            try
            {
                SyncTransactionsInSAP();

                //System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {

            }

        }
        #endregion

        #region Connection

        public void Setconnection()
        {
            try
            {
                string sCookie;
                //string sConnectionContext;

                // First initialize the Company object
                oCompany = new SAPbobsCOM.Company();
                if (oCompany.Connected == true)
                    return;

                sCookie = oCompany.GetContextCookie();
                //sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
                //oCompany.SetSboLoginContext(sConnectionContext);
                oCompany.Connect();
            }
            catch (Exception ex)
            {
                //oCompany = SBO_Application.Company.GetDICompany();
                //SBO_Application.MessageBox(ex.ToString().ToString(), 1, "ok", "", "");s
            }
        }

        private void SetApplication()
        {
            //SAPbouiCOM.SboGuiApi SboGuiApi = null;
            //string sConnectionString = null;
            //SboGuiApi = new SAPbouiCOM.SboGuiApi();
            //sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            //sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
            try
            {
                // If there's no active application the connection will fail
                //SboGuiApi.Connect(sConnectionString);
            }
            catch
            { //  Connection failed
                System.Windows.Forms.MessageBox.Show("No SAP Business One Application");
                System.Environment.Exit(0);
            }
            // get an initialized application object
            //SBO_Application = SboGuiApi.GetApplication(-1);
        }

        #endregion

        #region Other Methods       

        public void SyncTransactionsInSAP()
        {
            try
            {

                CompanyDB = System.Configuration.ConfigurationManager.AppSettings["CompanyDB"].ToString();
                oRecordSet = null;
                oRecordSet = ByQueryReturnDataTable("Select * from Schema.\"@ONRB\" t0 where t0.\"U_Active\" = 'Y'");
                if (oRecordSet.Rows.Count > 0)
                {
                    for (int i = 0; i < oRecordSet.Rows.Count; i++)
                    {
                        //API = oRecordSet.Rows[i]["U_NRB_API"].ToString();
                        CompanyDB = oRecordSet.Rows[i]["Name"].ToString();
                        UserName = oRecordSet.Rows[i]["U_UserName"].ToString();
                        Password = oRecordSet.Rows[i]["U_Password"].ToString();
                        ServerName = oRecordSet.Rows[i]["U_ServerName"].ToString();
                        ConnectSAPCompany();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public DataTable ByQueryReturnDataTable(string SQLQuery)
        {
            try
            {
                BeginTranscation();
                objCommand.CommandType = CommandType.Text;
                objCommand.CommandText = SQLQuery.Replace("Schema", CompanyDB);
                objAdpter.SelectCommand = objCommand;
                objDataTable = new DataTable();
                objAdpter.Fill(objDataTable);
                this.EndTranscation();
                return objDataTable;
            }
            catch (Exception Ex)
            {
                this.EndTranscation();
                return null;
            }
        }

        public void BeginTranscation()
        {
            try
            {
                //SchemaName = oCompany.CompanyDB;
                objConnection = new HanaConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ODBCConnection"].ToString());
                objCommand = new HanaCommand();
                objCommand.CommandTimeout = 1800;
                objCommand.Connection = objConnection;
                objCommand.Connection.Open();
            }
            catch (Exception ex)
            {

            }
        }

        public void EndTranscation()
        {
            try
            {
                if (objCommand.Connection.State == System.Data.ConnectionState.Open)
                {
                    objCommand.Connection.Close();
                    objCommand.Connection.Dispose();
                    objCommand.Dispose();
                }
            }
            catch (Exception Ex)
            {
            }
        }

        public static string ConnectSAPCompany()
        {
            string retMsg = String.Empty;
            int retValue = 0;
            DateTime date = DateTime.Now; // Replace with your date value

            // Format the date as "yyyy-MM-dd"
            string formattedDate = date.ToString("yyyy-MM-dd");

            try
            {
                oCompany = new SAPbobsCOM.Company();
                oCompany.Server = ServerName;
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                oCompany.CompanyDB = CompanyDB;
                oCompany.UserName = UserName;
                oCompany.Password = Password;
                retValue = oCompany.Connect();

                if (retValue != 0)
                {
                    oCompany.GetLastError(out retValue, out retMsg);
                    Log.WriteEventLog("SAP Connect: " + retMsg);
                    //return retValue + ": " + retMsg;
                }
                else
                {
                    Log.WriteEventLog("Connected with DBName = " + CompanyDB + " ");

                    var options = new RestClientOptions("https://www.nrb.org.np")
                    {
                        MaxTimeout = -1,
                    };
                    var client = new RestClient(options);
                    var request = new RestRequest("/api/forex/v1/rates?from=" + formattedDate + "&to=" + formattedDate + "&per_page=1&page=1", Method.Get);

                    RestResponse response = client.Execute(request);
                    string psReturn = response.Content.ToString();

                    JObject jsonObject = JObject.Parse(psReturn);

                    // Access the rates array and display rates
                    JArray rates = (JArray)jsonObject["data"]["payload"][0]["rates"];
                    foreach (JObject rate in rates)
                    {
                        string currencyName = (string)rate["currency"]["name"];
                        string iso3Code = (string)rate["currency"]["iso3"];
                        string buyRate = (string)rate["buy"];
                        string sellRate = (string)rate["sell"];

                        oSBObob = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                        if (iso3Code == "USD")
                        {
                            //oSBObob.SetCurrencyRate(iso3Code, DateTime.Now, Convert.ToDouble(buyRate), true);
                            //oSBObob.SetCurrencyRate("USS", DateTime.Now, Convert.ToDouble(sellRate), true);
                            //Log.WriteEventLog("Exchange Rate Updated USD = " + Convert.ToDouble(buyRate).ToString() + ", USS = " + Convert.ToDouble(sellRate).ToString() + " for a date : " + formattedDate + "");

                            oSBObob.SetCurrencyRate(iso3Code, DateTime.Now, Convert.ToDouble(sellRate), true);                          
                            Log.WriteEventLog("Exchange Rate Updated USD = " + Convert.ToDouble(sellRate).ToString() + ", USS = " + Convert.ToDouble(sellRate).ToString() + " for a date : " + formattedDate + "");


                        }
                        if (iso3Code == "INR")
                        {
                            oSBObob.SetCurrencyRate(iso3Code, DateTime.Now, Convert.ToDouble(sellRate) / 100, true);
                            Log.WriteEventLog("Exchange Rate Updated INR = " + (Convert.ToDouble(sellRate) / 100).ToString() + " for a date : " + formattedDate + "");
                        }
                        if (iso3Code == "EUR")
                        {
                            oSBObob.SetCurrencyRate(iso3Code, DateTime.Now, Convert.ToDouble(sellRate) / 100, true);
                            Log.WriteEventLog("Exchange Rate Updated " + iso3Code + " = " + (Convert.ToDouble(sellRate) / 100).ToString() + " for a date : " + formattedDate + "");
                        }
                        //Console.WriteLine($"Currency: {currencyName}");
                        //Console.WriteLine($"ISO3 Code: {iso3Code}");
                        //Console.WriteLine($"Buy Rate: {buyRate}");
                        //Console.WriteLine($"Sell Rate: {sellRate}");
                        //Console.WriteLine();
                    }
                    //return "Connected";

                }

                return "";
            }
            catch (Exception ex)
            {
                Log.WriteEventLog("SAP Connect" + ex.Message);
                return ex.Message;
            }
        }

        #endregion
    }
}
