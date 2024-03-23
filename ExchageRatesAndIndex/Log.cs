using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace ExchageRatesAndIndex
{
    public static class Log
    {
        public static void WriteEventLog(string Message)
        {
            StreamWriter sw = null;
            try
            {
                string oPath = System.Windows.Forms.Application.StartupPath + @"\SysLogs";
                string Date = System.DateTime.Now.ToString("dd-MM-yyyy");
                if (Directory.Exists("SysLogs") == false)
                {
                    Directory.CreateDirectory(oPath);
                }
                //oPath += @"\" + CommonFunctions.gSAPDatabase + "--- " + CommonFunctions.gPortalDatabase + DateTime.Now.ToString("dd-MM-yyyy") + "-AlertErrorLog.log";
                //StreamWriter sw = new StreamWriter(oPath, true);

                oPath += @"\HRISLog_" + Date + ".txt";
                sw = new StreamWriter(oPath, true);

                sw.WriteLine(DateTime.Now.ToString() + ":" + Message);
                sw.Flush();
                sw.Close();
            }
            catch (Exception Ex)
            {
            }
        }
    }
}
