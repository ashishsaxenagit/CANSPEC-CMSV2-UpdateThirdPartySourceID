using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
//using Microsoft.Data.SqlClient;
using System.Data;
using System.Data.SqlClient;
namespace CANSPEC_CMSV2_UpdateThirdPartySourceID
{
    class Program
    {

        //https://www.csharp-console-examples.com/general/reading-excel-file-in-c-console-application/
        private static string CMSv2ConStr = System.Configuration.ConfigurationManager.AppSettings["CMSv2ConStr"];
        private static string DataFilePath = System.Configuration.ConfigurationManager.AppSettings["DataFilePath"];
        private static string strFlag = System.Configuration.ConfigurationManager.AppSettings["Flag"];
        
        static void Main(string[] args)
        {
            int rows = 0;
            int cols = 0;
            int i = 0;
            string SqlScr = "";
            string LatestJanssenFAID = "";
            string OldJanssenFAID = "";
            int NoofRows = 0;
            try
            {
                Application excelApp = new Application();
                if (excelApp == null)
                {
                    Console.WriteLine("Excel is not installed!!");
                    return;
                }

                Workbook excelBook = excelApp.Workbooks.Open(DataFilePath);
                _Worksheet excelSheet = excelBook.Sheets[1];
                Range excelRange = excelSheet.UsedRange;

                rows = excelRange.Rows.Count;
                cols = excelRange.Columns.Count;
                SqlScr = "";
                LatestJanssenFAID = "";
                OldJanssenFAID = "";
                Console.WriteLine("\r\n" + "Connecting DB:-" + CMSv2ConStr);
                Console.WriteLine("\r\n" + "Reading File :-" + DataFilePath);
                Console.WriteLine("\r\n\r\n" + "FAID Mapping Excel Import started!!");
                for (i = 2; i <= rows; i++)
                {
                    //create new line
                    Console.Write("\r\n");
                    //write the console
                    if (
                       (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null)
                       &&
                       (excelRange.Cells[i, 2] != null && excelRange.Cells[i, 2].Value2 != null)
                       )
                    {
                        LatestJanssenFAID = excelRange.Cells[i, 1].Value2;
                        OldJanssenFAID = excelRange.Cells[i, 2].Value2;
                    }

                    SqlScr = "INSERT INTO [DBO].[ProjectFREEWAY_3340_FinancialAssistance_JanssenFAIDMapping] ";
                    SqlScr = SqlScr + "(LatestJanssenFAID,OldJanssenFAID,DateAdded)";
                    SqlScr = SqlScr + " VALUES (@LatestJanssenFAID, @OldJanssenFAID,@DateAdded)";

                    using (SqlConnection cn = new SqlConnection(CMSv2ConStr))
                    using (SqlCommand command = new SqlCommand(SqlScr, cn))
                    {
                        command.Parameters.AddWithValue("@LatestJanssenFAID", LatestJanssenFAID);
                        command.Parameters.AddWithValue("@OldJanssenFAID", OldJanssenFAID);
                        command.Parameters.AddWithValue("@DateAdded", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        cn.Open();
                        int rowsAffected = command.ExecuteNonQuery();
                        cn.Close();
                        Console.Write("RowNo:" + i.ToString() + "- LatestSALESFORCEID-" + LatestJanssenFAID + "- LEGACYSALESFORCEID-" + OldJanssenFAID + "\t");
                    }
                }
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                Console.WriteLine("\r\n\r\n" + "FAID Mapping Excel successfully Imported!!");

                if (strFlag == "True")
                {
                    Console.WriteLine("\r\n" + "FAID Mapping Started!!");
                    using (SqlConnection cnSP = new SqlConnection(CMSv2ConStr))
                    using (SqlCommand command = new SqlCommand("[dbo].[sProjectFREEWAY_3340_ImportedJanssenFAIDMapping_OneTime]", cnSP))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        cnSP.Open();
                        NoofRows = command.ExecuteNonQuery();
                        cnSP.Close();
                        Console.WriteLine("\r\n" + "FAID Mapping Completed!!");
                    }
                    Console.WriteLine("\r\n" + "FAID Mapping Progress Completed for [ " + Convert.ToString(NoofRows) + " ] records !!");
                }
                else
                {
                    Console.WriteLine("\r\n" + "FAID Mapping escaped!!");
                }
                Console.ReadLine();
            }

            catch (Exception ex)
            {
                //Console.Write("RowNo:" + i.ToString() + "- LatestSALESFORCEID-" + LatestJanssenFAID + "- LEGACYSALESFORCEID-" + OldJanssenFAID + "\t");
                Console.Write("Error:-" + ex.Message);
                Console.ReadLine();
            }

        }


    }
}