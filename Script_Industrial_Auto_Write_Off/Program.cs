using System;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using CommonFunctions;
using System.Collections.Generic;
using System.Threading;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Script_Industrial_Auto_Write_Off
{
    class Program
    {

        //properties
        private static CF.ScreenScraper ss { get; set; }

        static void Main(string[] args)
        {
            List<FacilityRow> lstFacility = new List<FacilityRow>();
            int intScriptingErrors = 0;
            DateTime datScriptStart = DateTime.Now;

            try
            {
                //log start
                WriteToLog("START", "Script started.", null, null, null);

                //check for existing data
                if (DataExists())
                {
                    WriteToLog("ERROR", "Data already exists for today.", null, null, null);
                    Environment.Exit(0);
                }

                //get list
                lstFacility = GetFacilityList();

                //initialize
                InitializeEmulator();

                //loop thru facility list
                foreach (FacilityRow facilityrow in lstFacility)
                {

                    try
                    {

                        //process facility
                        ProcessFacility(facilityrow);

                    }
                    catch (Exception ex)
                    {

                        //increment scripting errors
                        intScriptingErrors++;

                        //log error
                        try
                        {
                            WriteToLog("ERROR", ex.ToString(), facilityrow.FacilityNumber, null, null);
                            WriteToLog("ERROR", ss.GetText(), null, null, null);
                        }
                        catch
                        {
                            //ignore error
                        }

                        //initialize
                        InitializeEmulator();

                    }

                }

                //report
                CreateReport(datScriptStart);

                //log finish
                WriteToLog("FINISH", "Script finished with " + intScriptingErrors + " errors", null, null, null);

            }
            catch (Exception ex)
            {
                WriteToLog("ERROR", ex.ToString(), null, null, null);
            }
            finally
            {
                //stop emulator
                if (ss != null) ss.Dispose();

            }

        }


        static void InitializeEmulator()
        {

            //stop emulator
            if (ss != null) ss.Dispose();

            //start emulator
            ss = new CF.ScreenScraper("SVCNSHSQ", "P3ngu1n35971fhS", 200);

            //logon to 60000
            ss.LogonFrom60000LogonScreen();

            //pause 10 secs
            Thread.Sleep(10000);

        }


        static void ProcessFacility(FacilityRow facilityrow)
        {
            List<AccountRow> lstAccount = new List<AccountRow>();
            List<string> lstNote = new List<string>();
            decimal decTotal;
            string strTotal;
            string strTemp = null;

            //get account list
            lstAccount = GetAccountList(facilityrow.FacilityNumber);

            //check for no accounts
            if (lstAccount.Count == 0)
            {
                return;
            }

            //check for more than 997 accounts - 
            if (lstAccount.Count > 997)
            {
                return;
            }

            //login to facility
            ss.WaitForString("===>");
            ss.LogonToFacilityFrom60000CommandPrompt(Convert.ToInt32(facilityrow.FacilityNumber));

            //daily processing menu
            ss.WaitForString("===>");
            ss.SendText("GO DLYPRS@E");

            //cash or adjustments
            ss.WaitForString("===>");
            ss.SendText("7@E");

            //enter summary data screen
            ss.WaitForString("AR0P05");
            ss.SendText("@E");

            //get totals
            decTotal = lstAccount.Where(x => x.Balance > 0).Select(x => x.Balance).Sum();
            strTotal = Convert.ToInt32((Math.Abs(decTotal) * 100)).ToString();

            //enter summary data 
            ss.WaitForString("TYPE BATCH CONTROL INFORMATION");
            ss.SendText(DateTime.Now.ToString("MMddyy@A@+"), 7, 38);
            ss.SendText(lstAccount.Count + "@A@+", 8, 38);
            ss.SendText(strTotal, 9, 38);
            ss.SendText("@A@-");
            ss.SendText("A@E", 10, 38);

            //loop thru accounts
            foreach (AccountRow accountrow in lstAccount.Where(x => x.Balance > 0)) 
            {
                //account
                ss.WaitForString("CASH RECEIPTS ENTRY");
                ss.SendText(accountrow.Account + "@A@+", 8, 8);

                //cy no
                ss.SendText("99@A@+", 8, 17);

                //payor co
                ss.SendText("999@A@+", 8, 21);

                //amount
                strTemp = Convert.ToInt32((accountrow.Balance * 100)).ToString();
                ss.SendText(strTemp, 8, 26);
                ss.SendText("@A@-");

                //date
                ss.SendText(DateTime.Now.ToString("MMddyy") + "@A@+", 8, 40);

                //trans code
                ss.SendText(accountrow.Tcode + "@A@+", 8, 48);

                //comment
                ss.SendText("INDUST AUTO WRITEOFF@E", 8, 57);

                //cycle not found
                if (ss.GetText().ToUpper().Contains("CYCLE NUMBER NOT FOUND FOR PATIENT"))
                {
                    ss.SendText("99@A@+", 8, 17);
                    ss.SendText("@E");
                }

                //return to entry screen
                ss.WaitForString("DELETED ENTRIES APPEAR AS REVERSE IMAGES");
                ss.SendText("@E");

                //log it
                WriteToLog("INFO", "Adjustment.", facilityrow.FacilityNumber, accountrow.Account,
                    -1 * accountrow.Balance);

            }

            //batch screen
            ss.WaitForString("CASH RECEIPTS ENTRY");
            ss.SendText("@c");

            //end batch
            ss.WaitForString("TYPE BATCH CONTROL INFORMATION");
            ss.SendText("@2");

            //return to command prompt
            ss.SendText("@3");
            ss.SendText("@3");

            //post notes - loop thru accounts
            lstNote.Add("Weekly auto adjustment submitted to bring balance to zero.");
            foreach (AccountRow accountrow in lstAccount.Where(x => x.Balance > 0)) 
            {
                ss.PostNote(accountrow.Account, lstNote);
            }

            //logoff facility
            ss.WaitForString("===>");
            ss.SendText("SO@E");

        }


        static List<FacilityRow> GetFacilityList()
        {
            SqlConnection cnnCom = null;
            SqlDataAdapter daCom = null;
            DataTable dt = null;
            List<FacilityRow> lstFacility = new List<FacilityRow>();
            string strSql = null;

            try
            {
                //open connection
                cnnCom = CF.OpenSqlConnectionWithRetry(CF.GetConnectionString(CF.DatabaseName.COMMON), 10);

                //loop thru facilities                    
                strSql =
                    "SELECT " +
                    "FACNUMBER, " +
                    "FACCONSTR_VS " +
                    "FROM COM_FACILITIES " +
                    "WHERE ENABLED=1 " +
                    "AND CAST(FACNUMBER AS INT)=28 " +
                    "ORDER BY FACNAME";
                daCom = new SqlDataAdapter(strSql, cnnCom);
                daCom.Fill(dt = new System.Data.DataTable());
                foreach (DataRow dr in dt.Rows)
                {
                    lstFacility.Add(new FacilityRow
                    {
                        FacilityNumber = dr["FACNUMBER"].ToString(),
                        ConnectionString = dr["FACCONSTR_VS"].ToString()
                    });
                }

                //return list
                return lstFacility;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (cnnCom != null) cnnCom.Dispose();
                if (daCom != null) daCom.Dispose();
                if (dt != null) dt.Dispose();
            }

        }


        static List<AccountRow> GetAccountList(string FacilityNumber)
        {
            SqlConnection cnnCom = null;
            SqlDataAdapter daCom = null;
            OdbcConnection cnnHms = null;
            OdbcDataAdapter daHms = null;
            DataTable dtMrn = null;
            DataTable dtPlan = null;
            DataTable dt = null;
            string strSql = null;
            string strFacilityNumber = null;
            string strConnectionString = null;
            List<AccountRow> lstAccount = new List<AccountRow>();
            

            try
            {
                //open connection
                cnnCom = CF.OpenSqlConnectionWithRetry(CF.GetConnectionString(CF.DatabaseName.COMMON), 10);

                //loop thru mrns             
                strSql =
                    "SELECT " +
                    "MRN, " +
                    "TCODE " +
                    "FROM SCRIPT_INDUSTRIAL_AUTO_WRITE_OFF_MRN " +
                    "WHERE FACILITYNUMBER=" + FacilityNumber + " " +
                    "ORDER BY MRN";
                daCom = new SqlDataAdapter(strSql, cnnCom);
                daCom.Fill(dtMrn = new DataTable());
                foreach (DataRow drMrn in dtMrn.Rows)
                {

                    //get connection string
                    strSql =
                        "SELECT " +
                        "FACNUMBER, " +
                        "FACCONSTR_VS " +
                        "FROM COM_FACILITIES " +
                        "WHERE CAST(FACNUMBER AS INT)='" + FacilityNumber + "'";
                    daCom = new SqlDataAdapter(strSql, cnnCom);
                    daCom.Fill(dt = new DataTable());
                    strFacilityNumber = dt.Rows[0]["FACNUMBER"].ToString();
                    strConnectionString = dt.Rows[0]["FACCONSTR_VS"].ToString();

                    //connect to HMS
                    cnnHms = CF.OpenOdbcConnectionWithRetry(strConnectionString, 10);

                    //get accounts for that mrn
                    strSql =
                        "SELECT " +

                        "PATNO, " +

                        "CASE " +
                        "WHEN (SELECT ISBDWDT FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)<>'0001-01-01' " +
                        "THEN (SELECT CURBD FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT ISDTFBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)<>'0001-01-01' " +
                        "THEN (SELECT CURBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT COUNT(PATNO) FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)>0 " +
                        "THEN (SELECT CURBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT COUNT(PATNO) FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO)>0 " +
                        "THEN (SELECT BALAN FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO) " +
                        "ELSE NULL " +
                        "END AS BALANCE " +

                        "FROM " +

                        "(SELECT PATNO FROM HOSPF" + strFacilityNumber + ".PATIENTS " +
                        "WHERE HSTNUM='" + drMrn["MRN"].ToString() + "' " +
                        "UNION " +
                        "SELECT PATNO FROM HOSPF" + strFacilityNumber + ".ARMAST " +
                        "WHERE HSTNUM='" + drMrn["MRN"].ToString() + "') AS T1 " +

                        "WHERE " +

                        "CASE " +
                        "WHEN (SELECT ISBDWDT FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)<>'0001-01-01' " +
                        "THEN (SELECT CURBD FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT ISDTFBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)<>'0001-01-01' " +
                        "THEN (SELECT CURBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT COUNT(PATNO) FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)>0 " +
                        "THEN (SELECT CURBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT COUNT(PATNO) FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO)>0 " +
                        "THEN (SELECT BALAN FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO) " +
                        "ELSE NULL " +
                        "END<>0 AND " +

                        "CASE " +
                        "WHEN (SELECT NWARFC1 FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) IS NOT NULL " +
                        "THEN (SELECT NWARFC1 FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT NWFINCL FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO) IS NOT NULL " +
                        "THEN (SELECT NWFINCL FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO) " +
                        "ELSE NULL " +
                        "END='I' " +

                        "ORDER BY T1.PATNO";
                    daHms = new OdbcDataAdapter(strSql, cnnHms);
                    daHms.SelectCommand.CommandTimeout = 0;
                    daHms.Fill(dt = new DataTable());

                    //add to list
                    foreach (DataRow dr in dt.Rows)
                    {
                        lstAccount.Add(new AccountRow
                        {
                            Account = dr["PATNO"].ToString(),
                            Balance = Convert.ToDecimal(dr["BALANCE"]),
                            Tcode = drMrn["TCODE"].ToString()
                        });
                    }

                }

                //loop thru plans          
                strSql =
                    "SELECT " +
                    "FACILITYNUMBER, " +
                    "INSURANCECOMPANY, " +
                    "INSURANCEPLAN, " +
                    "TCODE " +
                    "FROM SCRIPT_INDUSTRIAL_AUTO_WRITE_OFF_PLAN " +
                    "WHERE FACILITYNUMBER=" + FacilityNumber + " " +
                    "ORDER BY INSURANCECOMPANY,INSURANCEPLAN";
                daCom = new SqlDataAdapter(strSql, cnnCom);
                daCom.Fill(dtPlan = new DataTable());
                foreach (DataRow drPlan in dtPlan.Rows)
                {
                  
                    //get connection string
                    strSql =
                        "SELECT " +
                        "FACNUMBER, " +
                        "FACCONSTR_VS " +
                        "FROM COM_FACILITIES " +
                        "WHERE CAST(FACNUMBER AS INT)='" + drPlan["FACILITYNUMBER"].ToString() + "'";
                    daCom = new SqlDataAdapter(strSql, cnnCom);
                    daCom.Fill(dt = new DataTable());
                    strFacilityNumber = dt.Rows[0]["FACNUMBER"].ToString();
                    strConnectionString = dt.Rows[0]["FACCONSTR_VS"].ToString();

                    //connect to HMS
                    cnnHms = CF.OpenOdbcConnectionWithRetry(strConnectionString, 10);

                    //get accounts for that plan
                    strSql =
                        "SELECT " +

                        "PATNO, " +

                        "CASE " +
                        "WHEN (SELECT ISBDWDT FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)<>'0001-01-01' " +
                        "THEN (SELECT CURBD FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT ISDTFBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)<>'0001-01-01' " +
                        "THEN (SELECT CURBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT COUNT(PATNO) FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)>0 " +
                        "THEN (SELECT CURBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT COUNT(PATNO) FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO)>0 " +
                        "THEN (SELECT BALAN FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO) " +
                        "ELSE NULL " +
                        "END AS BALANCE " +

                        "FROM " +

                        "( " +
                        "SELECT PATNO " +
                        "FROM HOSPF" + strFacilityNumber + ".PATIENTS " +
                        "WHERE " +
                        "(AINS1=" + drPlan["INSURANCECOMPANY"] + " AND " +
                        "APLN1=" + drPlan["INSURANCEPLAN"] + ") OR " +
                        "(AINS2=" + drPlan["INSURANCECOMPANY"] + " AND " +
                        "APLN2=" + drPlan["INSURANCEPLAN"] + ") OR " +
                        "(AINS3=" + drPlan["INSURANCECOMPANY"] + " AND " +
                        "APLN3=" + drPlan["INSURANCEPLAN"] + ")" +
                        "UNION " +
                        "SELECT PATNO " +
                        "FROM HOSPF" + strFacilityNumber + ".ARMAST " +
                        "WHERE INS1=" + drPlan["INSURANCECOMPANY"] + " AND IPL1=" + drPlan["INSURANCEPLAN"] + " " + 
                        ") AS T1 " +

                        "WHERE " +

                        "CASE " +
                        "WHEN (SELECT ISBDWDT FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)<>'0001-01-01' " +
                        "THEN (SELECT CURBD FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT ISDTFBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)<>'0001-01-01' " +
                        "THEN (SELECT CURBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT COUNT(PATNO) FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO)>0 " +
                        "THEN (SELECT CURBL FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT COUNT(PATNO) FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO)>0 " +
                        "THEN (SELECT BALAN FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO) " +
                        "ELSE NULL " +
                        "END<>0 AND " +

                        "CASE " +
                        "WHEN (SELECT NWARFC1 FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) IS NOT NULL " +
                        "THEN (SELECT NWARFC1 FROM HOSPF" + strFacilityNumber + ".ARMAST WHERE PATNO=T1.PATNO) " +
                        "WHEN (SELECT NWFINCL FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO) IS NOT NULL " +
                        "THEN (SELECT NWFINCL FROM HOSPF" + strFacilityNumber + ".PATIENTS WHERE PATNO=T1.PATNO) " +
                        "ELSE NULL " +
                        "END='I' " +

                        "ORDER BY T1.PATNO";
                    daHms = new OdbcDataAdapter(strSql, cnnHms);
                    daHms.SelectCommand.CommandTimeout = 0;
                    daHms.Fill(dt = new DataTable());

                    //add to list
                    foreach (DataRow dr in dt.Rows)
                    {
                        lstAccount.Add(new AccountRow
                        {
                            Account = dr["PATNO"].ToString(),
                            Balance = Convert.ToDecimal(dr["BALANCE"]),
                            Tcode = drPlan["TCODE"].ToString()
                        });
                    }

                    //dispose connection
                    if (cnnHms != null) cnnHms.Dispose();
                    if (daHms != null) daHms.Dispose();

                }

                //return list
                return lstAccount;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                //dispose
                if (cnnCom != null) cnnCom.Dispose();
                if (daCom != null) daCom.Dispose();
                if (cnnHms != null) cnnHms.Dispose();
                if (daHms != null) daHms.Dispose();
                if (dtMrn != null) dtMrn.Dispose();
                if (dtPlan != null) dtPlan.Dispose();
                if (dt != null) dt.Dispose();

            }


        }


        static bool DataExists()
        {
            SqlConnection cnnCom = null;
            SqlDataAdapter daCom = null;
            DataTable dt = null;
            string strSql = null;
            bool blnDataExists;

            //open cnn
            cnnCom = new SqlConnection(CF.GetConnectionString(CF.DatabaseName.COMMON));
            cnnCom.Open();

            //check data
            strSql =
                "SELECT * " +
                "FROM SCRIPT_INDUSTRIAL_AUTO_WRITE_OFF " +
                "WHERE CAST(TIMESTAMP AS DATE)='" + DateTime.Now.Date + "' AND LOGTYPE='INFO'";
            daCom = new SqlDataAdapter(strSql, cnnCom);
            daCom.SelectCommand.CommandTimeout = 0;
            daCom.Fill(dt = new DataTable());
            if (dt.Rows.Count == 0)
            {
                blnDataExists = false;
            }
            else
            {
                blnDataExists = true;
            }

            //dispose
            if (cnnCom != null) cnnCom.Dispose();
            if (daCom != null) daCom.Dispose();
            if (dt != null) dt.Dispose();

            //return
            return blnDataExists;

        }


        static void WriteToLog(string LogType, string Message, string FacilityNumber, string Account,
            decimal? Amount)
        {
            SqlConnection cnnSql = null;
            SqlCommand cmdSql = null;
            string strSql = null;
            SqlParameter sqp = null;

            //open cnn
            cnnSql = new SqlConnection(CF.GetConnectionString(CF.DatabaseName.COMMON));
            cnnSql.Open();

            //write to log
            strSql =
                "INSERT INTO Script_Industrial_Auto_Write_Off ( " +
                "LogType, " +
                "Message, " +
                "TimeStamp, " +
                "FacilityNumber, " +
                "Account, " +
                "Amount) " +
                "VALUES ( " +
                "@LogType, " +
                "@Message, " +
                "@TimeStamp, " +
                "@FacilityNumber, " +
                "@Account, " +
                "@Amount)";
            cmdSql = new SqlCommand(strSql, cnnSql);

            //logtype
            sqp = new SqlParameter("LogType", SqlDbType.NVarChar);
            sqp.Value = LogType.ToUpper();
            cmdSql.Parameters.Add(sqp);

            //message
            sqp = new SqlParameter("Message", SqlDbType.NVarChar);
            sqp.Value = Message;
            cmdSql.Parameters.Add(sqp);

            //timestamp
            sqp = new SqlParameter("TimeStamp", SqlDbType.DateTime2);
            sqp.Value = DateTime.Now;
            cmdSql.Parameters.Add(sqp);

            //facility number
            sqp = new SqlParameter("FacilityNumber", SqlDbType.NVarChar);
            if (FacilityNumber == null)
            {
                sqp.Value = DBNull.Value;
            }
            else
            {
                sqp.Value = FacilityNumber;
            }
            cmdSql.Parameters.Add(sqp);

            //account
            sqp = new SqlParameter("Account", SqlDbType.NVarChar);
            if (Account == null)
            {
                sqp.Value = DBNull.Value;
            }
            else
            {
                sqp.Value = Account;
            }
            cmdSql.Parameters.Add(sqp);

            //amount
            sqp = new SqlParameter("Amount", SqlDbType.Money);
            if (Amount == null)
            {
                sqp.Value = DBNull.Value;
            }
            else
            {
                sqp.Value = Amount;
            }
            cmdSql.Parameters.Add(sqp);

            //execute sql
            cmdSql.CommandTimeout = 0;
            cmdSql.ExecuteNonQuery();

            //dispose
            if (cnnSql != null) cnnSql.Dispose();
            if (cmdSql != null) cmdSql.Dispose();

        }


        static void CreateReport(DateTime ReportStart)
        {
            Excel.Application appExcel = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            SqlConnection cnnCom = null;
            SqlDataAdapter daCom = null;
            DataTable dt = null;
            string strSql = null;
            string strExcelReport = CF.Folder.WorkingArea + "\\Script_Industrial_Auto_Write_Off_" +
                DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + ".xlsx";

            try
            {
                //open connection
                cnnCom = new SqlConnection(CF.GetConnectionString(CF.DatabaseName.COMMON));
                cnnCom.Open();

                //open excel
                appExcel = new Excel.Application();
                Thread.Sleep(1000);
                wb = appExcel.Workbooks.Add();
                Thread.Sleep(1000);

                //get negative worksheet
                ws = wb.Worksheets.Add();
                ws.Name = "Negative Accounts";

                //copy to excel
                CF.ListToExcel<AccountRow>(lstAccount.Where(x => x.Balance < 0).ToList(), ws, true, 1, 1);

                //format
                ws.Columns.AutoFit();

                //get processed worksheet
                ws = wb.Worksheets.Add();
                ws.Name = "Processed Accounts";

                //get processed
                strSql =
                    "SELECT " +
                    "TIMESTAMP AS 'Time Stamp', " +
                    "FACILITYNUMBER AS 'Facility Number', " +
                    "ACCOUNT AS 'Account', " +
                    "AMOUNT AS 'Adjustment' " +
                    "FROM SCRIPT_INDUSTRIAL_AUTO_WRITE_OFF " +
                    "WHERE " +
                    "TIMESTAMP>'" + ReportStart.ToString() + "' AND " +
                    "LOGTYPE='INFO' ";
                daCom = new SqlDataAdapter(strSql, cnnCom);
                daCom.Fill(dt = new System.Data.DataTable());

                //copy to excel
                CF.DataTableToExcel(dt, ws, true, 1, 1);

                //format
                ws.Columns.AutoFit();

              
                //save excel
                wb.SaveAs(strExcelReport);

            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                //dispose 
                appExcel.Quit();
                CF.KillApp(appExcel.Hwnd);
                if (cnnCom != null) cnnCom.Dispose();
                if (daCom != null) daCom.Dispose();
                if (dt != null) dt.Dispose();
            }

        }

    }

    public class FacilityRow
    {
        public string FacilityNumber { get; set; }

        public string ConnectionString { get; set; }

    }


    class AccountRow
    {
        public string Account { get; set; }
        public decimal Balance { get; set; }
        public string Tcode { get; set; }
    }

}

