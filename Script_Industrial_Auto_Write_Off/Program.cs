using System;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using CommonFunctions;
using System.Collections.Generic;
using System.Threading;
using System.Linq;

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
            decimal decTotal;
            string strTotal;
            string strTemp = null;
            int intCount;

            //get account list
            lstAccount = GetAccountList(facilityrow);

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
            decTotal = lstAccount.Select(x => x.Amount).Sum();
            strTotal = Convert.ToInt32((Math.Abs(decTotal) * 100)).ToString();

            //enter summary data 
            ss.WaitForString("TYPE BATCH CONTROL INFORMATION");
            ss.SendText(DateTime.Now.ToString("MMddyy@A@+"), 7, 38);
            ss.SendText(lstAccount.Count + "@A@+", 8, 38);
            ss.SendText(strTotal, 9, 38);
            if (decTotal > 0)
            {
                ss.SendText("@A@+");
            }
            else
            {
                ss.SendText("@A@-");
            }
            ss.SendText("A@E", 10, 38);

            //loop thru accounts
            foreach (AccountRow accountrow in lstAccount)
            {
                //account
                ss.WaitForString("CASH RECEIPTS ENTRY");
                ss.SendText(accountrow.Account + "@A@+", 8, 8);

                //cy no
                ss.SendText("99@A@+", 8, 17);

                //payor co
                ss.SendText("999@A@+", 8, 21);

                //amount
                strTemp = Convert.ToInt32((Math.Abs(accountrow.Amount) * 100)).ToString();
                ss.SendText(strTemp, 8, 26);
                if (accountrow.Amount > 0)
                {
                    ss.SendText("@A@+");
                }
                else
                {
                    ss.SendText("@A@-");
                }

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
                //WriteToLog("INFO", accountrow.Comment, facilityrow.FacilityNumber, accountrow.Account,
                //   accountrow.Amount, accountrow.Ptype);

            }

            //batch screen
            ss.WaitForString("CASH RECEIPTS ENTRY");
            ss.SendText("@c");

            //end batch
            ss.WaitForString("TYPE BATCH CONTROL INFORMATION");
            ss.SendText("@2");

            //log it
            WriteToLog("INFO", "Batch complete.", facilityrow.FacilityNumber, null, null);

            //return
            ss.SendText("@3");
            ss.SendText("@3");

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


        static List<AccountRow> GetAccountList(FacilityRow facilityrow)
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
                    "WHERE FACILITYNUMBER=" + facilityrow.FacilityNumber + " " +
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
                        "WHERE CAST(FACNUMBER AS INT)='" + facilityrow.FacilityNumber + "'";
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
                        "END<>0 " +

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
                            Amount = Convert.ToDecimal(dr["BALANCE"]),
                            Tcode = drMrn["TCODE"].ToString()
                        });
                    }

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
                "FROM SCRIPT_INDUSTRIAL_ACCOUNTS " +
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


        public static void WriteToLog(string LogType, string Message, string FacilityNumber, string Account,
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
                "INSERT INTO Script_Industrial_Accounts ( " +
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

    }

    public class FacilityRow
    {
        public string FacilityNumber { get; set; }

        public string ConnectionString { get; set; }

    }


    class AccountRow
    {
        public string Account { get; set; }
        public decimal Amount { get; set; }
        public string Tcode { get; set; }
    }

}


