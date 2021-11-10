using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace AVFGService
{
    public class GetExternalData
    {
        static string ESerName = ConfigurationManager.AppSettings["ExternalServerName"];
        static string EDBName = ConfigurationManager.AppSettings["ExternalDBName"];
        static string EUID = ConfigurationManager.AppSettings["ExternalUserName"];
        static string EPWD = ConfigurationManager.AppSettings["ExternalPassword"];
        static string connection = $"data source={ESerName};initial catalog={EDBName};User ID={EUID};Password={EPWD};integrated security=True;MultipleActiveResultSets=True";
        //static string connection = $"data source={ESerName};initial catalog={EDBName};integrated security=True;MultipleActiveResultSets=True";
        SqlConnection conn = new SqlConnection(connection);
        public void SetLog(string content)
        {
            StreamWriter objSw = null;
            StreamWriter objSw2 = null;
            try
            {
                string AppLocation = "";
                AppLocation = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData); ;
                string folderName = AppLocation + "\\AVFG_LogFiles";
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName);
                }
                string sFilePath = folderName + "\\AVFG_Log-" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
                objSw = new StreamWriter(sFilePath, true);
                objSw.WriteLine(DateTime.Now.ToString() + " " + content + Environment.NewLine);

                string sFilePath2 = AppDomain.CurrentDomain.BaseDirectory + "AVFG_Log-" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
                objSw2 = new StreamWriter(sFilePath2, true);
                objSw2.WriteLine(DateTime.Now.ToString() + " " + content + Environment.NewLine);
            }
            catch (Exception ex)
            {
                SetLog("Error -" + ex.Message);
            }
            finally
            {
                if (objSw != null)
                {
                    objSw.Flush();
                    objSw.Dispose();
                }
            }
        }
        public DataSet GetPendingSalesOrder_Std(int Customerid)
        {
            #region Common Extrenal Data
            string sql = "";
            sql = $@"select a.sName Customer_Name,h.sVoucherNo Sales_Order_Doc_No,cast(dbo.IntToDate(h.iDate) as date) Sales_Order_Doc_Date,xh.Customer_LPO_No LPO_NO,cast(dbo.IntToDate(xh.Customer_LPO_Date) as date) LPO_Date,p.sName Product_Name,p.sCode Product_Code,l.Balance Balance_Qty,cast(dbo.IntToDate(xd.ETD)as date) ETD_Date
from vCore_Links369301505_0 l
                inner join tcore_data_0 d on l.iRefId = d.iTransactionId
				inner join tCore_Indta_0 i on i.iBodyId = d.iBodyId
                inner join mCore_Account a on l.iFilter = a.iMasterId
                inner join tCore_Header_0  h on d.iHeaderId = h.iHeaderId
                left join tCore_HeaderData5635_0 xh on h.iHeaderId = xh.iHeaderId
                inner join mCore_Product p on i.iProduct = p.iMasterId
                left join tCore_Data5635_0 xd on d.iBodyId=xd.iBodyId
                where l.bClosed = 0  and l.iFilter = {Customerid}";

            DataSet dst = GetData(sql);
            return dst;
            #endregion
        }

        public DataSet GetToEmailId()
        {
            //Getting Email Id of Customer
            #region Get Customer Email
            string sql = "";
            sql = $@"
                select iFilter,a.ProcurementEmail from vCore_Links369301505_0 l
                inner join muCore_Account_Details a on l.iFilter = a.iMasterId
                where l.bClosed = 0 and a.ProcurementEmail != ''
                group by iFilter,a.ProcurementEmail";

            DataSet dst = GetData(sql);
            return dst;

            #endregion
        }
        public DataSet GetEmailDetails()
        {
            //Getting Email Id of Customer
            #region Get Customer Email
            string sql = "";
            sql = $@"
                select Subject,Body,dbo.fCore_IntToTime(ScheduledTime)ScheduledTime,
case when ScheduleDay =1 then 'Sunday' 
when ScheduleDay =2 then 'Monday'
when ScheduleDay =3 then 'Tuesday'
when ScheduleDay =4 then 'Wednesday'
when ScheduleDay =5 then 'Thursday'
when ScheduleDay =6 then 'Friday'
when ScheduleDay =7 then 'Saturday'
end as ScheduledDay 

from muCore_EmailConfiguration where iMasterId=1";

            DataSet dst = GetData(sql);
            return dst;

            #endregion
        }

        public DataSet GetFromEmailId()
        {
            //Getting Email Id of Customer
            #region Get From Email
            string sql = "";
            sql = $@"
                select sValue from cCore_PreferenceText_0 where icategory=13 and iFieldId=0
                union all
                select sValue from cCore_PreferenceText_0 where icategory=13 and iFieldId=1
                union all
                select sValue from cCore_PreferenceText_0 where icategory=13 and iFieldId=2";

            DataSet dst = GetData(sql);
            return dst;

            #endregion
        }


        //Get Data from Focus Database Sql server 
        static string FSerName = ConfigurationManager.AppSettings["FocusServerName"];
        static string FDB = ConfigurationManager.AppSettings["FocusDBName"];
        static string FSQLUID = ConfigurationManager.AppSettings["FocusUserName"];
        static string FSQLPWD = ConfigurationManager.AppSettings["FocusPassword"];
        static string Fconnection = $"data source={FSerName};initial catalog={FDB};User ID={FSQLUID};Password={FSQLPWD};integrated security=True;MultipleActiveResultSets=True";
        SqlConnection con = new SqlConnection(Fconnection);


        public DataSet GetData(string Query)
        {
            conn.Open();
            SqlCommand cmd = new SqlCommand(Query, conn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            DataSet dst = ds;
            conn.Close();
            return dst;
        }

        public int Update(string Vouc)
        {
            int result = 0;
            using (SqlConnection connect = new SqlConnection(connection))
            {
                string sql = $"{Vouc}";
                using (SqlCommand command = new SqlCommand(sql, connect))
                {
                    connect.Open();
                    result = command.ExecuteNonQuery();
                    connect.Close();
                }
            }
            return result;
        }
    }
}
