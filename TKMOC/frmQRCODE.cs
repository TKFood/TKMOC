using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmQRCODE : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        public frmQRCODE()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
            comboBox3load();
        }

        #region FUNCTION

        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [CorporationId],[Name] FROM [HRMDB].[dbo].[Corporation] WHERE [Name]='老楊食品股份有限公司' UNION ALL  SELECT [CorporationId],[Name] FROM [HRMDB].[dbo].[Corporation] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("CorporationId", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "Name";
            comboBox1.DisplayMember = "Name";
            sqlConn.Close();


        }
        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [Code],[Name] FROM [HRMDB].[dbo].[Department] WHERE   [Name] NOT LIKE '%停用%' AND [Code]  LIKE '103%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("Code", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "Name";
            comboBox2.DisplayMember = "Name";
            sqlConn.Close();


        }
        public void comboBox3load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [Name] FROM [HRMDB].[dbo].[EmployeeState]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("Name", typeof(string));

            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "Name";
            comboBox3.DisplayMember = "Name";
            sqlConn.Close();


        }
        public void SETFASTREPORT()
        {

            string SQL;
            Report report1 = new Report();
            report1.Load(@"REPORT\員工QRCODE.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT [Employee].[EmployeeId],[Employee].[CnName],[Employee].[JobId],[Employee].[PartTimeJob],[Employee].[Code] ");
            FASTSQL.AppendFormat(@"  ,[Department].[Name],[Corporation].[Name],[EmployeeState].[Name]");
            FASTSQL.AppendFormat(@"  FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department],[HRMDB].[dbo].[EmployeeState],[HRMDB].[dbo].[Corporation]");
            FASTSQL.AppendFormat(@"  WHERE [Employee].[DepartmentId]=[Department].[DepartmentId]");
            FASTSQL.AppendFormat(@"  AND [EmployeeState].EmployeeStateId=[Employee].EmployeeStateId");
            FASTSQL.AppendFormat(@"  AND [Employee].[CorporationId]=[Corporation].[CorporationId]");
            FASTSQL.AppendFormat(@"  AND [Corporation].[Name]='{0}'",comboBox1.Text);
            FASTSQL.AppendFormat(@"  AND [Department].[Name]='{0}'", comboBox2.Text);
            FASTSQL.AppendFormat(@"  AND [EmployeeState].[Name]='{0}'", comboBox3.Text);
            FASTSQL.AppendFormat(@"  ORDER BY [Employee].[Code]");
            FASTSQL.AppendFormat(@"    ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }
        #endregion


        #region BUTTON

        private void button6_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion
    }
}
