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
namespace TKMOC
{
    public partial class frmEngineeringAddEditDel : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EquipmentID;
        Thread TD;

        public frmEngineeringAddEditDel()
        {
            InitializeComponent();
            combobox1load();
        }
        public frmEngineeringAddEditDel(string EquipmentID)
        {
            InitializeComponent();
            combobox1load();
            textBox1.Text = EquipmentID;

            if (!string.IsNullOrEmpty(EquipmentID))
            {
                Search(EquipmentID);
            }
        }

        #region FUNCTION

        public void combobox1load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [UNITID],[UNITNAME] FROM [TKMOC].[dbo].[ENDUNIT]";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("UNITID", typeof(string));
            dt.Columns.Add("UNITNAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "UNITID";
            comboBox1.DisplayMember = "UNITNAME";
            sqlConn.Close();

        }
        public void Search(string ID)
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();               

                sbSql.Append(@" SELECT [ID] AS '設備編號',[NAME]  AS '設備名稱',[UNIT]  AS '單位',[FACTORY]  AS '廠牌',[TYPE]  AS '型別',[MAINTENANCE]  AS '保養',[CHEKCK]  AS '點檢',[STATUS]  AS '狀況說明'  ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[ENGEQUIPMENT] WITH (NOLOCK)");
                sbSql.AppendFormat(@" WHERE [ID] ='{0}'",ID);
                sbSql.Append(@" ORDER BY [ID]  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        textBox1.Text = ds.Tables["TEMPds1"].Rows[0]["設備編號"].ToString();
                        textBox2.Text = ds.Tables["TEMPds1"].Rows[0]["設備名稱"].ToString();
                        textBox3.Text = ds.Tables["TEMPds1"].Rows[0]["廠牌"].ToString();
                        textBox4.Text = ds.Tables["TEMPds1"].Rows[0]["型別"].ToString();
                        textBox5.Text = ds.Tables["TEMPds1"].Rows[0]["保養"].ToString();
                        textBox6.Text = ds.Tables["TEMPds1"].Rows[0]["點檢"].ToString();
                        textBox7.Text = ds.Tables["TEMPds1"].Rows[0]["狀況說明"].ToString();
                        comboBox1.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["單位"].ToString();

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        #endregion

        #region BUTTON

        #endregion
    }
}
