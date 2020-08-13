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
using System.Globalization;
using Calendar.NET;

namespace TKMOC
{
    public partial class frmTRACEBACK : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();

        int result;

        public frmTRACEBACK()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SEARCHOUT(string MB001,string LOTNO)
        {
            StringBuilder sbSql = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT MF001,MF002,'0',MF003,MF004,MF005,MF006,MF010");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE MF001=ME001 AND MF002=ME002");
                sbSql.AppendFormat(@"  AND MF009 IN ('2','5')");
                sbSql.AppendFormat(@"  AND MF001='{0}' AND MF002='{1}'",MB001,LOTNO);
                sbSql.AppendFormat(@"  ORDER BY MF002,MF003,MF004,MF005");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {


                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDTRACEBACKOUT(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" WHERE [MMB001]='{0}' AND [MLOTNO]='{1}'",MB001,LOTNO);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                sbSql.AppendFormat(" SELECT MF001,MF002,'銷貨','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010");
                sbSql.AppendFormat(" FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK)");
                sbSql.AppendFormat(" WHERE MF001=ME001 AND MF002=ME002");
                sbSql.AppendFormat(" AND MF009 IN ('2','5')");
                sbSql.AppendFormat(" AND MF001='{0}' AND MF002='{1}'", MB001, LOTNO);
                sbSql.AppendFormat(" ORDER BY INVMF.MF002,MF003,MF004,MF005");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        #endregion

        #region BUTTON


        private void button1_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox1.Text)&& !string.IsNullOrEmpty(textBox2.Text))
            {
                if(comboBox1.Text.Trim().Equals("成品逆溯"))
                {
                    SEARCHOUT(textBox1.Text.Trim(), textBox2.Text.Trim());
                }
                else if (comboBox1.Text.Trim().Equals("原料順溯"))
                {
                   
                }


            }
        }

        #endregion
    }
}
