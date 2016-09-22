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
    public partial class frmMOCPRODUCTDAILYREPORT : Form
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
        string EDITID;
        int result;
        Thread TD;

        public frmMOCPRODUCTDAILYREPORT()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@" SELECT [PRODUCEDATE] AS '日期',[PRODUCEID]  AS '製令單號',[PRODUCENAME]  AS '品號/品名' ");
                sbSql.Append(@" ,[PREIN]  AS '預計投入量(kg)',[PROCESSTIME]  AS '製造時間(分)',[PRECESSPEOPEL]  AS '製造人數'");
                sbSql.Append(@" ,[MATERIAL]  AS '油酥原料(1)',[RECYCLEMATERIAL]  AS '油酥可回收餅麩(2)',[MATERIALTIME]  AS '水面製造時間(分)(3)'");
                sbSql.Append(@" ,[MATERIALPEOPLE]  AS '人數(4)',[WMATERIAL]  AS '水面原料(5)',[WRECYCLESIDE]  AS '水面可回收邊料(6)'");
                sbSql.Append(@" ,[WRECYCLECOOKIES]  AS '水面可回收餅麩(7)',[PACKAGETIME]  AS '包裝時間(內包裝區+罐裝)(分)(8)'");
                sbSql.Append(@" ,[PACKAGEPEOPLE]  AS '包裝人數(9)',[TRECYCLE]  AS '可回收餅麩(10)=(2)+(7)(kg)',[NGTATOL]  AS '未熟總量(11)(kg)'");
                sbSql.Append(@" ,[NGTIME]  AS '未熟烤焙時間(12)(分)',[WEIGHTBEFORE]  AS '烤前單片重量(g)',[WEIGHTAFTER]  AS '烤後單片重量(g)'");
                sbSql.Append(@" ,[COOKIENUM]  AS '每排數量',[BLADENUM]  AS '刀數',[INBEFORE]  AS '烤前實際總投入(kg)',[INAFTER]  AS '烤後實際總投入(kg)'");
                sbSql.Append(@" ,[PREOUT]  AS '預計產出(kg)',[NGSTIR]  AS '製造課不良-攪拌(kg)',[NGSIDE]  AS '製造課不良-成型邊料(kg)'");
                sbSql.Append(@" ,[NGCOOKIES]  AS '製造課不良-餅麩(kg)',[NGBAKE]  AS '製造課不良-烤焙(kg)',[NGNOGOOD]  AS '包裝課不良-包裝不良餅乾(kg)'");
                sbSql.Append(@" ,[NGNOCAN]  AS '包裝課不良-包裝(內袋(卷) 罐)',[PACKAGEWEIGHT]  AS '袋重(kg)',[PACKAGEIN]  AS '包裝投入(袋(卷),罐)'");
                sbSql.Append(@" ,[ACTUALOUT]  AS '實際產出(kg)(裸餅)',[HALFOUT]  AS '半成品入庫數(kg) (含袋重)',[REMARK]  AS '備註'");
                sbSql.Append(@" ,[MANUTIME]  AS '製造工時',[STIRPCT]  AS '攪拌成型製成率%',[EVARATE]  AS '蒸發率'");
                sbSql.Append(@" ,[LOSTRATE]  AS '製成損失率',[HALFRATE]  AS '半成品產出效率',[PACKAGETOTALTIME]  AS '包裝工時'");
                sbSql.Append(@",[WEIGHTTOTAL]  AS '袋重',[PACKAGELOSTRATE]  AS '包裝損耗率',[TOTALRATE]  AS '製成率' ");
                sbSql.Append(@" ,[CANCOOKIESWEIGHT]  AS '一箱裸餅重',[CANWEIGHT]   AS '一箱餅含袋重',[ID]");
                sbSql.Append(@" ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT]  WITH (NOLOCK)");
                //sbSql.AppendFormat(@" WHERE [ID] ='{0}'", ID);
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

                        labelSearch.Text = "有 " + ds.Tables["TEMPds1"].Rows.Count.ToString() + " 筆";
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = ds.Tables["TEMPds1"].Rows[0]["設備名稱"].ToString();
            textBox2.Text = ds.Tables["TEMPds1"].Rows[0]["故障情形"].ToString();
            textBox3.Text = ds.Tables["TEMPds1"].Rows[0]["維修時數"].ToString();
            dateTimePicker1.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["維修時間起"].ToString());
            dateTimePicker2.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["維修時間迄"].ToString());

        }
        public void UPDATE()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MAINRECORD] ");
                //sbSql.AppendFormat("  SET [EQUIPMENTID]='{1}',[EQUIPMENTNAME]='{2}',[UNIT]='{3}',[ERROR]='{4}',[MAINDATEBEGIN]='{5}',[MAINDATEEND]='{6}',[MAINDATHR]='{7}',[MAINEMP]='{8}',[MALFUNCIONID]='{9}',[MAINSTATUS]='{10}',[MAINUSED]='{11}' WHERE [ID]='{0}' ", textBox8.Text.ToString(), comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(), comboBox1.SelectedValue.ToString(), textBox2.Text.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm"), dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm"), textBox3.Text.ToString(), textBox4.Text.ToString(), comboBox3.SelectedValue.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString());
                sbSql.Append("   ");

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
                    this.Close();

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
        public void ADD()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MAINRECORD] ");
                sbSql.Append(" ([ID],[EQUIPMENTID],[EQUIPMENTNAME],[UNIT],[ERROR],[MAINDATEBEGIN],[MAINDATEEND],[MAINDATHR],[MAINEMP],[MALFUNCIONID],[MAINSTATUS],[MAINUSED])  ");
                //sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}') ", Guid.NewGuid(), comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(), comboBox1.SelectedValue.ToString(), textBox2.Text.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm"), dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm"), textBox3.Text.ToString(), textBox4.Text.ToString(), comboBox3.SelectedValue.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString());

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
                    this.Close();

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
        
        public void SetADD()
        {

        }

        public void SetUPDATE()
        {

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SetADD();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SetUPDATE();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textID.Text.ToString()))
            {
                UPDATE();
            }
            else
            {
                ADD();
            }
        }



        #endregion

       
    }
}
