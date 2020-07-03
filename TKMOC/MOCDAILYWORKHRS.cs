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
    public partial class MOCDAILYWORKHRS : Form
    {
        public MOCDAILYWORKHRS()
        {
            InitializeComponent();
        }



        #region FUNCTION
        public void SEARCH(string IDDATE)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT  CONVERT(NVARCHAR,[DATS],112) AS '日期',[MANU] AS '產線別',[TA001] AS '製令單',[TA002] AS '製令單號',[MB001] AS '品號',[MB002] AS '品名',[NUMS] AS '入庫量',[MOCNUM] AS '預計生產量'");
                sbSql.AppendFormat(@"  ,[WORKSTART] AS '開始時間',[WORKEND] AS '結束時間',[WORKHRS] AS '工時',[WORKTIMES] AS '工時(分)',[AVGWORKHRS] AS '平均工時'");
                sbSql.AppendFormat(@"  ,[WATERNOODLESEMP] AS '水麵攪拌',[WATERNOODLESSTART] AS '水麵攪拌開始時間',[WATERNOODLESEND] AS '水麵攪拌結束時間',[WATERNOODLESTIMES] AS '水麵攪拌工時'");
                sbSql.AppendFormat(@"  ,[OILPASTRYEMP] AS '油酥攪拌',[OILPASTRYSTART] AS '油酥攪拌開始時間',[OILPASTRYEND] AS '油酥攪拌結束時間',[OILPASTRYTIMES] AS '油酥攪拌工時'");
                sbSql.AppendFormat(@"  ,[FOLDEMP] AS '摺疊',[FOLDSTART] AS '摺疊開始時間',[FOLDEND] AS '摺疊結束時間',[FOLDTIMES] AS '摺疊工時'");
                sbSql.AppendFormat(@"  ,[TYPECOOKEMP] AS '舖餅',[TYPECOOKSTART] AS '舖餅開始時間',[TYPECOOKEND] AS '舖餅結束時間',[TYPECOOKTIMES] AS '舖餅工時'");
                sbSql.AppendFormat(@"  ,[TYPEEMP] AS '成型/烘烤',[TYPESTART] AS '成型/烘烤開始時間',[TYPEEND] AS '成型/烘烤結束時間',[TYPETIMES] AS '成型/烘烤工時'");
                sbSql.AppendFormat(@"  ,[OVENCOOKEMP] AS '烤箱篩餅',[OVENCOOKSTART] AS '烤箱篩餅開始時間',[OVENCOOKEND] AS '烤箱篩餅結束時間',[OVENCOOKTIMES] AS '烤箱篩餅工時'");
                sbSql.AppendFormat(@"  ,[COLDCOOKEMP] AS '冷卻篩餅',[COLDCOOKSTART] AS '冷卻篩餅開始時間',[COLDCOOKEND] AS '冷卻篩餅結束時間',[COLDCOOKTIMES] AS '冷卻篩餅工時'");
                sbSql.AppendFormat(@"  ,[ARRAYEMP] AS '排餅/裝罐',[ARRAYSTART] AS '排餅/裝罐開始時間',[ARRAYEND] AS '排餅/裝罐結束時間',[ARRAYTIMES] AS '排餅/裝罐工時'");
                sbSql.AppendFormat(@"  ,[PACKEMP] AS '包裝機',[PACKSTART] AS '包裝機開始時間',[PACKEND] AS '包裝機結束時間',[PACKTIMES] AS '包裝機工時'");
                sbSql.AppendFormat(@"  ,[PACKPICKEMP] AS '包裝篩餅',[PACKPICKSTART] AS '包裝篩餅開始時間',[PACKPICKEND] AS '包裝篩餅結束時間',[PACKPICKTIMES] AS '包裝篩餅工時'");
                sbSql.AppendFormat(@"  ,[BOXSEMP] AS '裝箱',[BOXSSTART] AS '裝箱開始時間',[BOXSEND] AS '裝箱結束時間',[BOXSTIMES] AS '裝箱工時'");
                sbSql.AppendFormat(@"  ,[HANDCOOKEMP] AS '撿餅',[HANDCOOKSTART] AS '撿餅開始時間',[HANDCOOKEND] AS '撿餅結束時間',[HANDCOOKTIMES] AS '撿餅工時'");
                sbSql.AppendFormat(@"  ,[SCALESWEIGHTEMP] AS '秤重',[SCALESWEIGHTSTART] AS '秤重',[SCALESWEIGHTEND] AS '秤重結束時間',[SCALESWEIGHTTIMES] AS '秤重工時'");
                sbSql.AppendFormat(@"  ,[OUTBOXSEMP] AS '外裝箱',[OUTBOXSSTART] AS '外裝箱開始時間',[OUTBOXSEND] AS '外裝箱結束時間',[OUTBOXSTIMES] AS '外裝箱工時'");
                sbSql.AppendFormat(@"  ,[SEALEMP] AS '封箱',[SEALSTART] AS '封箱開始時間',[SEALEND] AS '封箱結束時間',[SEALTIMES] AS '封箱工時'");
                sbSql.AppendFormat(@"  ,[THROWEMP] AS '倒餅',[THROWSTART] AS '倒餅開始時間',[THROWEND] AS '倒餅結束時間',[THROWTIMES] AS '倒餅工時'");
                sbSql.AppendFormat(@"  ,[BOXPACKEMP] AS '封盒機',[BOXPACKSTART] AS '封盒機開始時間',[BOXPACKEND] AS '封盒機結束時間',[BOXPACKTIMES] AS '封盒機工時'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCDAILYWORKHRS] ");
                sbSql.AppendFormat(@"  WHERE  CONVERT(NVARCHAR,[DATS],112)='{0}'", IDDATE);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["ds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
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
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
        private void button3_Click(object sender, EventArgs e)
        {

        }
        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        #endregion




    }
}
