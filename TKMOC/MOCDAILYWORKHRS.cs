﻿using System;
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
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        int result;
        string STATUS = null;
        string DELID = null;
     

        public MOCDAILYWORKHRS()
        {
            InitializeComponent();

            SETTIMES();
            comboBox1load();

        }



        #region FUNCTION

        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD002";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void SEARCH(string IDDATE)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
           
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
        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox11.Text)&& !string.IsNullOrEmpty(textBox12.Text))
            {
                SEARCHMOCTA(textBox11.Text.Trim(), textBox12.Text.Trim());
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox11.Text) && !string.IsNullOrEmpty(textBox12.Text))
            {
                SEARCHMOCTA(textBox11.Text.Trim(), textBox12.Text.Trim());
            }
        }

        public void SEARCHMOCTA(string TA001,string TA002)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

            SETTEXT1();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" SELECT TA006,TA034,TA015,TA017 FROM [TK].dbo.MOCTA WHERE TA001='{0}' AND TA002='{1}' ", TA001,TA002);
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
                    
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        textBox21.Text = ds1.Tables["ds1"].Rows[0]["TA006"].ToString();
                        textBox22.Text = ds1.Tables["ds1"].Rows[0]["TA034"].ToString();
                        textBox23.Text = ds1.Tables["ds1"].Rows[0]["TA017"].ToString();
                        textBox24.Text = ds1.Tables["ds1"].Rows[0]["TA015"].ToString();

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

        public void ADDMOCDAILYWORKHRS(string ID, string DATS, string MANU, string TA001, string TA002, string MB001, string MB002, string NUMS, string MOCNUM
            , string WORKSTART, string WORKEND, string WORKHRS, string WORKTIMES, string AVGWORKHRS
            , string WATERNOODLESEMP, string WATERNOODLESSTART, string WATERNOODLESEND, string WATERNOODLESTIMES
            , string OILPASTRYEMP, string OILPASTRYSTART, string OILPASTRYEND, string OILPASTRYTIMES
            , string FOLDEMP, string FOLDSTART, string FOLDEND, string FOLDTIMES
            , string TYPECOOKEMP, string TYPECOOKSTART, string TYPECOOKEND, string TYPECOOKTIMES
            , string TYPEEMP, string TYPESTART, string TYPEEND, string TYPETIMES
            , string OVENCOOKEMP, string OVENCOOKSTART, string OVENCOOKEND, string OVENCOOKTIMES
            , string COLDCOOKEMP, string COLDCOOKSTART, string COLDCOOKEND, string COLDCOOKTIMES
            , string ARRAYEMP, string ARRAYSTART, string ARRAYEND, string ARRAYTIMES
            , string PACKEMP, string PACKSTART, string PACKEND, string PACKTIMES
            , string PACKPICKEMP, string PACKPICKSTART, string PACKPICKEND, string PACKPICKTIMES
            , string BOXSEMP, string BOXSSTART, string BOXSEND, string BOXSTIMES
            , string HANDCOOKEMP, string HANDCOOKSTART, string HANDCOOKEND, string HANDCOOKTIMES
            , string SCALESWEIGHTEMP, string SCALESWEIGHTSTART, string SCALESWEIGHTEND, string SCALESWEIGHTTIMES
            , string OUTBOXSEMP, string OUTBOXSSTART, string OUTBOXSEND, string OUTBOXSTIMES
            , string SEALEMP, string SEALSTART, string SEALEND, string SEALTIMES
            , string THROWEMP, string THROWSTART, string THROWEND, string THROWTIMES
            , string BOXPACKEMP, string BOXPACKSTART, string BOXPACKEND, string BOXPACKTIMES                            
            )
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

                sbSql.AppendFormat("  INSERT INTO [TKMOC].[dbo].[MOCDAILYWORKHRS]");
                sbSql.AppendFormat("  (");
                sbSql.AppendFormat("  [ID],[DATS],[MANU],[TA001],[TA002],[MB001],[MB002],[NUMS],[MOCNUM]");
                sbSql.AppendFormat("  ,[WORKSTART],[WORKEND],[WORKHRS],[WORKTIMES],[AVGWORKHRS]");
                sbSql.AppendFormat("  ,[WATERNOODLESEMP],[WATERNOODLESSTART],[WATERNOODLESEND],[WATERNOODLESTIMES]");
                sbSql.AppendFormat("  ,[OILPASTRYEMP],[OILPASTRYSTART],[OILPASTRYEND],[OILPASTRYTIMES]");
                sbSql.AppendFormat("  ,[FOLDEMP],[FOLDSTART],[FOLDEND],[FOLDTIMES]");
                sbSql.AppendFormat("  ,[TYPECOOKEMP],[TYPECOOKSTART],[TYPECOOKEND],[TYPECOOKTIMES]");
                sbSql.AppendFormat("  ,[TYPEEMP],[TYPESTART],[TYPEEND],[TYPETIMES]");
                sbSql.AppendFormat("  ,[OVENCOOKEMP],[OVENCOOKSTART],[OVENCOOKEND],[OVENCOOKTIMES]");
                sbSql.AppendFormat("  ,[COLDCOOKEMP],[COLDCOOKSTART],[COLDCOOKEND],[COLDCOOKTIMES]");
                sbSql.AppendFormat("  ,[ARRAYEMP],[ARRAYSTART],[ARRAYEND],[ARRAYTIMES]");
                sbSql.AppendFormat("  ,[PACKEMP],[PACKSTART],[PACKEND],[PACKTIMES]");
                sbSql.AppendFormat("  ,[PACKPICKEMP],[PACKPICKSTART],[PACKPICKEND],[PACKPICKTIMES]");
                sbSql.AppendFormat("  ,[BOXSEMP],[BOXSSTART],[BOXSEND],[BOXSTIMES]");
                sbSql.AppendFormat("  ,[HANDCOOKEMP],[HANDCOOKSTART],[HANDCOOKEND],[HANDCOOKTIMES]");
                sbSql.AppendFormat("  ,[SCALESWEIGHTEMP],[SCALESWEIGHTSTART],[SCALESWEIGHTEND],[SCALESWEIGHTTIMES]");
                sbSql.AppendFormat("  ,[OUTBOXSEMP],[OUTBOXSSTART],[OUTBOXSEND],[OUTBOXSTIMES]");
                sbSql.AppendFormat("  ,[SEALEMP],[SEALSTART],[SEALEND],[SEALTIMES]");
                sbSql.AppendFormat("  ,[THROWEMP],[THROWSTART],[THROWEND],[THROWTIMES]");
                sbSql.AppendFormat("  ,[BOXPACKEMP],[BOXPACKSTART],[BOXPACKEND],[BOXPACKTIMES]");
                sbSql.AppendFormat("  )");
                sbSql.AppendFormat("  VALUES");
                sbSql.AppendFormat("  (");
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',", ID, DATS, MANU, TA001, TA002, MB001, MB002, NUMS, MOCNUM);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}','{4}',", WORKSTART, WORKEND, WORKHRS, WORKTIMES, AVGWORKHRS);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", WATERNOODLESEMP, WATERNOODLESSTART, WATERNOODLESEND, WATERNOODLESTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", OILPASTRYEMP, OILPASTRYSTART, OILPASTRYEND, OILPASTRYTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", FOLDEMP, FOLDSTART, FOLDEND, FOLDTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", TYPECOOKEMP, TYPECOOKSTART, TYPECOOKEND, TYPECOOKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", TYPEEMP, TYPESTART, TYPEEND, TYPETIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", OVENCOOKEMP, OVENCOOKSTART, OVENCOOKEND, OVENCOOKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", COLDCOOKEMP, COLDCOOKSTART, COLDCOOKEND, COLDCOOKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", ARRAYEMP, ARRAYSTART, ARRAYEND, ARRAYTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", PACKEMP, PACKSTART, PACKEND, PACKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", PACKPICKEMP, PACKPICKSTART, PACKPICKEND, PACKPICKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", BOXSEMP, BOXSSTART, BOXSEND, BOXSTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", HANDCOOKEMP, HANDCOOKSTART, HANDCOOKEND, HANDCOOKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", SCALESWEIGHTEMP, SCALESWEIGHTSTART, SCALESWEIGHTEND, SCALESWEIGHTTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", OUTBOXSEMP, OUTBOXSSTART, OUTBOXSEND, OUTBOXSTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", SEALEMP, SEALSTART, SEALEND, SEALTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", THROWEMP, THROWSTART, THROWEND, THROWTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}'", BOXPACKEMP, BOXPACKSTART, BOXPACKEND, BOXPACKTIMES);
                sbSql.AppendFormat("  )");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");

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

        public void UPDATEMOCDAILYWORKHRS()
        {

        }

        public void DELMOCDAILYWORKHRS()
        {

        }

        private void dateTimePicker37_ValueChanged(object sender, EventArgs e)
        {
            SETSTARTTIMES(dateTimePicker37.Value);
            CALTIMES();
        }

        private void dateTimePicker38_ValueChanged(object sender, EventArgs e)
        {
            SETENDTIMES(dateTimePicker38.Value);
            CALTIMES();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker5_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker7_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker9_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker10_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker11_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker12_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker13_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker14_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker15_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker16_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker17_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker18_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker19_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker20_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker21_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker22_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker23_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker24_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker25_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker26_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker27_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker28_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker29_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker30_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker31_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker32_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker33_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker34_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker35_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker36_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }
        public void CALTIMES()
        {
            TimeSpan TS = new TimeSpan();
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker38.Value.Ticks - dateTimePicker37.Value.Ticks).TotalHours) > 0)
            {            
                numericUpDown11.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker38.Value.Ticks - dateTimePicker37.Value.Ticks).TotalHours);
            }

            if (Convert.ToDecimal(new TimeSpan(dateTimePicker3.Value.Ticks - dateTimePicker2.Value.Ticks).TotalHours)>0)
            {
                numericUpDown31.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker3.Value.Ticks - dateTimePicker2.Value.Ticks).TotalHours);
            }

            if (Convert.ToDecimal(new TimeSpan(dateTimePicker6.Value.Ticks - dateTimePicker5.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown41.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker6.Value.Ticks - dateTimePicker5.Value.Ticks).TotalHours);
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker8.Value.Ticks - dateTimePicker7.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown51.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker8.Value.Ticks - dateTimePicker7.Value.Ticks).TotalHours);
               
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker10.Value.Ticks - dateTimePicker9.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown61.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker10.Value.Ticks - dateTimePicker9.Value.Ticks).TotalHours);
               
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker12.Value.Ticks - dateTimePicker11.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown71.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker12.Value.Ticks - dateTimePicker11.Value.Ticks).TotalHours);
                
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker14.Value.Ticks - dateTimePicker13.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown81.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker14.Value.Ticks - dateTimePicker13.Value.Ticks).TotalHours);
               
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker16.Value.Ticks - dateTimePicker15.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown91.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker16.Value.Ticks - dateTimePicker15.Value.Ticks).TotalHours);
               
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker18.Value.Ticks - dateTimePicker17.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown101.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker18.Value.Ticks - dateTimePicker17.Value.Ticks).TotalHours);
                
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker20.Value.Ticks - dateTimePicker19.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown111.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker20.Value.Ticks - dateTimePicker19.Value.Ticks).TotalHours);
               
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker22.Value.Ticks - dateTimePicker21.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown121.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker22.Value.Ticks - dateTimePicker21.Value.Ticks).TotalHours);
                
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker24.Value.Ticks - dateTimePicker23.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown131.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker24.Value.Ticks - dateTimePicker23.Value.Ticks).TotalHours);

              
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker26.Value.Ticks - dateTimePicker25.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown32.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker26.Value.Ticks - dateTimePicker25.Value.Ticks).TotalHours);
              
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker28.Value.Ticks - dateTimePicker27.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown42.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker28.Value.Ticks - dateTimePicker27.Value.Ticks).TotalHours);
               
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker30.Value.Ticks - dateTimePicker29.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown52.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker30.Value.Ticks - dateTimePicker29.Value.Ticks).TotalHours);
              
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker32.Value.Ticks - dateTimePicker31.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown62.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker32.Value.Ticks - dateTimePicker31.Value.Ticks).TotalHours);
                
            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker34.Value.Ticks - dateTimePicker33.Value.Ticks).TotalHours) > 0)
            {               
                numericUpDown72.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker34.Value.Ticks - dateTimePicker33.Value.Ticks).TotalHours);
                            }
            if (Convert.ToDecimal(new TimeSpan(dateTimePicker36.Value.Ticks - dateTimePicker35.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown82.Value = Convert.ToDecimal(new TimeSpan(dateTimePicker36.Value.Ticks - dateTimePicker35.Value.Ticks).TotalHours);

            }



        }

        public void SETTEXT1()
        {
            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
        }

        public void SETTIMES()
        {
            DateTime dt = Convert.ToDateTime(DateTime.Now.Year+"/"+ DateTime.Now.Month + "/" + DateTime.Now.Day+" "+ DateTime.Now.Hour+":"+ DateTime.Now.Minute);
            dateTimePicker2.Value = dt;
            dateTimePicker5.Value = dt;
            dateTimePicker7.Value = dt;
            dateTimePicker9.Value = dt;
            dateTimePicker11.Value = dt;
            dateTimePicker13.Value = dt;
            dateTimePicker15.Value = dt;
            dateTimePicker17.Value = dt;
            dateTimePicker19.Value = dt;
            dateTimePicker21.Value = dt;
            dateTimePicker23.Value = dt;
            dateTimePicker25.Value = dt;
            dateTimePicker27.Value = dt;
            dateTimePicker29.Value = dt;
            dateTimePicker31.Value = dt;
            dateTimePicker33.Value = dt;
            dateTimePicker35.Value = dt;

            dateTimePicker3.Value = dt;
            dateTimePicker6.Value = dt;
            dateTimePicker8.Value = dt;
            dateTimePicker10.Value = dt;
            dateTimePicker12.Value = dt;
            dateTimePicker14.Value = dt;
            dateTimePicker16.Value = dt;
            dateTimePicker18.Value = dt;
            dateTimePicker20.Value = dt;
            dateTimePicker22.Value = dt;
            dateTimePicker24.Value = dt;
            dateTimePicker26.Value = dt;
            dateTimePicker28.Value = dt;
            dateTimePicker30.Value = dt;
            dateTimePicker32.Value = dt;
            dateTimePicker34.Value = dt;
            dateTimePicker36.Value = dt;

            dateTimePicker37.Value = dt;
            dateTimePicker38.Value = dt;
        }

        public void SETSTARTTIMES(DateTime dt)
        {
            dateTimePicker2.Value = dt;
            dateTimePicker5.Value = dt;
            dateTimePicker7.Value = dt;
            dateTimePicker9.Value = dt;
            dateTimePicker11.Value = dt;
            dateTimePicker13.Value = dt;
            dateTimePicker15.Value = dt;
            dateTimePicker17.Value = dt;
            dateTimePicker19.Value = dt;
            dateTimePicker21.Value = dt;
            dateTimePicker23.Value = dt;
            dateTimePicker25.Value = dt;
            dateTimePicker27.Value = dt;
            dateTimePicker29.Value = dt;
            dateTimePicker31.Value = dt;
            dateTimePicker33.Value = dt;
            dateTimePicker35.Value = dt;
        }

        public void SETENDTIMES(DateTime dt)
        {
            dateTimePicker3.Value = dt;
            dateTimePicker6.Value = dt;
            dateTimePicker8.Value = dt;
            dateTimePicker10.Value = dt;
            dateTimePicker12.Value = dt;
            dateTimePicker14.Value = dt;
            dateTimePicker16.Value = dt;
            dateTimePicker18.Value = dt;
            dateTimePicker20.Value = dt;
            dateTimePicker22.Value = dt;
            dateTimePicker24.Value = dt;
            dateTimePicker26.Value = dt;
            dateTimePicker28.Value = dt;
            dateTimePicker30.Value = dt;
            dateTimePicker32.Value = dt;
            dateTimePicker34.Value = dt;
            dateTimePicker36.Value = dt;
        }


        public void SETTEXTBOX1()
        {
            textBox11.Text = null;
            textBox12.Text = null;
            textBox31.Text = null;
            textBox41.Text = null;
            textBox51.Text = null;
            textBox61.Text = null;
            textBox71.Text = null;
            textBox81.Text = null;
            textBox91.Text = null;
            textBox101.Text = null;
            textBox111.Text = null;
            textBox121.Text = null;
            textBox131.Text = null;
            textBox32.Text = null;
            textBox42.Text = null;
            textBox52.Text = null;
            textBox62.Text = null;
            textBox72.Text = null;
            textBox82.Text = null;
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            STATUS = "ADD";
            label26.Text= "ADD";
            //SETTEXTBOX1();

        }
        private void button3_Click(object sender, EventArgs e)
        {
            STATUS = "EDIT";
            label26.Text = "EDIT";

        }
        private void button4_Click(object sender, EventArgs e)
        {
            if(STATUS.Equals("ADD"))
            {
                ADDMOCDAILYWORKHRS(Guid.NewGuid().ToString(), dateTimePicker4.Value.ToString("yyyy/MM/dd"), comboBox1.Text.Trim(), textBox11.Text, textBox12.Text, textBox21.Text, textBox22.Text, textBox23.Text, textBox24.Text
                , dateTimePicker37.Value.ToString("HH:ss"), dateTimePicker38.Value.ToString("HH:ss"), numericUpDown11.Value.ToString(), "0","0"
                , textBox31.Text, dateTimePicker2.Value.ToString("HH:ss"), dateTimePicker3.Value.ToString("HH:ss"), numericUpDown31.Value.ToString()
                , textBox41.Text, dateTimePicker5.Value.ToString("HH:ss"), dateTimePicker6.Value.ToString("HH:ss"), numericUpDown41.Value.ToString()
                , textBox51.Text, dateTimePicker7.Value.ToString("HH:ss"), dateTimePicker8.Value.ToString("HH:ss"), numericUpDown51.Value.ToString()
                , textBox61.Text, dateTimePicker9.Value.ToString("HH:ss"), dateTimePicker10.Value.ToString("HH:ss"), numericUpDown61.Value.ToString()
                , textBox71.Text, dateTimePicker11.Value.ToString("HH:ss"), dateTimePicker12.Value.ToString("HH:ss"), numericUpDown71.Value.ToString()
                , textBox81.Text, dateTimePicker13.Value.ToString("HH:ss"), dateTimePicker14.Value.ToString("HH:ss"), numericUpDown81.Value.ToString()
                , textBox91.Text, dateTimePicker15.Value.ToString("HH:ss"), dateTimePicker16.Value.ToString("HH:ss"), numericUpDown91.Value.ToString()
                , textBox101.Text, dateTimePicker17.Value.ToString("HH:ss"), dateTimePicker18.Value.ToString("HH:ss"), numericUpDown101.Value.ToString()
                , textBox111.Text, dateTimePicker19.Value.ToString("HH:ss"), dateTimePicker20.Value.ToString("HH:ss"), numericUpDown111.Value.ToString()
                , textBox121.Text, dateTimePicker21.Value.ToString("HH:ss"), dateTimePicker22.Value.ToString("HH:ss"), numericUpDown121.Value.ToString()
                , textBox131.Text, dateTimePicker23.Value.ToString("HH:ss"), dateTimePicker24.Value.ToString("HH:ss"), numericUpDown131.Value.ToString()
                , textBox32.Text, dateTimePicker25.Value.ToString("HH:ss"), dateTimePicker26.Value.ToString("HH:ss"), numericUpDown32.Value.ToString()
                , textBox42.Text, dateTimePicker27.Value.ToString("HH:ss"), dateTimePicker28.Value.ToString("HH:ss"), numericUpDown42.Value.ToString()
                , textBox52.Text, dateTimePicker29.Value.ToString("HH:ss"), dateTimePicker30.Value.ToString("HH:ss"), numericUpDown52.Value.ToString()
                , textBox62.Text, dateTimePicker31.Value.ToString("HH:ss"), dateTimePicker32.Value.ToString("HH:ss"), numericUpDown62.Value.ToString()
                , textBox72.Text, dateTimePicker33.Value.ToString("HH:ss"), dateTimePicker34.Value.ToString("HH:ss"), numericUpDown72.Value.ToString()
                , textBox82.Text, dateTimePicker35.Value.ToString("HH:ss"), dateTimePicker36.Value.ToString("HH:ss"), numericUpDown82.Value.ToString()
                );
            }
            else if(STATUS.Equals("EDIT"))
            {

            }

            STATUS = null;
            label26.Text = "STATUS";
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }








        #endregion

       
    }
}
