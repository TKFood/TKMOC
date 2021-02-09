﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmMOCCHANGE : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();

        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();

        DataSet dsCHECKMOCTDMOCTG = new DataSet();
        DataTable dt = new DataTable();
        SqlTransaction tran;
        int result;

        string tablename = null;
        int rownum = 0;

        string TA001;
        string TA002;
        string OLDMB001;
        string NEWMB001;

        public frmMOCCHANGE()
        {
            InitializeComponent();

            comboBox1load();
            comboBox1load2("");
            comboBox1load3();
        }

        private void frmMOCCHANGE_Load(object sender, EventArgs e)
        {
            //dataGridView1
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            //dataGridView3
            dataGridView3.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol3 = new DataGridViewCheckBoxColumn();
            cbCol3.Width = 50;   //設定寬度
            cbCol3.HeaderText = "選擇";
            cbCol3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol3.TrueValue = true;
            cbCol3.FalseValue = false;
            dataGridView3.Columns.Insert(0, cbCol3);

            //dataGridView5
            dataGridView5.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol5 = new DataGridViewCheckBoxColumn();
            cbCol5.Width = 50;   //設定寬度
            cbCol5.HeaderText = "選擇";
            cbCol5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol5.TrueValue = true;
            cbCol5.FalseValue = false;
            dataGridView5.Columns.Insert(0, cbCol5);

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol7 = new DataGridViewCheckBoxColumn();
            cbCol7.Width = 50;   //設定寬度
            cbCol7.HeaderText = "選擇";
            cbCol7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol7.TrueValue = true;
            cbCol7.FalseValue = false;
            dataGridView7.Columns.Insert(0, cbCol7);
        }

        #region FUNCTION
        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT BOMMB.MB001,RTRIM(LTRIM(BOMMB.MB001))+' '+INVMB.MB002 AS MB002,BOMMB.MB004 FROM [TK].dbo.BOMMB,[TK] .dbo.INVMB WHERE BOMMB.MB001=INVMB.MB001 AND BOMMB.MB001 LIKE '1%'  GROUP BY BOMMB.MB001,INVMB.MB002,BOMMB.MB004  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MB001";
            comboBox1.DisplayMember = "MB002";
            sqlConn.Close();


        }

        public void comboBox1load2(string MB001)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT BOMMB.MB004,RTRIM(LTRIM(BOMMB.MB004))+' '+INVMB.MB002 AS MB002 FROM [TK].dbo.BOMMB,[TK] .dbo.INVMB WHERE BOMMB.MB004=INVMB.MB001  AND BOMMB.MB001 LIKE '1%' AND BOMMB.MB001='{0}' ", MB001);
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB004", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MB004";
            comboBox2.DisplayMember = "MB002";
            sqlConn.Close();


        }

        public void comboBox1load3()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001,MC002 FROM [TK].dbo.CMSMC    WHERE MC001 IN ('20006','20004','20005') ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "MC001";
            comboBox3.DisplayMember = "MC001";
            sqlConn.Close();


        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox1load2(comboBox1.SelectedValue.ToString().Trim());
        }


        public void SEARCH()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位',TA021 AS '線別',TA026 AS '訂單',TA027 AS '單號',TA028 AS '序號'
                                    FROM [TK].dbo.MOCTA
                                    WHERE TA003>='{0}' AND TA003<='{1}' 
                                    ORDER BY TA001,TA002,TA003
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds"];
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

            }
        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }
        public void CHANGEMULTI()
        {
            OLDMB001 = comboBox1.SelectedValue.ToString().Trim();
            NEWMB001 = comboBox2.SelectedValue.ToString().Trim();

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    TA001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString();
                    TA002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString();

                    //MessageBox.Show(OLDTA001+"-"+ OLDTA002);
                    if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002) && !string.IsNullOrEmpty(OLDMB001) && !string.IsNullOrEmpty(NEWMB001))
                    {
                        UPDATEMOCTB(TA001, TA002, OLDMB001, NEWMB001);
                    }
                }
                else
                {
                    TA001 = null;
                    TA002 = null;
                }
            }

        }

        public void UPDATEMOCTB(string TA001, string TA002, string OLDMB001, string NEWMB001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" UPDATE [TK].dbo.MOCTB ");
                sbSql.AppendFormat(" SET TB003=INVMB.MB001,TB012=INVMB.MB002,TB013=INVMB.MB003");
                sbSql.AppendFormat(" FROM [TK].dbo.INVMB");
                sbSql.AppendFormat(" WHERE INVMB.MB001='{0}'",NEWMB001);
                sbSql.AppendFormat(" AND TB001='{0}' AND TB002='{1}' AND TB003='{2}' ",TA001, TA002,OLDMB001);
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    TA001 = row.Cells["製令"].Value.ToString().Trim();
                    TA002 = row.Cells["單號"].Value.ToString().Trim();
                    
                    SEARCHMOCTB(TA001, TA002);

                }
                else
                {
                    SEARCHMOCTB("", "");

                }
            }
        }

        public void SEARCHMOCTB(string ta001,string TA002)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TB003 AS '材料品號',TB012 AS '材料品名',TB004 AS '需領用量'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTB");
                sbSql.AppendFormat(@"  WHERE TB001='{0}' AND TB002='{1}'",TA001,TA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        public void SEARCH2(string SDAY,string EDAY)
        {
            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();
          
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
          
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA009 AS '預計開工',TA012 AS '實際開工',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位',TA021 AS '線別',TA026 AS '訂單',TA027 AS '單號',TA028 AS '序號'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  WHERE TA003>='{0}' AND TA003<='{1}' ", SDAY,EDAY);
                sbSql.AppendFormat(@"  ORDER BY TA001,TA002,TA003");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds.Tables["TEMPds"];
                        dataGridView3.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        public void SEARCH3(string SDAY, string EDAY)
        {
            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA001 AS '製令',TA002 AS '單號',TB003 AS '品號',TB012 AS '品名',TB004 AS '需領用量',TB009 AS '庫別' ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND TB003='106061011'");
                sbSql.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}'",SDAY,EDAY);
                sbSql.AppendFormat(@"  ORDER BY TA001,TA002,TB003");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds.Tables["TEMPds"];
                        dataGridView5.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    string TA001 = row.Cells["製令"].Value.ToString().Trim();
                    string TA002 = row.Cells["單號"].Value.ToString().Trim();

                    SEARCHMOCTB2(TA001, TA002);

                }
                else
                {
                    SEARCHMOCTB2("", "");

                }
            }
        }

        public void SEARCHMOCTB2(string TA001, string TA002)
        {
            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA009 AS '預計開工',TA012 AS '實際開工' ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  WHERE TA001='{0}' AND TA002='{1}'", TA001, TA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds.Tables["ds"];
                        dataGridView4.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        public void CHANGEMULTI2()
        {
            string NEWDATES = dateTimePicker5.Value.ToString("yyyyMMdd");

            foreach (DataGridViewRow dr in this.dataGridView3.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    string TA001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString();
                    string TA002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString();

                    //MessageBox.Show(OLDTA001+"-"+ OLDTA002);
                    if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002) && !string.IsNullOrEmpty(NEWDATES) )
                    {
                        UPDATEMOCTA(TA001, TA002, NEWDATES);
                    }
                }
                else
                {
                    TA001 = null;
                    TA002 = null;
                }
            }

        }

        public void UPDATEMOCTA(string TA001, string TA002, string TA009)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                sbSql.AppendFormat(" UPDATE [TK].dbo.MOCTA");
                sbSql.AppendFormat(" SET TA009='{0}',TA012='{0}'", TA009);
                sbSql.AppendFormat(" WHERE TA001='{0}' AND TA002='{1}'",TA001,TA002);
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

        public void CHANGEMULTI3()
        {
            foreach (DataGridViewRow dr in this.dataGridView5.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    string TB001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString();
                    string TB002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString();
                    string TB009 = comboBox3.Text;

                    if(!string.IsNullOrEmpty(TB001)&& !string.IsNullOrEmpty(TB002) && !string.IsNullOrEmpty(TB009) )
                    {
                        UPDATEMOCTB(TB001.Trim(), TB002.Trim(), TB009.Trim());
                    }

                    //MessageBox.Show(OLDTA001+"-"+ OLDTA002);

                }
                else
                {
                  
                }
            }

        }

        public void UPDATEMOCTB(string TB001, string TB002, string TB009)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                sbSql.AppendFormat(" UPDATE [TK].dbo.MOCTB");
                sbSql.AppendFormat(" SET TB009='{0}'", TB009);
                sbSql.AppendFormat(" WHERE TB003='106061011' AND TB001='{0}' AND TB002='{1}'", TB001, TB002);
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


        public void SEARCH4(string SDATES,string EDATES)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位',TA021 AS '線別',TA026 AS '訂單',TA027 AS '單號',TA028 AS '序號'
                                    FROM [TK].dbo.MOCTA
                                    WHERE TA003>='{0}' AND TA003<='{1}' 
                                    ORDER BY TA001,TA002,TA003
                                    ", SDATES, EDATES);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView7.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView7.DataSource = ds.Tables["TEMPds"];
                        dataGridView7.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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


        public void CHANGEMULTI4()
        {
            foreach (DataGridViewRow dr in this.dataGridView7.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    string TB001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString();
                    string TB002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString();

                    if (!string.IsNullOrEmpty(TB001) && !string.IsNullOrEmpty(TB002) )
                    {
                        UPDATEMOCTB2(TB001.Trim(), TB002.Trim());
                    }

                    //MessageBox.Show(OLDTA001+"-"+ OLDTA002);

                }
                else
                {

                }
            }

        }

        public void UPDATEMOCTB2(string TB001, string TB002)
        {
            string SQLLIKE = SEARCHMOCCHANGE();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                sbSql.AppendFormat(@" 
                                    UPDATE [TK].dbo.MOCTB
                                    SET TB004=ROUND(TB004,0)
                                    WHERE ( {2} )
                                    AND TB001='{0}' AND TB002='{1}'
                                    ", TB001, TB002, SQLLIKE.ToString());



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

        public string SEARCHMOCCHANGE()
        {
            StringBuilder MB001 = new StringBuilder();
            MB001.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT [MB001] FROM [TKMOC].[dbo].[MOCCHANGE]
                                    ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count >= 1)
                {
                    for (int i = 0; i < ds4.Tables["ds4"].Rows.Count; i++)
                    {
                        MB001.AppendFormat(@" (TB003 LIKE '{0}%') OR ", ds4.Tables["ds4"].Rows[i]["MB001"].ToString());
                    }

                    MB001.AppendFormat(@" (TB003 LIKE 'NA%') ");
                    return MB001.ToString();

                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {

            }
        }
        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView7.CurrentRow != null)
            {
                int rowindex = dataGridView7.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView7.Rows[rowindex];
                    TA001 = row.Cells["製令"].Value.ToString().Trim();
                    TA002 = row.Cells["單號"].Value.ToString().Trim();

                    SEARCHMOCTB3(TA001, TA002);

                }
                else
                {
                    SEARCHMOCTB3("", "");

                }
            }
        }

        public void SEARCHMOCTB3(string ta001, string TA002)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                SELECT  LTRIM(RTRIM(TB003)) AS '材料品號',LTRIM(RTRIM(TB012)) AS '材料品名',TB004 AS '需領用量'
                                FROM [TK].dbo.MOCTB
                                WHERE TB001='{0}' AND TB002='{1}'
                                ", TA001, TA002);

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView8.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView8.DataSource = ds3.Tables["ds3"];
                        dataGridView8.AutoResizeColumns();

                        dataGridView8.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView8.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        dataGridView8.Columns[0].Width = 120;
                        dataGridView8.Columns[1].Width = 100;
                        dataGridView8.Columns[2].Width = 100;
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
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMULTI();
                SEARCH();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }




        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH2(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMULTI2();
                SEARCH2(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SEARCH3(dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker7.Value.ToString("yyyyMMdd"));
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMULTI3();
                SEARCH3(dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker7.Value.ToString("yyyyMMdd"));

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }


        private void button7_Click(object sender, EventArgs e)
        {
            SEARCH4(dateTimePicker8.Value.ToString("yyyyMMdd"), dateTimePicker9.Value.ToString("yyyyMMdd"));
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMULTI4();
                SEARCH4(dateTimePicker8.Value.ToString("yyyyMMdd"), dateTimePicker9.Value.ToString("yyyyMMdd"));
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        #endregion

       
    }
}
