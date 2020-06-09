using System;
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

        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();

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
        }

        private void frmMOCCHANGE_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);
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

                sbSql.AppendFormat(@"  SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位',TA021 AS '線別',TA026 AS '訂單',TA027 AS '單號',TA028 AS '序號'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  WHERE TA003>='{0}' AND TA003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
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



        #endregion

       
    }
}
