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
    public partial class frmMOCOVEN : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlM = new StringBuilder();
        StringBuilder sbSqlDETAIL = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlDataAdapter adapterM = new SqlDataAdapter();
        SqlDataAdapter adapterDETAIL = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderM = new SqlCommandBuilder();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet dsMOCOVEN = new DataSet();
        DataSet dsMOCOVENDTAIL = new DataSet();
        DataTable dt = new DataTable();
        DataGridViewRow drEMPLOYEE = new DataGridViewRow();
        string tablename = null;
        string ID;
        int result;
        int rownum = 0;
        int rownumDETAIL = 0;
        Thread TD;
        DataGridViewRow drMOCOVEN = new DataGridViewRow();
        DataGridViewRow drMOCOVENDTAIL = new DataGridViewRow();

        public frmMOCOVEN()
        {
            InitializeComponent();
            combobox1load();
            combobox2load();
            combobox3load();
            combobox4load();
            combobox5load();
        }

        #region FUNCTION
        public void combobox1load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT   [ID] ,[DEPNAME]  FROM [TKMOC].[dbo].[MANUDEP] ";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("DEPNAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "DEPNAME";
            sqlConn.Close();

        }
        public void combobox2load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE]  WHERE ID IN ('100002','130036','140045','160114','970007','160130','160131','160132','160133','150063','160055','160057','160134','160138','040002') ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "ID";
            comboBox2.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void combobox3load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE]  WHERE ID IN ('100002','130036','140045','160114','970007','160130','160131','160132','160133','150063','160055','160057','160134','160138','040002') ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "ID";
            comboBox3.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void combobox4load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE ID IN ('100002','130036','140045','160114','970007','160130','160131','160132','160133','150063','160055','160057','160134','160138','040002') ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "ID";
            comboBox4.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void combobox5load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE ID IN ('100002','130036','140045','160114','970007','160130','160131','160132','160133','150063','160055','160057','160134','130138','040002') ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "ID";
            comboBox5.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void Search()
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                sqlConn.Open();
                
                sbSqlM.Clear();

                sbSqlM.AppendFormat(@" SELECT  CONVERT(varchar(100),[OVENDATE], 112) AS '日期',[MANUDEP].[DEPNAME] AS '組',CONVERT(varchar(100),[PREHEARTSTART], 108)  AS '預熱時間(起)',CONVERT(varchar(100),[PREHEARTEND], 108)   AS '預熱時間(迄)',[GAS]  AS '瓦斯磅數',EMP1.NAME  AS '折疊人員1',EMP2.NAME    AS '折疊人員2', EMP3.NAME   AS '主管',EMP4.NAME    AS '操作人員',");
                sbSqlM.AppendFormat(@" [MANUDEP] AS '組別',[MOCOVEN].[ID],[OVENDATE],[MANUDEP],[PREHEARTSTART],[PREHEARTEND],[GAS],[FLODPEOPLE1],[FLODPEOPLE2],[MANAGER],[OPERATOR]");
                sbSqlM.AppendFormat(@" FROM [TKMOC].[dbo].[MOCOVEN] WITH(NOLOCK)");
                sbSqlM.AppendFormat(@" LEFT JOIN [TKMOC].[dbo].[MANUEMPLOYEE] EMP1  ON [FLODPEOPLE1]=EMP1.ID");
                sbSqlM.AppendFormat(@" LEFT JOIN [TKMOC].[dbo].[MANUEMPLOYEE] EMP2 ON [FLODPEOPLE2]=EMP2.ID");
                sbSqlM.AppendFormat(@" LEFT JOIN [TKMOC].[dbo].[MANUEMPLOYEE]  EMP3 ON [MANAGER]=EMP3.ID");
                sbSqlM.AppendFormat(@" LEFT JOIN [TKMOC].[dbo].[MANUEMPLOYEE]  EMP4 ON [OPERATOR]=EMP4.ID");
                sbSqlM.AppendFormat(@" LEFT JOIN [TKMOC].[dbo].[MANUDEP] ON [MANUDEP].ID=[MOCOVEN].[MANUDEP]");
                sbSqlM.AppendFormat(@" WHERE  CONVERT(varchar(100),[OVENDATE], 112)='{0}'", dateTimePicker4.Value.ToString("yyyyMMdd"));
                sbSqlM.AppendFormat(@" ");

                adapterM = new SqlDataAdapter(@"" + sbSqlM, sqlConn);

                sqlCmdBuilderM = new SqlCommandBuilder(adapterM);
                
                dsMOCOVEN.Clear();
                adapterM.Fill(dsMOCOVEN, "TEMPds1");
                sqlConn.Close();


                if (dsMOCOVEN.Tables["TEMPds1"].Rows.Count == 0)
                {
                    //label1.Text = "找不到資料";
                    
                    SearchMOCOVENDTAIL(null);
                }
                else
                {
                    if (dsMOCOVEN.Tables["TEMPds1"].Rows.Count >= 1)
                    {                       
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = dsMOCOVEN.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        dataGridView1.CurrentCell = dataGridView1[0, rownum];


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
        public void SearchMOCOVENDTAIL(string ID)
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);


                sbSqlDETAIL.Clear();
                sbSqlDETAIL.AppendFormat(@" SELECT  [MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[TEMPER] AS '溫度',[HUMIDITY] AS '溼度',[WEATHER] AS '天氣',CONVERT(varchar(100),[MANUTIME],108) AS '時間'");
                sbSqlDETAIL.AppendFormat(@" ,[FURANACEUP1] AS '上爐-第一爐',[FURANACEUP2] AS '上爐-第二爐',[FURANACEUP3] AS '上爐-第三爐',[FURANACEUP4] AS '上爐-第四爐',[FURANACEUP5] AS '上爐-第五爐'");
                sbSqlDETAIL.AppendFormat(@" ,[FURANACEDOWN1] AS '下爐-第一爐',[FURANACEDOWN2] AS '下爐-第二爐',[FURANACEDOWN3] AS '下爐-第三爐',[FURANACEDOWN4] AS '下爐-第四爐',[FURANACEDOWN5] AS '下爐-第五爐'");
                sbSqlDETAIL.AppendFormat(@" ,[ID],[SOURCEID]");
                sbSqlDETAIL.AppendFormat(@" FROM [TKMOC].[dbo].[MOCOVENDTAIL]");
                sbSqlDETAIL.AppendFormat(@" WHERE [SOURCEID]='{0}'",ID.ToString());
                sbSqlDETAIL.AppendFormat(@" ");

                adapterDETAIL = new SqlDataAdapter(@" " + sbSqlDETAIL, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                dsMOCOVENDTAIL.Clear();
                dsMOCOVENDTAIL.Tables.Clear();
                adapterDETAIL.Fill(dsMOCOVENDTAIL, "TEMPds2");
                sqlConn.Close();


                if (dsMOCOVENDTAIL.Tables["TEMPds2"].Rows.Count == 0)
                {
                    
                    dataGridView2.DataSource = null;
                  
                }
                else
                {
                    if (dsMOCOVENDTAIL.Tables["TEMPds2"].Rows.Count >= 1)
                    {                        
                        dataGridView2.DataSource = dsMOCOVENDTAIL.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns();
                        dataGridView2.CurrentCell = dataGridView2[0, rownumDETAIL];


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

        public void SETADD()
        {
            dateTimePicker1.Enabled = true;
            dateTimePicker2.Enabled = true;
            dateTimePicker3.Enabled = true;
            textBox1.ReadOnly = false;
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            comboBox4.Enabled = true;
            comboBox5.Enabled = true;

            textBoxID.Text = null;
        }

        public void SETADDNEW()
        {
            dateTimePicker1.Value = dateTimePicker4.Value;
            textBoxID.Text = null;
            textBox1.Text = null;
            comboBox1.SelectedValue = "01";
            comboBox2.SelectedValue = "000002";
            comboBox3.SelectedValue = "000002";
            comboBox4.SelectedValue = "000002";
            comboBox5.SelectedValue = "000002";
        }

        public void SETUPDATE()
        {
            dateTimePicker1.Enabled = true;
            dateTimePicker2.Enabled = true;
            dateTimePicker3.Enabled = true;
            textBox1.ReadOnly = false;
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            comboBox4.Enabled = true;
            comboBox5.Enabled = true;
        }

        public void SETFINISH()
        {
            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;
            dateTimePicker3.Enabled = false;
            textBox1.ReadOnly = true;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            comboBox3.Enabled = false;
            comboBox4.Enabled = false;
            comboBox5.Enabled = false;
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
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MOCOVEN] ");
                sbSql.AppendFormat("  SET [OVENDATE]='{1}',[MANUDEP]='{2}',[PREHEARTSTART]='{3}',[PREHEARTEND]='{4}',[GAS]='{5}',[FLODPEOPLE1]='{6}',[FLODPEOPLE2]='{7}',[MANAGER]='{8}',[OPERATOR]='{9}'  WHERE [ID]='{0}' ", textBoxID.Text.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"), comboBox1.SelectedValue.ToString(), dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"), textBox1.Text.ToString(), comboBox2.SelectedValue.ToString(), comboBox3.SelectedValue.ToString(), comboBox4.SelectedValue.ToString(), comboBox5.SelectedValue.ToString());
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
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MOCOVEN]  ");
                sbSql.Append(" ([ID],[OVENDATE],[MANUDEP],[PREHEARTSTART],[PREHEARTEND],[GAS],[FLODPEOPLE1],[FLODPEOPLE2],[MANAGER],[OPERATOR])  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}') ", Guid.NewGuid(), dateTimePicker1.Value.ToString("yyyy/MM/dd"),comboBox1.SelectedValue.ToString(), dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"),textBox1.Text.ToString(),comboBox2.SelectedValue.ToString(), comboBox3.SelectedValue.ToString(), comboBox4.SelectedValue.ToString(), comboBox5.SelectedValue.ToString());

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
            if (dataGridView1.Rows.Count >= 1)
            {
                drMOCOVEN = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];

                textBoxID.Text = drMOCOVEN.Cells["ID"].Value.ToString();
                dateTimePicker1.Value = Convert.ToDateTime(drMOCOVEN.Cells["OVENDATE"].Value.ToString());
                comboBox1.SelectedValue = drMOCOVEN.Cells["MANUDEP"].Value.ToString();
                dateTimePicker2.Value = Convert.ToDateTime(drMOCOVEN.Cells["PREHEARTSTART"].Value.ToString());
                dateTimePicker3.Value = Convert.ToDateTime(drMOCOVEN.Cells["PREHEARTEND"].Value.ToString());
                textBox1.Text= drMOCOVEN.Cells["GAS"].Value.ToString();
                comboBox2.SelectedValue = drMOCOVEN.Cells["FLODPEOPLE1"].Value.ToString();
                comboBox3.SelectedValue = drMOCOVEN.Cells["FLODPEOPLE2"].Value.ToString();
                comboBox4.SelectedValue = drMOCOVEN.Cells["MANAGER"].Value.ToString();
                comboBox5.SelectedValue = drMOCOVEN.Cells["OPERATOR"].Value.ToString();


                SearchMOCOVENDTAIL(drMOCOVEN.Cells["ID"].Value.ToString());
                textBoxSID.Text = drMOCOVEN.Cells["ID"].Value.ToString();
                
            }
            else
            {
                SearchMOCOVENDTAIL("5C85A7F3-B942-4DF6-8804-5EC7ADF57870");
            }
        }

        public void SETDTEAILADD()
        {
            textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            textBox6.ReadOnly = false;
            textBox7.ReadOnly = false;
            textBox8.ReadOnly = false;
            textBox9.ReadOnly = false;
            textBox10.ReadOnly = false;
            textBox11.ReadOnly = false;
            textBox12.ReadOnly = false;
            textBox13.ReadOnly = false;
            textBox14.ReadOnly = false;
            textBox15.ReadOnly = false;
            textBox16.ReadOnly = false;
            comboBox6.Enabled = true;
            dateTimePicker5.Enabled = true;

           
        }
        public void SETADDDETAILNEW()
        {
            textBoxDETAILID.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;
            textBox14.Text = null;
            textBox15.Text = null;
            textBox16.Text = null;

        }
        public void SETDETAILUPDATE()
        {
            textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            textBox6.ReadOnly = false;
            textBox7.ReadOnly = false;
            textBox8.ReadOnly = false;
            textBox9.ReadOnly = false;
            textBox10.ReadOnly = false;
            textBox11.ReadOnly = false;
            textBox12.ReadOnly = false;
            textBox13.ReadOnly = false;
            textBox14.ReadOnly = false;
            textBox15.ReadOnly = false;
            textBox16.ReadOnly = false;
            comboBox6.Enabled = true;
            dateTimePicker5.Enabled = true;
        }
        public void SETDETAILFINISH()
        {
            textBox2.ReadOnly = true;
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
            textBox5.ReadOnly = true;
            textBox6.ReadOnly = true;
            textBox7.ReadOnly = true;
            textBox8.ReadOnly = true;
            textBox9.ReadOnly = true;
            textBox10.ReadOnly = true;
            textBox11.ReadOnly = true;
            textBox12.ReadOnly = true;
            textBox13.ReadOnly = true;
            textBox14.ReadOnly = true;
            textBox15.ReadOnly = true;
            textBox16.ReadOnly = true;
            comboBox6.Enabled = false;
            dateTimePicker5.Enabled = false;
        }
        public void DETAILUPDATE()
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
                sbSql.AppendFormat("  UPDATE [TKMOC].[dbo].[MOCOVENDTAIL]");
                sbSql.AppendFormat("  SET [MB001]='{1}',[MB002]='{2}',[MB003]='{3}',[TEMPER]='{4}',[HUMIDITY]='{5}',[WEATHER]='{6}',[MANUTIME]='{7}' ,[FURANACEUP1]='{8}',[FURANACEUP2]='{9}',[FURANACEUP3]='{10}',[FURANACEUP4]='{11}',[FURANACEUP5]='{12}' ,[FURANACEDOWN1]='{13}',[FURANACEDOWN2]='{14}',[FURANACEDOWN3]='{15}',[FURANACEDOWN4]='{16}',[FURANACEDOWN5]='{17}' WHERE [ID]='{0}'", textBoxDETAILID.Text.ToString(),textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), comboBox6.Text.ToString(), dateTimePicker5.Value.ToString("HH:mm"), textBox7.Text.ToString(), textBox8.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString(), textBox14.Text.ToString(), textBox15.Text.ToString(), textBox16.Text.ToString());
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
        public void DETAILADD()
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
                sbSql.AppendFormat("  INSERT INTO [TKMOC].[dbo].[MOCOVENDTAIL]");
                sbSql.AppendFormat("  ([ID],[SOURCEID],[MB001],[MB002],[MB003],[TEMPER],[HUMIDITY],[WEATHER],[MANUTIME]");
                sbSql.AppendFormat("  ,[FURANACEUP1],[FURANACEUP2],[FURANACEUP3],[FURANACEUP4],[FURANACEUP5]");
                sbSql.AppendFormat("  ,[FURANACEDOWN1],[FURANACEDOWN2],[FURANACEDOWN3],[FURANACEDOWN4],[FURANACEDOWN5])");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}') ", Guid.NewGuid(),textBoxSID.Text.ToString(),textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(),comboBox6.Text.ToString(), dateTimePicker5.Value.ToString("HH:mm"), textBox7.Text.ToString(), textBox8.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString(), textBox14.Text.ToString(), textBox15.Text.ToString(), textBox16.Text.ToString());
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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox3.Text = FINDMB002(textBox2.Text.ToString());
            textBox4.Text = FINDMB003(textBox2.Text.ToString());
        }
        public string FINDMB002(string MB001)
        {
            DataSet ds = new DataSet();
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            sbSql.Clear();
            sbSqlQuery.Clear();

            sbSqlQuery.AppendFormat(@" SELECT [MB001],[MB002],[MB003] FROM [TKMOC].[dbo].[ERPINVMB] WHERE [MB001]='{0}'", MB001.ToString());

            adapter = new SqlDataAdapter(@"" + sbSqlQuery, sqlConn);
            sqlCmdBuilder = new SqlCommandBuilder(adapter);

            sqlConn.Open();
            ds.Clear();
            adapter.Fill(ds, "TEMPds1");
            sqlConn.Close();


            if (ds.Tables["TEMPds1"].Rows.Count >= 1)
            {
                return ds.Tables["TEMPds1"].Rows[0]["MB002"].ToString();
            }
            else
            {
                return "";
            }


        }
        public string FINDMB003(string MB001)
        {
            DataSet ds = new DataSet();
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            sbSql.Clear();
            sbSqlQuery.Clear();

            sbSql.AppendFormat(@" SELECT [MB001],[MB002],[MB003] FROM [TKMOC].[dbo].[ERPINVMB] WHERE [MB001]='{0}'", MB001.ToString());

            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
            sqlCmdBuilder = new SqlCommandBuilder(adapter);

            sqlConn.Open();
            ds.Clear();
            adapter.Fill(ds, "TEMPds1");
            sqlConn.Close();


            if (ds.Tables["TEMPds1"].Rows.Count >= 1)
            {
                return ds.Tables["TEMPds1"].Rows[0]["MB003"].ToString();
            }
            else
            {
                return "";
            }


        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dsMOCOVENDTAIL.Tables["TEMPds2"].Rows.Count>=1)
            {
                drMOCOVENDTAIL = dataGridView2.Rows[dataGridView2.SelectedCells[0].RowIndex];

                textBox2.Text = drMOCOVENDTAIL.Cells["品號"].Value.ToString();
                textBox3.Text = drMOCOVENDTAIL.Cells["品名"].Value.ToString();
                textBox4.Text = drMOCOVENDTAIL.Cells["規格"].Value.ToString();
                textBox5.Text = drMOCOVENDTAIL.Cells["溫度"].Value.ToString();
                textBox6.Text = drMOCOVENDTAIL.Cells["溼度"].Value.ToString();
                textBox7.Text = drMOCOVENDTAIL.Cells["上爐-第一爐"].Value.ToString();
                textBox8.Text = drMOCOVENDTAIL.Cells["上爐-第二爐"].Value.ToString();
                textBox9.Text = drMOCOVENDTAIL.Cells["上爐-第三爐"].Value.ToString();
                textBox10.Text = drMOCOVENDTAIL.Cells["上爐-第四爐"].Value.ToString();
                textBox11.Text = drMOCOVENDTAIL.Cells["上爐-第五爐"].Value.ToString();
                textBox12.Text = drMOCOVENDTAIL.Cells["下爐-第一爐"].Value.ToString();
                textBox13.Text = drMOCOVENDTAIL.Cells["下爐-第二爐"].Value.ToString();
                textBox14.Text = drMOCOVENDTAIL.Cells["下爐-第三爐"].Value.ToString();
                textBox15.Text = drMOCOVENDTAIL.Cells["下爐-第四爐"].Value.ToString();
                textBox16.Text = drMOCOVENDTAIL.Cells["下爐-第五爐"].Value.ToString();
                comboBox6.Text = drMOCOVENDTAIL.Cells["天氣"].Value.ToString();
                dateTimePicker5.Value = Convert.ToDateTime(drMOCOVENDTAIL.Cells["時間"].Value.ToString());

                textBoxDETAILID.Text = drMOCOVENDTAIL.Cells["ID"].Value.ToString();
                textBoxSID.Text = drMOCOVENDTAIL.Cells["SOURCEID"].Value.ToString();
            }
            else
            {
                SETADDDETAILNEW();
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion

        #region BUTTOON
        private void button1_Click(object sender, EventArgs e)
        {
            SETADD();
            SETADDNEW();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETUPDATE();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBoxID.Text.ToString()))
            {
                UPDATE();
            }
            else
            {
                ADD();
            }
            if (dsMOCOVEN.Tables["TEMPds1"].Rows.Count >= 1)
            {
                rownum = dataGridView1.CurrentCell.RowIndex;
            }
            
            Search();
            SETFINISH();
            
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SETDTEAILADD();
            SETADDDETAILNEW();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SETDETAILUPDATE();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBoxDETAILID.Text.ToString()))
            {
                DETAILUPDATE();
            }
            else
            {
                DETAILADD();
            }
            if (dsMOCOVEN.Tables["TEMPds1"].Rows.Count >= 1)
            {
                rownum = dataGridView1.CurrentCell.RowIndex;
            }
            if (dsMOCOVENDTAIL.Tables["TEMPds2"].Rows.Count >= 1)
            {
                rownumDETAIL = dataGridView2.CurrentCell.RowIndex;
            }
                       
            Search();
            SearchMOCOVENDTAIL(textBoxID.Text.ToString());
            SETDETAILFINISH();
        }


        #endregion


    }
}
