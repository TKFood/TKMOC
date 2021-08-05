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
using FastReport;
using FastReport.Data;

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

            textBox3.Text = DateTime.Now.Year.ToString();
        }

        #region FUNCTION
        private void frmTRACEBACK_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色
            dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView3.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView4.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView5.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 40;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);


            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView1
            dataGridView1.Controls.Add(cbHeader);
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);


            //先建立個 CheckBox 欄
            cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 40;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView2.Columns.Insert(0, cbCol);


            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            rect = dataGridView2.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;
            //将 CheckBox 加入到 dataGridView2
            dataGridView2.Controls.Add(cbHeader);
            dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);



            //先建立個 CheckBox 欄
            cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 40;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView3.Columns.Insert(0, cbCol);

  

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            rect = dataGridView3.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;
            //将 CheckBox 加入到 dataGridView2
            dataGridView3.Controls.Add(cbHeader);
            dataGridView3.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);


            //先建立個 CheckBox 欄
            cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 40;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView4.Columns.Insert(0, cbCol);



            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            rect = dataGridView4.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;
            //将 CheckBox 加入到 dataGridView2
            dataGridView4.Controls.Add(cbHeader);
            dataGridView4.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);


            //先建立個 CheckBox 欄
            cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 40;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView5.Columns.Insert(0, cbCol);



            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            rect = dataGridView5.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;
            //将 CheckBox 加入到 dataGridView2
            dataGridView5.Controls.Add(cbHeader);
            dataGridView5.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);



        }
        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }
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

                sbSql.AppendFormat(@"  
                                    SELECT MF001,MF002,'0',MF003,MF004,MF005,MF006,MF010
                                    FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK)
                                    WHERE MF001=ME001 AND MF002=ME002
                                    AND MF009 IN ('1','2','5')
                                    AND MF001='{0}' AND MF002='{1}'
                                    ORDER BY MF002,MF003,MF004,MF005
                                    ", MB001, LOTNO);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    ADDTRACEBACKOUT(MB001, LOTNO);
                    ADDTRACEBACKMOC(MB001, LOTNO);
                    ADDTRACEBACKMOCOUTIN(MB001, LOTNO);
                    ADDTRACEBACKINVMF(MB001, LOTNO);
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

        public void SEARCHOUT2(string MB001, string LOTNO)
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

                sbSql.AppendFormat(@" 
                                SELECT MF001,MF002,'1入庫','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND MQ003 IN ('34','58')
                                AND MF001='{0}' AND MF002='{1}'
                                ORDER BY INVMF.MF002,MF003,MF004,MF005
                                    ", MB001, LOTNO);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    ADDTRACEBACKOUT2(MB001, LOTNO);
                    ADDTRACEBACKMOC2(MB001, LOTNO);
                    ADDTRACEBACKMOCOUTIN2(MB001, LOTNO);
                    ADDTRACEBACKINVMFSALE2(MB001, LOTNO);
                    ADDTRACEBACKINVMF2(MB001, LOTNO);
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

        public void DELETETRACEBACK()
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@" 
                                     DELETE [TKMOC].[dbo].[TRACEBACK]     
                                     ");

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

         
                sbSql.AppendFormat(@" 
 
                                     INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                     ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])
                                     SELECT MF001,MF002,'1銷貨','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010
                                     FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK)
                                     WHERE MF001=ME001 AND MF002=ME002
                                     AND MF009 IN ('2','5')
                                     AND MF001='{0}' AND MF002='{1}'
                                     ORDER BY INVMF.MF002,MF003,MF004,MF005

                                     ", MB001, LOTNO);

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

        public void ADDTRACEBACKOUT2(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
               
                   
                                    INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])
                                    SELECT MF001,MF002,'1入庫','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010
                                    FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK)
                                    WHERE MF001=ME001 AND MF002=ME002
                                    AND MQ001=MF004
                                    AND MQ003='34'
                                    AND MF001='{0}' AND MF002='{1}'
                                    ORDER BY INVMF.MF002,MF003,MF004,MF005
                    ", MB001,LOTNO);


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

        public void ADDTRACEBACKMOC(string MB001, string LOTNO)
        {
            int LEVELNOW = 0;
            int LEVELNEXT = 1;
            int MAXCOUNT = 1;
            int DSCEHCK = 1;


            //新增成品的LEVEL=0
            ADDTRACEBACKPRODUCTLEVEL0(MB001, LOTNO);

            while (DSCEHCK >= 1 && MAXCOUNT <= 10)
            {
                ADDTRACEBACKLEVELPRODUCTNEXT(MB001, LOTNO,LEVELNOW.ToString(), LEVELNEXT.ToString());

                LEVELNOW = LEVELNOW + 1;
                LEVELNEXT = LEVELNEXT + 1;
                MAXCOUNT = MAXCOUNT + 1;

                DSCEHCK = CHECKPRODUCTLEVEL(MB001, LOTNO, LEVELNOW.ToString());


            }

            //try
            //{
            //    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //    sqlConn = new SqlConnection(connectionString);

            //    sqlConn.Close();
            //    sqlConn.Open();
            //    tran = sqlConn.BeginTransaction();

            //    sbSql.Clear();

            //    sbSql.AppendFormat(@" 
            //                         WITH RTABLES AS 
            //                         ( SELECT 0 AS LEVELS,[TG001],[TG002],[TG003],[TG004],[TG011],[TG017],[TG014],[TG015],[TE001],[TE002],[TE003],[TE004],[TE005],[TE010] 
            //                         FROM [TK].[dbo].[VMOCTGMOCTE] WITH (NOLOCK) 
            //                         WHERE [VMOCTGMOCTE].TG004 ='{0}' AND [VMOCTGMOCTE].TG017 ='{1}'  
            //                         UNION ALL 
            //                         SELECT LEVELS+1,B.[TG001], B.[TG002], B.[TG003], B.[TG004], B.[TG011], B.[TG017], B.[TG014], B.[TG015], B.[TE001], B.[TE002],B.[TE003], B.[TE004], B.[TE005], B.[TE010] 
            //                         FROM [TK].[dbo].[VMOCTGMOCTE] B WITH (NOLOCK) 
            //                         INNER JOIN RTABLES ON RTABLES.[TE004]=B.[TG004] AND RTABLES.[TE010]=B.[TG017] )   


            //                        INSERT INTO [TKMOC].[dbo].[TRACEBACK] 
            //                        ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS],[TG014],[TG015]) 

            //                         SELECT '{0}','{1}','2生產',LEVELS  
            //                         ,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF WHERE TF001=TG001 AND TF002=TG002 ORDER BY TF003) 
            //                         ,[TG001],[TG002],[TG003],[TG004], '',[TG017],[TG011]  ,[TG014],[TG015]
            //                         FROM RTABLES 
            //                         GROUP BY LEVELS,[TG001],[TG002],[TG003],[TG004],[TG017],[TG011] ,[TG014],[TG015]
            //                         ORDER BY LEVELS,[TG004] 

            //                        ", MB001, LOTNO);



            //    cmd.Connection = sqlConn;
            //    cmd.CommandTimeout = 60;
            //    cmd.CommandText = sbSql.ToString();
            //    cmd.Transaction = tran;
            //    result = cmd.ExecuteNonQuery();

            //    if (result == 0)
            //    {
            //        tran.Rollback();    //交易取消
            //    }
            //    else
            //    {
            //        tran.Commit();      //執行交易  


            //    }

            //}
            //catch
            //{

            //}

            //finally
            //{
            //    sqlConn.Close();
            //}
        }

        public void ADDTRACEBACKPRODUCTLEVEL0(string MB001, string LOTNO)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                     
                                   INSERT INTO [TKMOC].[dbo].[TRACEBACK] 
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS],[TG014],[TG015]) 

                                    SELECT '{0}','{1}','2生產','0'  
                                    ,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF WHERE TF001=TG001 AND TF002=TG002 ORDER BY TF003) 
                                    ,[TG001],[TG002],'****' [TG003],[TG004], '',[TG017],SUM([TG011]) [TG011]  ,[TG014],[TG015]
                                    FROM [TK].dbo.MOCTE,[TK].dbo.MOCTG
                                    WHERE  TG014 = TE011 AND TG015 = TE012
                                    AND TG004='{0}' AND TG017='{1}'
                                    GROUP BY [TG001],[TG002],[TG004],[TG017] ,[TG014],[TG015]
                                    ORDER BY [TG004] 


                                    ", MB001, LOTNO);

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


        public void ADDTRACEBACKLEVELPRODUCTNEXT(string MB001,string LOTNO, string LEVELNOW, string LEVELNEXT)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"    
                                    INSERT INTO [TKMOC].[dbo].[TRACEBACK] 
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS],[TG014],[TG015]) 
                                 
                                    SELECT 
                                    '{0}','{1}','2生產','{2}'
                                    ,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF WHERE TF001 = TG001 AND TF002 = TG002 ORDER BY TF003)
                                    ,[TG001],[TG002],'****' [TG003],[TG004], '',[TG017],SUM([TG011]) [TG011]  ,[TG014],[TG015]
                                    FROM [TK].dbo.MOCTE
	                                    ,[TK].dbo.MOCTG
                                    WHERE TG014 = TE011
	                                    AND TG015 = TE012
	                                    AND LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017)) IN (
		                                    SELECT LTRIM(RTRIM([TE004])) + LTRIM(RTRIM([TE010]))
		                                    FROM [TK].dbo.MOCTE
			                                    ,[TK].dbo.MOCTG
		                                    WHERE TG014 = TE011
			                                    AND TG015 = TE012
			                                    AND LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017)) IN (
				                                    SELECT LTRIM(RTRIM(MB001)) + LTRIM(RTRIM(LOTNO))
				                                    FROM [TKMOC].[dbo].[TRACEBACK]
				                                    WHERE LEVELS = '{3}'
				                                    )
		                                    GROUP BY LTRIM(RTRIM([TE004])) + LTRIM(RTRIM([TE010]))
		                                    )
                                    GROUP BY [TG001],[TG002],[TG004],[TG017] ,[TG014],[TG015]
                                    ORDER BY [TG004]


                                    ", MB001, LOTNO, LEVELNEXT, LEVELNOW);

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

        public int CHECKPRODUCTLEVEL(string MB001, string LOTNO, string LEVELS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                   SELECT 
                                    '{0}','{1}','2生產','1'
                                    ,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF WHERE TF001 = TG001 AND TF002 = TG002 ORDER BY TF003)
                                    ,[TG001],[TG002],[TG003],[TG004],'',[TG017],[TG011],[TG014],[TG015]
                                    FROM [TK].dbo.MOCTE
	                                    ,[TK].dbo.MOCTG
                                    WHERE TG014 = TE011
	                                    AND TG015 = TE012
	                                    AND LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017)) IN (
		                                    SELECT LTRIM(RTRIM([TE004])) + LTRIM(RTRIM([TE010]))
		                                    FROM [TK].dbo.MOCTE
			                                    ,[TK].dbo.MOCTG
		                                    WHERE TG014 = TE011
			                                    AND TG015 = TE012
			                                    AND LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017)) IN (
				                                    SELECT LTRIM(RTRIM(MB001)) + LTRIM(RTRIM(LOTNO))
				                                    FROM [TKMOC].[dbo].[TRACEBACKTEMP]
				                                    WHERE LEVELS = '{2}'
				                                    )
		                                    GROUP BY LTRIM(RTRIM([TE004])) + LTRIM(RTRIM([TE010]))
		                                    )
                                    GROUP BY [TG001],[TG002],[TG003],[TG004],[TG017],[TG011],[TG014],[TG015]
                                    ORDER BY [TG004]
                                    ", MB001, LOTNO, LEVELS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {

                    return ds.Tables["ds"].Rows.Count;
                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDTRACEBACKMOC2(string MB001, string LOTNO)
        {
            int LEVELNOW = 0;
            int LEVELNEXT = 1;
            int MAXCOUNT = 1;
            int DSCEHCK = 1;

            //新增LEVEL=0
            ADDTRACEBACKLEVEL0(MB001.Trim(), LOTNO.Trim());

            while(DSCEHCK>=1 && MAXCOUNT<=10)
            {
                ADDTRACEBACKLEVELNEXT(LEVELNOW.ToString(), LEVELNEXT.ToString());

                LEVELNOW = LEVELNOW + 1;
                LEVELNEXT = LEVELNEXT + 1;
                MAXCOUNT = MAXCOUNT + 1;

                DSCEHCK = CHECKLEVEL(LEVELNOW);


            }

            //try
            //{
            //    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //    sqlConn = new SqlConnection(connectionString);

            //    sqlConn.Close();
            //    sqlConn.Open();
            //    tran = sqlConn.BeginTransaction();

            //    sbSql.Clear();

            //    sbSql.AppendFormat(@"                                     
            //                    INSERT INTO [TKMOC].[dbo].[TRACEBACK]
            //                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[TG014],[TG015],[MB001],[MB002],[LOTNO],[NUMS])

            //                    SELECT DISTINCT MF001
            //                     ,MF002
            //                     ,'2領退料'
            //                     ,'0'
            //                     ,MF003
            //                     ,MF004
            //                     ,MF005
            //                     ,MF006
            //                     ,TG014
            //                     ,TG015
            //                     ,TG004
            //                     ,''
            //                     ,TG017
            //                     ,TG011
            //                    FROM [TK].dbo.INVME WITH (NOLOCK)
            //                     ,[TK].dbo.INVMF WITH (NOLOCK)
            //                     ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                     ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                     ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                    WHERE MF001 = ME001
            //                     AND MF002 = ME002
            //                     AND MQ001 = MF004
            //                     AND TE001 = MF004
            //                     AND TE002 = MF005
            //                     AND TE004 = MF001
            //                     AND TE010 = MF002
            //                     AND TG014 = TE011
            //                     AND TG015 = TE012
            //                     AND MQ003 IN (
            //                      '54'
            //                      ,'56'
            //                      )
            //                     AND MF001 = '{0}'
            //                     AND MF002 = '{1}'

            //                    UNION ALL

            //                    SELECT DISTINCT MF001
            //                     ,MF002
            //                     ,'2領退料'
            //                     ,'1'
            //                     ,MF003
            //                     ,MF004
            //                     ,MF005
            //                     ,MF006
            //                     ,TG014
            //                     ,TG015
            //                     ,TG004
            //                     ,''
            //                     ,TG017
            //                     ,TG011
            //                    FROM [TK].dbo.INVME WITH (NOLOCK)
            //                     ,[TK].dbo.INVMF WITH (NOLOCK)
            //                     ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                     ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                     ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                    WHERE MF001 = ME001
            //                     AND MF002 = ME002
            //                     AND MQ001 = MF004
            //                     AND TE001 = MF004
            //                     AND TE002 = MF005
            //                     AND TE004 = MF001
            //                     AND TE010 = MF002
            //                     AND TG014 = TE011
            //                     AND TG015 = TE012
            //                     AND MQ003 IN (
            //                      '54'
            //                      ,'56'
            //                      )
            //                     AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN (
            //                      SELECT LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                      FROM [TK].dbo.INVME WITH (NOLOCK)
            //                       ,[TK].dbo.INVMF WITH (NOLOCK)
            //                       ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                       ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                       ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                      WHERE MF001 = ME001
            //                       AND MF002 = ME002
            //                       AND MQ001 = MF004
            //                       AND TE001 = MF004
            //                       AND TE002 = MF005
            //                       AND TE004 = MF001
            //                       AND TE010 = MF002
            //                       AND TG014 = TE011
            //                       AND TG015 = TE012
            //                       AND MQ003 IN (
            //                        '54'
            //                        ,'56'
            //                        )
            //                       AND MF001 = '{0}'
            //                       AND MF002 = '{1}'
            //                      GROUP BY LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                      )

            //                    UNION ALL

            //                    SELECT DISTINCT MF001
            //                     ,MF002
            //                     ,'2領退料'
            //                     ,'2'
            //                     ,MF003
            //                     ,MF004
            //                     ,MF005
            //                     ,MF006
            //                     ,TG014
            //                     ,TG015
            //                     ,TG004
            //                     ,''
            //                     ,TG017
            //                     ,TG011
            //                    FROM [TK].dbo.INVME WITH (NOLOCK)
            //                     ,[TK].dbo.INVMF WITH (NOLOCK)
            //                     ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                     ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                     ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                    WHERE MF001 = ME001
            //                     AND MF002 = ME002
            //                     AND MQ001 = MF004
            //                     AND TE001 = MF004
            //                     AND TE002 = MF005
            //                     AND TE004 = MF001
            //                     AND TE010 = MF002
            //                     AND TG014 = TE011
            //                     AND TG015 = TE012
            //                     AND MQ003 IN (
            //                      '54'
            //                      ,'56'
            //                      )
            //                     AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN (
            //                      SELECT LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                      FROM [TK].dbo.INVME WITH (NOLOCK)
            //                       ,[TK].dbo.INVMF WITH (NOLOCK)
            //                       ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                       ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                       ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                      WHERE MF001 = ME001
            //                       AND MF002 = ME002
            //                       AND MQ001 = MF004
            //                       AND TE001 = MF004
            //                       AND TE002 = MF005
            //                       AND TE004 = MF001
            //                       AND TE010 = MF002
            //                       AND TG014 = TE011
            //                       AND TG015 = TE012
            //                       AND MQ003 IN (
            //                        '54'
            //                        ,'56'
            //                        )
            //                       AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN (
            //                        SELECT LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                        FROM [TK].dbo.INVME WITH (NOLOCK)
            //                         ,[TK].dbo.INVMF WITH (NOLOCK)
            //                         ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                         ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                         ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                        WHERE MF001 = ME001
            //                         AND MF002 = ME002
            //                         AND MQ001 = MF004
            //                         AND TE001 = MF004
            //                         AND TE002 = MF005
            //                         AND TE004 = MF001
            //                         AND TE010 = MF002
            //                         AND TG014 = TE011
            //                         AND TG015 = TE012
            //                         AND MQ003 IN (
            //                          '54'
            //                          ,'56'
            //                          )
            //                         AND MF001 = '{0}'
            //                         AND MF002 = '{1}'
            //                        GROUP BY LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                        )
            //                      GROUP BY LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                      )

            //                    UNION ALL

            //                    SELECT DISTINCT MF001
            //                     ,MF002
            //                     ,'2領退料'
            //                     ,'3'
            //                     ,MF003
            //                     ,MF004
            //                     ,MF005
            //                     ,MF006
            //                     ,TG014
            //                     ,TG015
            //                     ,TG004
            //                     ,''
            //                     ,TG017
            //                     ,TG011
            //                    FROM [TK].dbo.INVME WITH (NOLOCK)
            //                     ,[TK].dbo.INVMF WITH (NOLOCK)
            //                     ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                     ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                     ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                    WHERE MF001 = ME001
            //                     AND MF002 = ME002
            //                     AND MQ001 = MF004
            //                     AND TE001 = MF004
            //                     AND TE002 = MF005
            //                     AND TE004 = MF001
            //                     AND TE010 = MF002
            //                     AND TG014 = TE011
            //                     AND TG015 = TE012
            //                     AND MQ003 IN (
            //                      '54'
            //                      ,'56'
            //                      )
            //                     AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN (
            //                      SELECT LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                      FROM [TK].dbo.INVME WITH (NOLOCK)
            //                       ,[TK].dbo.INVMF WITH (NOLOCK)
            //                       ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                       ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                       ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                      WHERE MF001 = ME001
            //                       AND MF002 = ME002
            //                       AND MQ001 = MF004
            //                       AND TE001 = MF004
            //                       AND TE002 = MF005
            //                       AND TE004 = MF001
            //                       AND TE010 = MF002
            //                       AND TG014 = TE011
            //                       AND TG015 = TE012
            //                       AND MQ003 IN (
            //                        '54'
            //                        ,'56'
            //                        )
            //                       AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN (
            //                        SELECT LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                        FROM [TK].dbo.INVME WITH (NOLOCK)
            //                         ,[TK].dbo.INVMF WITH (NOLOCK)
            //                         ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                         ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                         ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                        WHERE MF001 = ME001
            //                         AND MF002 = ME002
            //                         AND MQ001 = MF004
            //                         AND TE001 = MF004
            //                         AND TE002 = MF005
            //                         AND TE004 = MF001
            //                         AND TE010 = MF002
            //                         AND TG014 = TE011
            //                         AND TG015 = TE012
            //                         AND MQ003 IN (
            //                          '54'
            //                          ,'56'
            //                          )
            //                         AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN (
            //                          SELECT LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                          FROM [TK].dbo.INVME WITH (NOLOCK)
            //                           ,[TK].dbo.INVMF WITH (NOLOCK)
            //                           ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                           ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                           ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                          WHERE MF001 = ME001
            //                           AND MF002 = ME002
            //                           AND MQ001 = MF004
            //                           AND TE001 = MF004
            //                           AND TE002 = MF005
            //                           AND TE004 = MF001
            //                           AND TE010 = MF002
            //                           AND TG014 = TE011
            //                           AND TG015 = TE012
            //                           AND MQ003 IN (
            //                            '54'
            //                            ,'56'
            //                            )
            //                           AND MF001 = '{0}'
            //                           AND MF002 = '{1}'
            //                          GROUP BY LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                          )
            //                        GROUP BY LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                        )
            //                      GROUP BY LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                      )

            //                    UNION ALL

            //                    SELECT DISTINCT MF001
            //                     ,MF002
            //                     ,'2領退料'
            //                     ,'4'
            //                     ,MF003
            //                     ,MF004
            //                     ,MF005
            //                     ,MF006
            //                     ,TG014
            //                     ,TG015
            //                     ,TG004
            //                     ,''
            //                     ,TG017
            //                     ,TG011
            //                    FROM [TK].dbo.INVME WITH (NOLOCK)
            //                     ,[TK].dbo.INVMF WITH (NOLOCK)
            //                     ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                     ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                     ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                    WHERE MF001 = ME001
            //                     AND MF002 = ME002
            //                     AND MQ001 = MF004
            //                     AND TE001 = MF004
            //                     AND TE002 = MF005
            //                     AND TE004 = MF001
            //                     AND TE010 = MF002
            //                     AND TG014 = TE011
            //                     AND TG015 = TE012
            //                     AND MQ003 IN (
            //                      '54'
            //                      ,'56'
            //                      )
            //                     AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN (
            //                      SELECT LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                      FROM [TK].dbo.INVME WITH (NOLOCK)
            //                       ,[TK].dbo.INVMF WITH (NOLOCK)
            //                       ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                       ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                       ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                      WHERE MF001 = ME001
            //                       AND MF002 = ME002
            //                       AND MQ001 = MF004
            //                       AND TE001 = MF004
            //                       AND TE002 = MF005
            //                       AND TE004 = MF001
            //                       AND TE010 = MF002
            //                       AND TG014 = TE011
            //                       AND TG015 = TE012
            //                       AND MQ003 IN (
            //                        '54'
            //                        ,'56'
            //                        )
            //                       AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN (
            //                        SELECT LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                        FROM [TK].dbo.INVME WITH (NOLOCK)
            //                         ,[TK].dbo.INVMF WITH (NOLOCK)
            //                         ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                         ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                         ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                        WHERE MF001 = ME001
            //                         AND MF002 = ME002
            //                         AND MQ001 = MF004
            //                         AND TE001 = MF004
            //                         AND TE002 = MF005
            //                         AND TE004 = MF001
            //                         AND TE010 = MF002
            //                         AND TG014 = TE011
            //                         AND TG015 = TE012
            //                         AND MQ003 IN (
            //                          '54'
            //                          ,'56'
            //                          )
            //                         AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN (
            //                          SELECT LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                          FROM [TK].dbo.INVME WITH (NOLOCK)
            //                           ,[TK].dbo.INVMF WITH (NOLOCK)
            //                           ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                           ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                           ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                          WHERE MF001 = ME001
            //                           AND MF002 = ME002
            //                           AND MQ001 = MF004
            //                           AND TE001 = MF004
            //                           AND TE002 = MF005
            //                           AND TE004 = MF001
            //                           AND TE010 = MF002
            //                           AND TG014 = TE011
            //                           AND TG015 = TE012
            //                           AND MQ003 IN (
            //                            '54'
            //                            ,'56'
            //                            )
            //                           AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN (
            //                            SELECT LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                            FROM [TK].dbo.INVME WITH (NOLOCK)
            //                             ,[TK].dbo.INVMF WITH (NOLOCK)
            //                             ,[TK].dbo.CMSMQ WITH (NOLOCK)
            //                             ,[TK].dbo.MOCTE WITH (NOLOCK)
            //                             ,[TK].dbo.MOCTG WITH (NOLOCK)
            //                            WHERE MF001 = ME001
            //                             AND MF002 = ME002
            //                             AND MQ001 = MF004
            //                             AND TE001 = MF004
            //                             AND TE002 = MF005
            //                             AND TE004 = MF001
            //                             AND TE010 = MF002
            //                             AND TG014 = TE011
            //                             AND TG015 = TE012
            //                             AND MQ003 IN (
            //                              '54'
            //                              ,'56'
            //                              )
            //                             AND MF001 = '{0}'
            //                             AND MF002 = '{1}'
            //                            GROUP BY LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                            )
            //                          GROUP BY LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                          )
            //                        GROUP BY LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                        )
            //                      GROUP BY LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017))
            //                      )
            //                    ORDER BY INVMF.MF001
            //                     ,INVMF.MF002
            //                     ,MF004
            //                     ,MF005

            //                        ", MB001,LOTNO);

            //    cmd.Connection = sqlConn;
            //    cmd.CommandTimeout = 60;
            //    cmd.CommandText = sbSql.ToString();
            //    cmd.Transaction = tran;
            //    result = cmd.ExecuteNonQuery();

            //    if (result == 0)
            //    {
            //        tran.Rollback();    //交易取消
            //    }
            //    else
            //    {
            //        tran.Commit();      //執行交易  


            //    }

            //}
            //catch
            //{

            //}

            //finally
            //{
            //    sqlConn.Close();
            //}
        }

        public void ADDTRACEBACKLEVEL0(string MB001, string LOTNO)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                     
                                    INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[TG014],[TG015],[MB001],[MB002],[LOTNO],[NUMS])

                                    SELECT MF001
                                     ,MF002
                                     ,'2領退料'
                                     ,'0'
                                     ,MF003
                                     ,MF004
                                     ,MF005
                                     ,MF006
                                     ,TG014
                                     ,TG015
                                     ,TG004
                                     ,''
                                     ,TG017
                                     ,SUM(TG011)
                                    FROM [TK].dbo.INVME WITH (NOLOCK)
                                     ,[TK].dbo.INVMF WITH (NOLOCK)
                                     ,[TK].dbo.CMSMQ WITH (NOLOCK)
                                     ,[TK].dbo.MOCTE WITH (NOLOCK)
                                     ,[TK].dbo.MOCTG WITH (NOLOCK)
                                    WHERE MF001 = ME001
                                     AND MF002 = ME002
                                     AND MQ001 = MF004
                                     AND TE001 = MF004
                                     AND TE002 = MF005
                                     AND TE004 = MF001
                                     AND TE010 = MF002
                                     AND TG014 = TE011
                                     AND TG015 = TE012
                                     AND MQ003 IN (
                                      '54'
                                      ,'56'
                                      )
                                     AND MF001 = '{0}'
                                     AND MF002 = '{1}'
                                        GROUP BY  MF001
                                        ,MF002
                                        ,MF003
                                        ,MF004
                                        ,MF005
                                        ,MF006
                                        ,TG014
                                        ,TG015
                                        ,TG004
                                        ,TG017

                                    ", MB001, LOTNO);

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

        public int CHECKLEVEL(int LEVELS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT *
                                    FROM [TK].dbo.INVME WITH (NOLOCK)
                                    ,[TK].dbo.INVMF WITH (NOLOCK)
                                    ,[TK].dbo.CMSMQ WITH (NOLOCK)
                                    ,[TK].dbo.MOCTE WITH (NOLOCK)
                                    ,[TK].dbo.MOCTG WITH (NOLOCK)
                                    WHERE MF001 = ME001
                                    AND MF002 = ME002
                                    AND MQ001 = MF004
                                    AND TE001 = MF004
                                    AND TE002 = MF005
                                    AND TE004 = MF001
                                    AND TE010 = MF002
                                    AND TG014 = TE011
                                    AND TG015 = TE012
                                    AND MQ003 IN (
                                    '54'
                                    ,'56'
                                    )
                                    AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN 
                                    (SELECT  LTRIM(RTRIM(MB001)) + LTRIM(RTRIM([LOTNO]))
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    WHERE [KINDS]='2領退料' AND LEVELS='{0}'
                                    GROUP BY MB001,[LOTNO])
                                    ",LEVELS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {

                    return ds.Tables["ds"].Rows.Count;
                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDTRACEBACKLEVELNEXT(string LEVELNOW,string LEVELNEXT)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                     
                                    INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[TG014],[TG015],[MB001],[MB002],[LOTNO],[NUMS])

                                    SELECT  MF001
                                    ,MF002
                                    ,'2領退料'
                                    ,'{0}'
                                    ,MF003
                                    ,MF004
                                    ,MF005
                                    ,MF006
                                    ,TG014
                                    ,TG015
                                    ,TG004
                                    ,''
                                    ,TG017
                                    ,SUM(TG011)
                                    FROM [TK].dbo.INVME WITH (NOLOCK)
                                    ,[TK].dbo.INVMF WITH (NOLOCK)
                                    ,[TK].dbo.CMSMQ WITH (NOLOCK)
                                    ,[TK].dbo.MOCTE WITH (NOLOCK)
                                    ,[TK].dbo.MOCTG WITH (NOLOCK)
                                    WHERE MF001 = ME001
                                    AND MF002 = ME002
                                    AND MQ001 = MF004
                                    AND TE001 = MF004
                                    AND TE002 = MF005
                                    AND TE004 = MF001
                                    AND TE010 = MF002
                                    AND TG014 = TE011
                                    AND TG015 = TE012
                                    AND MQ003 IN (
                                    '54'
                                    ,'56'
                                    )
                                    AND LTRIM(RTRIM(MF001)) + LTRIM(RTRIM(MF002)) IN 
                                    (SELECT  LTRIM(RTRIM(MB001)) + LTRIM(RTRIM([LOTNO]))
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    WHERE [KINDS]='2領退料' AND LEVELS='{1}'
                                    GROUP BY MB001,[LOTNO])
                                    GROUP BY  MF001
                                    ,MF002
                                    ,MF003
                                    ,MF004
                                    ,MF005
                                    ,MF006
                                    ,TG014
                                    ,TG015
                                    ,TG004
                                    ,TG017
                                    ", LEVELNEXT, LEVELNOW);

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

        public void ADDTRACEBACKMOCOUTIN(string MB001, string LOTNO)
        {
            int LEVELNOW = 0;
            int LEVELNEXT = 1;
            int MAXCOUNT = 1;
            int DSCEHCK = 1;


            //新增成品的LEVEL=0
            ADDTRACEBACKPRODUCTMOCOUTINLEVEL0(MB001, LOTNO);

            while (DSCEHCK >= 1 && MAXCOUNT <= 10)
            {
                ADDTRACEBACKLEVELPRODUCTMOCOUTINNEXT(MB001, LOTNO, LEVELNOW.ToString(), LEVELNEXT.ToString());

                LEVELNOW = LEVELNOW + 1;
                LEVELNEXT = LEVELNEXT + 1;
                MAXCOUNT = MAXCOUNT + 1;

                DSCEHCK = CHECKPRODUCTMOCOUTINLEVEL(MB001, LOTNO, LEVELNEXT.ToString(), LEVELNOW.ToString());


            }

            //try
            //{
            //    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //    sqlConn = new SqlConnection(connectionString);

            //    sqlConn.Close();
            //    sqlConn.Open();
            //    tran = sqlConn.BeginTransaction();

            //    sbSql.Clear();


            //    sbSql.AppendFormat(@" 
            //                        WITH RTABLES
            //                        AS (
            //                        SELECT 0 AS LEVELS,[TG001],[TG002],[TG003],[TG004],[TG011],[TG017],[TG014],[TG015],[TE001],[TE002],[TE003],[TE004],[TE005],[TE010]
            //                        FROM [TK].[dbo].[VMOCTGMOCTE] WITH (NOLOCK)
            //                        WHERE [VMOCTGMOCTE].TG004 ='{0}' AND [VMOCTGMOCTE].TG017 ='{1}' 
            //                        UNION ALL
            //                        SELECT LEVELS+1,B.[TG001], B.[TG002], B.[TG003], B.[TG004], B.[TG011], B.[TG017], B.[TG014], B.[TG015], B.[TE001], B.[TE002],B.[TE003], B.[TE004], B.[TE005], B.[TE010]
            //                        FROM [TK].[dbo].[VMOCTGMOCTE] B WITH (NOLOCK)
            //                        INNER JOIN RTABLES ON RTABLES.[TE004]=B.[TG004] AND RTABLES.[TE010]=B.[TG017]
            //                        )

            //                        INSERT INTO [TKMOC].[dbo].[TRACEBACK]
            //                        ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS],[TG014] ,[TG015])

            //                        SELECT '{0}','{1}','3領退料'
            //                        ,LEVELS
            //                        ,(SELECT TOP 1 TC003 FROM [TK].dbo.MOCTC WHERE TC001=TE001  AND TC002=TE002)
            //                        ,[TE001],[TE002],[TE003],[TE004],'',[TE010],[TE005]
            //                        ,(SELECT TOP 1 TE011 FROM [TK].dbo.MOCTE WHERE MOCTE.TE001=RTABLES.TE001  AND MOCTE.TE002=RTABLES.TE002 AND MOCTE.TE003=RTABLES.TE003)
            //                        ,(SELECT TOP 1 TE012 FROM [TK].dbo.MOCTE WHERE MOCTE.TE001=RTABLES.TE001  AND MOCTE.TE002=RTABLES.TE002 AND MOCTE.TE003=RTABLES.TE003)
            //                        FROM RTABLES
            //                        GROUP BY LEVELS,[TE001],[TE002],[TE003],[TE004],[TE010],[TE005]
            //                        ORDER BY LEVELS,[TE001],[TE002],[TE003],[TE004],[TE010],[TE005]

            //                        ", MB001, LOTNO);

            //    cmd.Connection = sqlConn;
            //    cmd.CommandTimeout = 60;
            //    cmd.CommandText = sbSql.ToString();
            //    cmd.Transaction = tran;
            //    result = cmd.ExecuteNonQuery();

            //    if (result == 0)
            //    {
            //        tran.Rollback();    //交易取消
            //    }
            //    else
            //    {
            //        tran.Commit();      //執行交易  


            //    }

            //}
            //catch
            //{

            //}

            //finally
            //{
            //    sqlConn.Close();
            //}
        }

        public void ADDTRACEBACKPRODUCTMOCOUTINLEVEL0(string MB001, string LOTNO)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS],[TG014] ,[TG015])


                                    SELECT '{0}','{1}','3領退料','0'
                                    ,(SELECT TOP 1 TC003 FROM [TK].dbo.MOCTC WHERE TC001=TE001  AND TC002=TE002)
                                    ,[TE001],[TE002],[TE003],[TE004],'',[TE010],[TE005]
                                    ,(SELECT TOP 1 TE011 FROM [TK].dbo.MOCTE TE WHERE TE.TE001=MOCTE.TE001  AND TE.TE002=MOCTE.TE002 AND TE.TE003=MOCTE.TE003)
                                    ,(SELECT TOP 1 TE012 FROM [TK].dbo.MOCTE TE WHERE TE.TE001=MOCTE.TE001  AND TE.TE002=MOCTE.TE002 AND TE.TE003=MOCTE.TE003)
                                    FROM  [TK].dbo.MOCTG  WITH (NOLOCK) 
                                    LEFT OUTER JOIN [TK].dbo.MOCTE  WITH (NOLOCK) ON TE011 = TG014 AND TE012 = TG015
                                    WHERE  TG004='{0}' AND TG017='{1}'
                                    GROUP BY [TE001],[TE002],[TE003],[TE004],[TE010],[TE005]
                                    ORDER BY [TE001],[TE002],[TE003],[TE004],[TE010],[TE005]

                                    ", MB001, LOTNO);

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

        public void ADDTRACEBACKLEVELPRODUCTMOCOUTINNEXT(string MB001, string LOTNO, string LEVELNOW , string LEVELNEXT)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"    
                                    INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS],[TG014] ,[TG015])

                                    SELECT '{0}','{1}','3領退料','{2}'
                                    ,(SELECT TOP 1 TC003 FROM [TK].dbo.MOCTC WHERE TC001=TE001  AND TC002=TE002)
                                    ,[TE001],[TE002],[TE003],[TE004],'',[TE010],[TE005]
                                    ,(SELECT TOP 1 TE011 FROM [TK].dbo.MOCTE TE WHERE TE.TE001=MOCTE.TE001  AND TE.TE002=MOCTE.TE002 AND TE.TE003=MOCTE.TE003)
                                    ,(SELECT TOP 1 TE012 FROM [TK].dbo.MOCTE TE WHERE TE.TE001=MOCTE.TE001  AND TE.TE002=MOCTE.TE002 AND TE.TE003=MOCTE.TE003)
                                    FROM  [TK].dbo.MOCTG  WITH (NOLOCK) 
                                    LEFT OUTER JOIN [TK].dbo.MOCTE  WITH (NOLOCK) ON TE011 = TG014 AND TE012 = TG015
                                    WHERE   LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017)) IN (
		                                                                        SELECT LTRIM(RTRIM([TE004])) + LTRIM(RTRIM([TE010]))
		                                                                        FROM [TK].dbo.MOCTE
			                                                                        ,[TK].dbo.MOCTG
		                                                                        WHERE TG014 = TE011
			                                                                        AND TG015 = TE012
			                                                                        AND LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017)) IN (
				                                                                        SELECT LTRIM(RTRIM(MB001)) + LTRIM(RTRIM(LOTNO))
				                                                                        FROM [TKMOC].[dbo].[TRACEBACK]
				                                                                        WHERE LEVELS = '{3}'
				                                                                        )
		                                                                        GROUP BY LTRIM(RTRIM([TE004])) + LTRIM(RTRIM([TE010]))
		                                                                        )
                                    AND ISNULL([TE010],'')<>''
                                    GROUP BY [TE001],[TE002],[TE003],[TE004],[TE010],[TE005]
                                    ORDER BY [TE001],[TE002],[TE003],[TE004],[TE010],[TE005]


                                    ", MB001, LOTNO, LEVELNEXT, LEVELNOW);

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

        public int CHECKPRODUCTMOCOUTINLEVEL(string MB001, string LOTNO, string LEVELNEXT, string LEVELS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                  SELECT '{0}','{1}','3領退料','{2}'
                                ,(SELECT TOP 1 TC003 FROM [TK].dbo.MOCTC WHERE TC001=TE001  AND TC002=TE002)
                                ,[TE001],[TE002],[TE003],[TE004],'',[TE010],[TE005]
                                ,(SELECT TOP 1 TE011 FROM [TK].dbo.MOCTE TE WHERE TE.TE001=MOCTE.TE001  AND TE.TE002=MOCTE.TE002 AND TE.TE003=MOCTE.TE003)
                                ,(SELECT TOP 1 TE012 FROM [TK].dbo.MOCTE TE WHERE TE.TE001=MOCTE.TE001  AND TE.TE002=MOCTE.TE002 AND TE.TE003=MOCTE.TE003)
                                FROM  [TK].dbo.MOCTG  WITH (NOLOCK) 
                                LEFT OUTER JOIN [TK].dbo.MOCTE  WITH (NOLOCK) ON TE011 = TG014 AND TE012 = TG015
                                WHERE   LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017)) IN (
		                                                                    SELECT LTRIM(RTRIM([TE004])) + LTRIM(RTRIM([TE010]))
		                                                                    FROM [TK].dbo.MOCTE
			                                                                    ,[TK].dbo.MOCTG
		                                                                    WHERE TG014 = TE011
			                                                                    AND TG015 = TE012
			                                                                    AND LTRIM(RTRIM(TG004)) + LTRIM(RTRIM(TG017)) IN (
				                                                                    SELECT LTRIM(RTRIM(MB001)) + LTRIM(RTRIM(LOTNO))
				                                                                    FROM [TKMOC].[dbo].[TRACEBACKTEMP]
				                                                                    WHERE LEVELS = '{3}'
				                                                                    )
		                                                                    GROUP BY LTRIM(RTRIM([TE004])) + LTRIM(RTRIM([TE010]))
		                                                                    )
                                AND ISNULL([TE010],'')<>''
                                GROUP BY [TE001],[TE002],[TE003],[TE004],[TE010],[TE005]
                                ORDER BY [TE001],[TE002],[TE003],[TE004],[TE010],[TE005]
                                    ", MB001, LOTNO, LEVELNEXT, LEVELS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {

                    return ds.Tables["ds"].Rows.Count;
                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDTRACEBACKMOCOUTIN2(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS],TG014,TG015)

                                    SELECT MF001,MF002,'3生產入庫','0',MF003,MF004,MF005,'****' MF006,MF001,'',MF002,SUM(MF010) MF010,TG014,TG015
                                    FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                    WHERE MF001=ME001 AND MF002=ME002
                                    AND MQ001=MF004
                                    AND TG001=MF004 AND TG002=MF005 AND TG003=MF006
                                    AND MQ003='58'
                                    AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002))+RTRIM(LTRIM([TG014]))+RTRIM(LTRIM([TG015])) IN
                                    (
                                    SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO]))+RTRIM(LTRIM([TG014]))+RTRIM(LTRIM([TG015]))
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    WHERE KINDS='2領退料'
                                    )
                                    GROUP BY MF001,MF002,MF003,MF004,MF005,MF001,MF002,TG014,TG015
                                    ORDER BY INVMF.MF002,MF003,MF004,MF005

                                    ", MB001,LOTNO);

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

        public void ADDTRACEBACKINVMF(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                sbSql.AppendFormat(" SELECT '{0}','{1}','4入庫','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010", MB001, LOTNO);
                sbSql.AppendFormat(" FROM [TK].dbo.INVMF WITH (NOLOCK)");
                sbSql.AppendFormat(" WHERE INVMF.MF009 IN ('1')");
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN (SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO])) FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')", MB001, LOTNO);
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF004)) +RTRIM(LTRIM(MF005)) +RTRIM(LTRIM(MF006))  NOT IN (SELECT RTRIM(LTRIM([MID])) +RTRIM(LTRIM([DID])) +RTRIM(LTRIM([SID]))  FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')",MB001,LOTNO);
                sbSql.AppendFormat(" ORDER BY INVMF.MF002,MF004,MF005,MF006,MF001");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                sbSql.AppendFormat(" SELECT '{0}','{1}','5調整','0',MF002,MF004,MF005,MF006,MF001,'',MF002,MF010", MB001, LOTNO);
                sbSql.AppendFormat(" FROM [TK].dbo.INVMF WITH (NOLOCK)");
                sbSql.AppendFormat(" WHERE INVMF.MF009 IN ('5')");
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN (SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO])) FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')", MB001, LOTNO);
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF004)) +RTRIM(LTRIM(MF005)) +RTRIM(LTRIM(MF006))  NOT IN (SELECT RTRIM(LTRIM([MID])) +RTRIM(LTRIM([DID])) +RTRIM(LTRIM([SID]))  FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')", MB001, LOTNO);
                sbSql.AppendFormat(" ORDER BY INVMF.MF002,MF004,MF005,MF006,MF001");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                sbSql.AppendFormat(" SELECT '{0}','{1}','6其他','0',MF002,MF004,MF005,MF006,MF001,'',MF002,MF010", MB001, LOTNO);
                sbSql.AppendFormat(" FROM [TK].dbo.INVMF WITH (NOLOCK)");
                sbSql.AppendFormat(" WHERE INVMF.MF009 IN ('3') AND MF004 LIKE 'A1%'");
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN (SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO])) FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')", MB001, LOTNO);
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF004)) +RTRIM(LTRIM(MF005)) +RTRIM(LTRIM(MF006))  NOT IN (SELECT RTRIM(LTRIM([MID])) +RTRIM(LTRIM([DID])) +RTRIM(LTRIM([SID]))  FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')", MB001, LOTNO);
                sbSql.AppendFormat(" ORDER BY INVMF.MF002,MF004,MF005,MF006,MF001");
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

        public void ADDTRACEBACKINVMFSALE2(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])

                                    SELECT MF001,MF002,'5銷貨','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010
                                    FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK)
                                    WHERE MF001=ME001 AND MF002=ME002
                                    AND MQ001=MF004
                                    AND MF009 IN ('2','5')
                                    AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN
                                    (
                                    SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO]))
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    )
                                    ORDER BY INVMF.MF002,MF003,MF004,MF005

                                ");

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

        public void ADDTRACEBACKINVMF2(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])

                                    SELECT MF001,MF002,'6其他','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010
                                    FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK)
                                    WHERE MF001=ME001 AND MF002=ME002
                                    AND MQ001=MF004
                                    AND MQ003 IN ('11','13','14','15','16','17')
                                    AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN
                                    (
                                    SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO]))
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    )
                                    ORDER BY INVMF.MF002,MF003,MF004,MF005

                                ");

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

        public void UPDATETRACEBACK()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    UPDATE [TKMOC].[dbo].[TRACEBACK]
                                    SET [MMB002]=INVMB.MB002
                                    FROM [TK].dbo.INVMB
                                    WHERE [MMB001]=INVMB.MB001

                                    UPDATE [TKMOC].[dbo].[TRACEBACK]
                                    SET [MB002]=INVMB.MB002
                                    FROM [TK].dbo.INVMB
                                    WHERE [TRACEBACK].[MB001]=INVMB.MB001

                                ");

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

        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();
            Report report1 = new Report();

            SQL = SETSQL();

            report1.Load(@"REPORT\追踨表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();
         
            SB.AppendFormat(@" 
                            SELECT [MMB001] AS '主品號',[MMB002] AS '主品名',[MLOTNO] AS '主批號',[KINDS] AS '類別',[LEVELS] AS '層別',[DATES] AS '日期',[MID] AS '單別',[DID] AS '單號',[SID] AS '序號',[TG014] AS '製令',[TG015] AS '製令號',[MB001] AS '品號',[MB002] AS '品名',[LOTNO] AS '批號',[NUMS] AS '數量'
                            FROM [TKMOC].[dbo].[TRACEBACK]
                            ORDER BY [KINDS],[MMB001],[MLOTNO],[MID],[DID],[SID],[TG014],[TG015]


                            ");

            return SB;

        }

       

        public void SEARCHTRACEBACK1(string STATUS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();
               

                sbSql.AppendFormat(@"  
                                    SELECT [MID] AS '銷貨單別',[DID] AS '銷貨單號',[SID] AS '銷貨序號',[MMB001] AS '主品號',[MMB002] AS '主品名',[MLOTNO] AS '主批號',[KINDS] AS '類別',[LEVELS] AS '層別',[DATES] AS '日期',[TG014] AS '製令',[TG015] AS '製令號',[MB001] AS '品號',[MB002] AS '品名',[LOTNO] AS '批號',[NUMS] AS '數量'
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    WHERE [KINDS] IN ('1銷貨','5銷貨')
                                    ORDER BY [KINDS],[MMB001],[MLOTNO],[MID],[DID],[SID],[TG014],[TG015]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        dataGridView1.AutoResizeColumns();
                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                       
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
        
        public void SETFASTREPORT1(string STATUS)
        {
            StringBuilder SQL = new StringBuilder();
            string SELECT = SELECT1();
            Report report1 = new Report();

            if(!string.IsNullOrEmpty(SELECT))
            {
                SQL.AppendFormat(@"  
                                    SELECT 
                                    CONVERT(NVARCHAR,CONVERT(datetime,TG003),111)  AS '銷貨日期'
                                    ,TG001+'-'+TG002 AS '銷貨單號'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TG042),111)   AS '單據日期'
                                    ,TG004 AS '客戶代號'
                                    ,TG007 AS '客戶簡稱'
                                    ,TG033 AS '總數量'
                                    ,TG020 AS '單頭備註'
                                    ,TH003 AS '序號'
                                    ,TH004 AS '品號'
                                    ,TH005 AS '品名'
                                    ,TH006 AS '規格'
                                    ,TH007 AS '庫別代號'
                                    ,MC002 AS '庫別名稱'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TH106),111)  AS '有效日期'
                                    ,(TH008+TH024) AS '銷貨數量'
                                    ,TH009 AS '單位'
                                    ,TH014+'-'+TH015+'-'+TH016 AS '訂單單號'
                                    ,TH017 AS '批號'
                                    ,TH018 AS '單身備註'
                                    FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.CMSMC
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND MC001=TH007
                                    AND TH001+TH002+TH003 IN ({0})

                                    UNION ALL
                                    SELECT 
                                    CONVERT(NVARCHAR,CONVERT(datetime,TI003),111)  AS '銷貨日期'
                                    ,TI001+'-'+TI002 AS '銷貨單號'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TI034),111)   AS '單據日期'
                                    ,TI004 AS '客戶代號'
                                    ,TI021 AS '客戶簡稱'
                                    ,TI029*-1 AS '總數量'
                                    ,TI020 AS '單頭備註'
                                    ,TJ003 AS '序號'
                                    ,TJ004 AS '品號'
                                    ,TJ005 AS '品名'
                                    ,TJ006 AS '規格'
                                    ,TJ013 AS '庫別代號'
                                    ,MC002 AS '庫別名稱'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TJ096),111)  AS '有效日期'
                                    ,TJ007*-1 AS '銷貨數量'
                                    ,TJ008 AS '單位'
                                    ,TJ018+'-'+TJ019+'-'+TJ020 AS '訂單單號'
                                    ,TJ014 AS '批號'
                                    ,TJ023 AS '單身備註'
                                    FROM [TK].dbo.COPTI,[TK].dbo.COPTJ,[TK].dbo.CMSMC
                                    WHERE TI001=TJ001 AND TI002=TJ002
                                    AND MC001=TJ013
                                    AND TJ001+TJ002+TJ003 IN ({0})
                                    ", SELECT.ToString());

                report1.Load(@"REPORT\銷貨單明細表.frx");

                report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
                Table.SelectCommand = SQL.ToString();

                report1.Preview = previewControl3;
                report1.Show();
            }

           
        }

        public string SELECT1()
        {
            StringBuilder ADDSQL = new StringBuilder();

            foreach (DataGridViewRow dgR in this.dataGridView1.Rows)
            {
                try
                {
                    DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dgR.Cells[0];
                    if ((bool)cbx.FormattedValue)
                    {
                        ADDSQL.AppendFormat(@" '{0}', ", dgR.Cells["銷貨單別"].Value.ToString().Trim()+ dgR.Cells["銷貨單號"].Value.ToString().Trim()+ dgR.Cells["銷貨序號"].Value.ToString().Trim());
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            ADDSQL.AppendFormat(@" '' ");

            return ADDSQL.ToString();

        }

        public void SEARCHTRACEBACK2(string STATUS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT [TG014] AS '製令',[TG015] AS '製令單號',[MID] AS '入庫單別',[DID] AS '入庫單號',[SID] AS '入庫序號',[MMB001] AS '主品號',[MMB002] AS '主品名',[MLOTNO] AS '主批號',[KINDS] AS '類別',[LEVELS] AS '層別',[DATES] AS '日期',[MB001] AS '品號',[MB002] AS '品名',[LOTNO] AS '批號',[NUMS] AS '數量',TA006 AS '生產品號'
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    LEFT JOIN [TK].dbo.MOCTA ON  TA001=TG014 AND TA002=TG015
                                    WHERE [KINDS] IN ('2生產','3生產入庫')
                                    ORDER BY [KINDS],[MMB001],[MLOTNO],[MID],[DID],[SID],[TG014],[TG015]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds.Tables["ds"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        dataGridView2.AutoResizeColumns();
                        dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);

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

        public void SETFASTREPORT2(string STATUS)
        {
            StringBuilder SQL = new StringBuilder();
            string SELECT = SELECT2();
            Report report1 = new Report();

            if (!string.IsNullOrEmpty(SELECT))
            {
                SQL.AppendFormat(@"  
                                    SELECT 
                                    CONVERT(NVARCHAR,CONVERT(datetime,TF003),111)  AS '入庫日期'
                                    ,TF001+'-'+TF002 AS '單別-單號'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TF012),111)  AS '單據日期'
                                    ,TG004 AS '品號'
                                    ,TG005 AS '品名'
                                    ,TG006 AS '規格'
                                    ,TG011 AS '入庫數量'
                                    ,TG007 AS '單位'
                                    ,TG014+'-'+TG015 AS '製令編號'
                                    ,TG017 AS '批號'
                                    ,TA026+'-'+TA027 AS '訂單單號'
                                    ,TG020 AS '備註'
                                    FROM [TK].dbo.MOCTF, [TK].dbo.MOCTG
                                    LEFT JOIN [TK].dbo.MOCTA ON TA001=TG014 AND TA002=TG015
                                    WHERE TF001=TG001 AND TF002=TG002
                                    AND TG014+TG015 IN ({0})
                                    ORDER BY TF001,TF002,TG004
                                    ", SELECT.ToString());

                report1.Load(@"REPORT\生產入庫單明細表.frx");

                report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
                Table.SelectCommand = SQL.ToString();

                report1.Preview = previewControl4;
                report1.Show();
            }


        }

        public string SELECT2()
        {
            StringBuilder ADDSQL = new StringBuilder();

            foreach (DataGridViewRow dgR in this.dataGridView2.Rows)
            {
                try
                {
                    DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dgR.Cells[0];
                    if ((bool)cbx.FormattedValue)
                    {
                        ADDSQL.AppendFormat(@" '{0}', ", dgR.Cells["製令"].Value.ToString().Trim() + dgR.Cells["製令單號"].Value.ToString().Trim());
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            ADDSQL.AppendFormat(@" '' ");

            return ADDSQL.ToString();

        }

        public void SEARCHTRACEBACK3(string STATUS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT [TG014] AS '製令',[TG015] AS '製令單號',[MID] AS '單別',[DID] AS '單號',[SID] AS '序號',[MMB001] AS '主品號',[MMB002] AS '主品名',[MLOTNO] AS '主批號',[KINDS] AS '類別',[LEVELS] AS '層別',[DATES] AS '日期',[MB001] AS '品號',[MB002] AS '品名',[LOTNO] AS '批號',[NUMS] AS '數量',TA006 AS '生產品號'
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    LEFT JOIN [TK].dbo.MOCTA ON  TA001=TG014 AND TA002=TG015 
                                    WHERE [KINDS] IN ('3領退料','2領退料') AND [MID] LIKE 'A54%'
                                    ORDER BY [KINDS],[MMB001],[MLOTNO],[MID],[DID],[SID],[TG014],[TG015]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds.Tables["ds"];
                        dataGridView3.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        dataGridView3.AutoResizeColumns();
                        dataGridView3.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);

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

        public void SETFASTREPORT3(string STATUS)
        {
            StringBuilder SQL = new StringBuilder();
            string SELECT = SELECT3();
            Report report1 = new Report();

            if (!string.IsNullOrEmpty(SELECT))
            {
                SQL.AppendFormat(@"  
                                    SELECT CONVERT(NVARCHAR,CONVERT(datetime,TC003),111) AS '領料日期'
                                    ,TC001+'-'+TC002 AS '領料單號'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TC014),111) AS '單據日期'
                                    ,TE004 AS '材料品號'
                                    ,TE017 AS '品名'
                                    ,TE018 AS '規格'
                                    ,TE005 AS '領料數量'
                                    ,TE006 AS '單位'
                                    ,TE011+'-'+TE012 AS '製令單號'
                                    ,MC002 AS '庫別名稱'
                                    ,TE010 AS '批號'
                                    ,TE013 AS '領料說明'
                                    ,TE014 AS '備註'
                                    FROM [TK].dbo.MOCTC,[TK].dbo.MOCTE,[TK].dbo.CMSMC,[TK].dbo.CMSMQ
                                    WHERE TC001=TE001 AND TC002=TE002
                                    AND TE008=MC001
                                    AND TC001=MQ001 AND MQ003 IN ('54','55')
                                    AND TE011+TE012 IN ({0})
                                    ORDER BY TC001,TC002,TE003
                                    ", SELECT.ToString());

                report1.Load(@"REPORT\領料單明細表.frx");

                report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
                Table.SelectCommand = SQL.ToString();

                report1.Preview = previewControl5;
                report1.Show();
            }


        }

        public string SELECT3()
        {
            StringBuilder ADDSQL = new StringBuilder();

            foreach (DataGridViewRow dgR in this.dataGridView3.Rows)
            {
                try
                {
                    DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dgR.Cells[0];
                    if ((bool)cbx.FormattedValue)
                    {
                        ADDSQL.AppendFormat(@" '{0}', ", dgR.Cells["製令"].Value.ToString().Trim()+ dgR.Cells["製令單號"].Value.ToString().Trim());
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            ADDSQL.AppendFormat(@" '' ");

            return ADDSQL.ToString();

        }

        public void SEARCHTRACEBACK4(string STATUS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT [TG014] AS '製令',[TG015] AS '製令單號',[MID] AS '單別',[DID] AS '單號',[SID] AS '序號',[MMB001] AS '主品號',[MMB002] AS '主品名',[MLOTNO] AS '主批號',[KINDS] AS '類別',[LEVELS] AS '層別',[DATES] AS '日期',[MB001] AS '品號',[MB002] AS '品名',[LOTNO] AS '批號',[NUMS] AS '數量',TA006 AS '生產品號'
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    LEFT JOIN [TK].dbo.MOCTA ON  TA001=TG014 AND TA002=TG015 
                                    WHERE [KINDS] IN ('3領退料','2領退料') AND [MID] LIKE 'A56%'
                                    ORDER BY [KINDS],[MMB001],[MLOTNO],[MID],[DID],[SID],[TG014],[TG015]
                                    ");

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

                        dataGridView4.AutoResizeColumns();
                        dataGridView4.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);

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

        public void SETFASTREPORT4(string STATUS)
        {
            StringBuilder SQL = new StringBuilder();
            string SELECT = SELECT4();
            Report report1 = new Report();

            if (!string.IsNullOrEmpty(SELECT))
            {
                SQL.AppendFormat(@"  
                                    SELECT CONVERT(NVARCHAR,CONVERT(datetime,TC003),111) AS '退料日期'
                                    ,TC001+'-'+TC002 AS '退料單號'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TC014),111) AS '單據日期'
                                    ,TE004 AS '材料品號'
                                    ,TE017 AS '品名'
                                    ,TE018 AS '規格'
                                    ,TE005 AS '退料數量'
                                    ,TE006 AS '單位'
                                    ,TE011+'-'+TE012 AS '製令單號'
                                    ,MC002 AS '庫別名稱'
                                    ,TE010 AS '批號'
                                    ,TE013 AS '退料說明'
                                    ,TE014 AS '備註'
                                    FROM [TK].dbo.MOCTC,[TK].dbo.MOCTE,[TK].dbo.CMSMC,[TK].dbo.CMSMQ
                                    WHERE  TC001=TE001 AND TC002=TE002
                                    AND TE008=MC001
                                    AND TC001=MQ001 AND MQ003 IN ('56','57')
                                    AND TE011+TE012 IN ({0})
                                    ORDER BY TC001,TC002,TE003
                                    ", SELECT.ToString());

                report1.Load(@"REPORT\退料單明細表.frx");

                report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
                Table.SelectCommand = SQL.ToString();

                report1.Preview = previewControl6;
                report1.Show();
            }


        }

        public string SELECT4()
        {
            StringBuilder ADDSQL = new StringBuilder();

            foreach (DataGridViewRow dgR in this.dataGridView4.Rows)
            {
                try
                {
                    DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dgR.Cells[0];
                    if ((bool)cbx.FormattedValue)
                    {
                        ADDSQL.AppendFormat(@" '{0}', ", dgR.Cells["製令"].Value.ToString().Trim() + dgR.Cells["製令單號"].Value.ToString().Trim());
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            ADDSQL.AppendFormat(@" '' ");

            return ADDSQL.ToString();

        }

        public void SERACHDYCOPTGCOPTH(string LOTNO)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT TH004 AS '品號',TH005 AS '品名',TH017 AS '批號'
                                    FROM [DY].dbo.COPTG,[DY].dbo.COPTH
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND TH017 LIKE '{0}%'
                                    GROUP BY TH004,TH005,TH017
                                    ORDER BY TH004,TH005,TH017
                                    ", LOTNO);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds.Tables["ds"];
                        dataGridView5.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        dataGridView5.AutoResizeColumns();
                        dataGridView5.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);

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

        public void SETFASTREPORT5(string STATUS)
        {
            StringBuilder SQL = new StringBuilder();
            string SELECT = SELECT5();
            Report report1 = new Report();

            if (!string.IsNullOrEmpty(SELECT))
            {
                SQL.AppendFormat(@"  
                                    SELECT 
                                    CONVERT(NVARCHAR,CONVERT(datetime,TG003),111)  AS '銷貨日期'
                                    ,TG001+'-'+TG002 AS '銷貨單號'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TG042),111)   AS '單據日期'
                                    ,TG004 AS '客戶代號'
                                    ,TG007 AS '客戶簡稱'
                                    ,TG033 AS '總數量'
                                    ,TG020 AS '單頭備註'
                                    ,TH003 AS '序號'
                                    ,TH004 AS '品號'
                                    ,TH005 AS '品名'
                                    ,TH006 AS '規格'
                                    ,TH007 AS '庫別代號'
                                    ,MC002 AS '庫別名稱'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TH106),111)  AS '有效日期'
                                    ,TH008 AS '銷貨數量'
                                    ,TH009 AS '單位'
                                    ,TH014+'-'+TH015+'-'+TH016 AS '訂單單號'
                                    ,TH017 AS '批號'
                                    ,TH018 AS '單身備註'
                                    FROM [DY].dbo.COPTG,[DY].dbo.COPTH,[DY].dbo.CMSMC
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND MC001=TH007
                                    AND TH004+TH017 IN ({0})
                                    ", SELECT.ToString());

                report1.Load(@"REPORT\銷貨單明細表.frx");

                report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
                Table.SelectCommand = SQL.ToString();

                report1.Preview = previewControl7;
                report1.Show();
            }


        }

        public string SELECT5()
        {
            StringBuilder ADDSQL = new StringBuilder();

            foreach (DataGridViewRow dgR in this.dataGridView5.Rows)
            {
                try
                {
                    DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dgR.Cells[0];
                    if ((bool)cbx.FormattedValue)
                    {
                        ADDSQL.AppendFormat(@" '{0}', ", dgR.Cells["品號"].Value.ToString().Trim() + dgR.Cells["批號"].Value.ToString().Trim() );
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            ADDSQL.AppendFormat(@" '' ");

            return ADDSQL.ToString();

        }

        public void DG1CHECKALL()
        {

            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = true;

            }
        }
        public void DG2CHECKALL()
        {

            dataGridView2.EndEdit();

            foreach (DataGridViewRow dr in dataGridView2.Rows)
            {
                dr.Cells[0].Value = true;

            }
        }
        public void DG2CHECKALL1()
        {

            dataGridView2.EndEdit();

            foreach (DataGridViewRow dr in dataGridView2.Rows)
            {
                dr.Cells[0].Value = false;

            }

            foreach (DataGridViewRow dr in dataGridView2.Rows)
            {

                if (dr.Cells["生產品號"].Value.ToString().StartsWith("3"))
                {
                    dr.Cells[0].Value = true;
                }

                    

            }
        }
        public void DG2CHECKALL2()
        {

            dataGridView2.EndEdit();


            foreach (DataGridViewRow dr in dataGridView2.Rows)
            {
                dr.Cells[0].Value = false;

            }

            foreach (DataGridViewRow dr in dataGridView2.Rows)
            {

                if (dr.Cells["生產品號"].Value.ToString().StartsWith("4"))
                {
                    dr.Cells[0].Value = true;
                }
                  

            }
        }
        public void DG3CHECKALL()
        {

            dataGridView3.EndEdit();

            foreach (DataGridViewRow dr in dataGridView3.Rows)
            {
                dr.Cells[0].Value = true;

            }
        }

        public void DG3CHECKALL1()
        {

            dataGridView3.EndEdit();

            foreach (DataGridViewRow dr in dataGridView3.Rows)
            {
                dr.Cells[0].Value = false;

            }

            foreach (DataGridViewRow dr in dataGridView3.Rows)
            {
                if(dr.Cells["生產品號"].Value.ToString().StartsWith("3"))
                {
                    dr.Cells[0].Value = true;
                }
                

            }
        }

        public void DG3CHECKALL2()
        {

            dataGridView3.EndEdit();

            foreach (DataGridViewRow dr in dataGridView3.Rows)
            {
                dr.Cells[0].Value = false;

            }

            foreach (DataGridViewRow dr in dataGridView3.Rows)
            {
                if (dr.Cells["生產品號"].Value.ToString().StartsWith("4"))
                {
                    dr.Cells[0].Value = true;
                }                  

            }
        }
        public void DG4CHECKALL()
        {

            dataGridView4.EndEdit();

            foreach (DataGridViewRow dr in dataGridView4.Rows)
            {
                dr.Cells[0].Value = true;

            }
        }
        public void DG4CHECKALL1()
        {

            dataGridView4.EndEdit();


            foreach (DataGridViewRow dr in dataGridView4.Rows)
            {
                dr.Cells[0].Value = false;

            }

            foreach (DataGridViewRow dr in dataGridView4.Rows)
            {
                if (dr.Cells["生產品號"].Value.ToString().StartsWith("3"))
                {
                    dr.Cells[0].Value = true;
                }

            }
        }
        public void DG4CHECKALL2()
        {

            dataGridView4.EndEdit();


            foreach (DataGridViewRow dr in dataGridView4.Rows)
            {
                dr.Cells[0].Value = false;

            }

            foreach (DataGridViewRow dr in dataGridView4.Rows)
            {
                if (dr.Cells["生產品號"].Value.ToString().StartsWith("4"))
                {
                    dr.Cells[0].Value = true;
                }

            }
        }
        public void DG5CHECKALL()
        {

            dataGridView5.EndEdit();

            foreach (DataGridViewRow dr in dataGridView5.Rows)
            {
                dr.Cells[0].Value = true;

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
                    DELETETRACEBACK();
                    SEARCHOUT(textBox1.Text.Trim(), textBox2.Text.Trim());
                    UPDATETRACEBACK();
                }
                else if (comboBox1.Text.Trim().Equals("原料順溯"))
                {
                    DELETETRACEBACK();
                    SEARCHOUT2(textBox1.Text.Trim(), textBox2.Text.Trim());
                    UPDATETRACEBACK();
                }


            }

            SETFASTREPORT();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHTRACEBACK1("1銷貨");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT1("1銷貨");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SEARCHTRACEBACK2("2生產");
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2("2生產");
        }
        private void button6_Click(object sender, EventArgs e)
        {
            SEARCHTRACEBACK3("3領退料");
        }
        private void button7_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3("3領退料");
        }
        private void button8_Click(object sender, EventArgs e)
        {
            SEARCHTRACEBACK4("3領退料");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4("3領退料");
        }
        private void button10_Click(object sender, EventArgs e)
        {
            SERACHDYCOPTGCOPTH(textBox3.Text.Trim());
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SETFASTREPORT5("1銷貨");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            DG3CHECKALL();
        }
        private void button13_Click(object sender, EventArgs e)
        {
            DG4CHECKALL();
        }
        private void button14_Click(object sender, EventArgs e)
        {
            DG2CHECKALL();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DG1CHECKALL();
        }
        private void button16_Click(object sender, EventArgs e)
        {
            DG5CHECKALL();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            DG3CHECKALL1();
        }
        private void button18_Click(object sender, EventArgs e)
        {
            DG3CHECKALL2();
        }
        private void button19_Click(object sender, EventArgs e)
        {
            DG4CHECKALL1();
        }

        private void button21_Click(object sender, EventArgs e)
        {
            DG4CHECKALL2();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            DG2CHECKALL1();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            DG2CHECKALL2();
        }
        #endregion


    }
}
