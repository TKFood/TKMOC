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
using System.Text.RegularExpressions;

namespace TKMOC
{
    public partial class frmMOCPLANWEEK : Form
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
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        DataTable dtTemp = new DataTable();
        DataTable dtTemp2 = new DataTable();
        int result;
        string tablename = null;
        decimal COPNum = 0;
        decimal TOTALCOPNum = 0;
        double BOMNum = 0;
        double FinalNum = 0;
        decimal COOKIES = 1;
        decimal BATCH = 1;
        Thread TD;
        string CHECKYN = "N";
        decimal MOCBATCH = 1;
        string TD001 = null;
        string TD002 = null;
        string TD003 = null;


        public frmMOCPLANWEEK()
        {
            InitializeComponent();
            FINDWEKKDATE();

            dtTemp.Columns.Add("日期");
            dtTemp.Columns.Add("品號");
            dtTemp.Columns.Add("品名");
            dtTemp.Columns.Add("數量");
            dtTemp.Columns.Add("標準批量");
            dtTemp.Columns.Add("桶數");
            dtTemp.Columns.Add("標準時間");

            



        }

        #region FUNCTION
        public void Search()
        {
            StringBuilder TD001 = new StringBuilder();
            StringBuilder TC027 = new StringBuilder();
            StringBuilder PALNQUERY = new StringBuilder();

            if (checkBox1.Checked == true)
            {
                TD001.Append("'A221',");
            }
            if (checkBox2.Checked == true)
            {
                TD001.Append("'A222',");
            }

            if (checkBox4.Checked == true)
            {
                TD001.Append("'A225',");
            }
            if (checkBox5.Checked == true)
            {
                TD001.Append("'A226',");
            }
            if (checkBox6.Checked == true)
            {
                TD001.Append("'A227',");
            }
            if (checkBox7.Checked == true)
            {
                TD001.Append("'A223',");
            }
            TD001.Append("''");

            if (comboBox1.Text.ToString().Equals("已確認"))
            {
                TC027.Append(" 'Y',");
            }
            else if (comboBox1.Text.ToString().Equals("未確認"))
            {
                TC027.Append(" 'N',");
            }
            else if (comboBox1.Text.ToString().Equals("全部"))
            {
                TC027.Append(" 'Y','N', ");
            }
            TC027.Append("''");

            if (comboBox2.Text.ToString().Equals("未排計畫"))
            {
                PALNQUERY.AppendFormat("AND NOT  EXISTS  (SELECT TD001 FROM [TKMOC].[dbo].[MOCPLANWEEK] WHERE [MOCPLANWEEK].TD001=COPTD.TD001 AND [MOCPLANWEEK].TD002=COPTD.TD002 AND [MOCPLANWEEK].TD003=COPTD.TD003 AND [YEARS]='{0}' AND [WEEKS]='{1}')    ",numericUpDown1.Value.ToString(),numericUpDown2.Value.ToString());
            }
            else if(comboBox2.Text.ToString().Equals("已排計畫"))
            {
                PALNQUERY.AppendFormat("AND   EXISTS  (SELECT TD001 FROM [TKMOC].[dbo].[MOCPLANWEEK] WHERE [MOCPLANWEEK].TD001=COPTD.TD001 AND [MOCPLANWEEK].TD002=COPTD.TD002 AND [MOCPLANWEEK].TD003=COPTD.TD003 AND [YEARS]='{0}' AND [WEEKS]='{1}')    ", numericUpDown1.Value.ToString(), numericUpDown2.Value.ToString());
            }
            else if (comboBox2.Text.ToString().Equals("未排計畫"))
            {
                PALNQUERY.Append("  ");
            }

            dtTemp.Clear();
            dtTemp2.Clear();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@"  SELECT 客戶,日期,品號,品名,規格,CONVERT(INT,SUM(訂單數量)) AS 訂單數量,單位 ,單別,單號,序號  ");
                sbSql.Append(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20001' AND LA001=品號) AS '成品倉庫存'");
                sbSql.Append(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20002' AND LA001=品號) AS '外銷倉庫存'");
                sbSql.Append(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(TA015-TA017-TA018),0)) FROM [TK].dbo.MOCTA  WITH (NOLOCK) WHERE TA011 NOT IN ('Y','y') AND TA006=品號 ) AS '未完成的製令' ");
                sbSql.Append(@"  ,(SELECT CONVERT(INT,ISNULL(MC004,0))  FROM [TK].dbo.BOMMC WHERE MC001=品號) AS 標準批量");
                sbSql.Append(@"  FROM (");
                sbSql.Append(@"  SELECT   TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TC053  AS '客戶' ,TD013 AS '日期',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格'");
                sbSql.Append(@"  ,(CASE WHEN MB004=TD010 THEN (TD008-TD009) ELSE (TD008-TD009)*MD004 END) AS '訂單數量'");
                sbSql.Append(@"  ,MB004 AS '單位'");
                sbSql.Append(@"  ,(TD008-TD009) AS '訂單量'");
                sbSql.Append(@"  ,TD010 AS '訂單單位' ");
                sbSql.Append(@"  ,(CASE WHEN ISNULL(MD002,'')<>'' THEN MD002 ELSE TD010 END ) AS '換算單位'");
                sbSql.Append(@"  ,(CASE WHEN MD003>0 THEN MD003 ELSE 1 END) AS '分子'");
                sbSql.Append(@"  ,(CASE WHEN MD004>0 THEN MD004 ELSE (TD008-TD009) END ) AS '分母'");
                sbSql.Append(@"  FROM [TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.Append(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.Append(@"  WHERE TD004=MB001");
                sbSql.Append(@"  AND TC001=TD001 AND TC002=TD002");
                sbSql.Append(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TC001 IN ({0}) ", TD001.ToString());
                sbSql.Append(@"  AND (TD008-TD009)>0  ");
                sbSql.AppendFormat(@" AND TC027 IN ({0})  ", TC027.ToString());
                sbSql.AppendFormat(@"  {0}", PALNQUERY.ToString());                
                //sbSql.Append(@"  AND ( TD004 LIKE '40109916000740%'  ) ");
                sbSql.Append(@"  ) AS TEMP");
                sbSql.Append(@"  GROUP  BY 客戶,日期,品號,品名,規格,單位,單別,單號,序號");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if(CHECKYN.Equals("N"))
                {
                    //建立一個DataGridView的Column物件及其內容
                    DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
                    dgvc.Width = 40;
                    dgvc.Name = "選取";

                    this.dataGridView1.Columns.Insert(0, dgvc);
                    CHECKYN = "Y";
                }
                

                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    labelget.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        labelget.Text = "有 " + ds.Tables["TEMPds1"].Rows.Count.ToString() + " 筆";
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

        public void ADDTOMOCPLANWEEK()
        {
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    try
                    {
                        connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                        sqlConn = new SqlConnection(connectionString);

                        sqlConn.Close();
                        sqlConn.Open();
                        tran = sqlConn.BeginTransaction();

                        sbSql.Clear();
                        sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MOCPLANWEEK] ");
                        sbSql.Append(" ([ID],[YEARS],[WEEKS],[SDATE],[EDATE],[TD001],[TD002],[TD003],[TD004],[TD005],[TD006],[TD008],[TD009],[TD013],[MC004]) ");
                        sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}') ", Guid.NewGuid(), numericUpDown1.Value.ToString(),numericUpDown2.Value.ToString(),dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"), dr.Cells["單別"].Value.ToString(), dr.Cells["單號"].Value.ToString(), dr.Cells["序號"].Value.ToString(), dr.Cells["品號"].Value.ToString(), dr.Cells["品名"].Value.ToString(), dr.Cells["規格"].Value.ToString(), dr.Cells["訂單數量"].Value.ToString(), dr.Cells["單位"].Value.ToString(), dr.Cells["日期"].Value.ToString(), dr.Cells["標準批量"].Value.ToString());

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
            }

        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            FINDWEKKDATE();
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            FINDWEKKDATE();
        }

        public void FINDWEKKDATE()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  declare @num int,@year varchar(4),@date datetime");
                sbSql.AppendFormat(@"  select @num={0}",numericUpDown2.Value.ToString());
                sbSql.AppendFormat(@"  select @year='{0}'", numericUpDown1.Value.ToString() + "/1/1");
                sbSql.AppendFormat(@"  select @date=dateadd(wk,@num-1,@year)");
                sbSql.AppendFormat(@"  select CONVERT(varchar(100),(dateadd(dd,1-datepart(dw,@date),@date)), 111)  AS 'SDATE'");
                sbSql.AppendFormat(@"  ,CONVERT(varchar(100),dateadd(dd,7-datepart(dw,@date),@date), 111) AS 'EDATE'");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dateTimePicker3.Value = Convert.ToDateTime(numericUpDown1.Value.ToString()+"/1/1");
                    dateTimePicker4.Value = Convert.ToDateTime(numericUpDown1.Value.ToString()+"/1/1");
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dateTimePicker3.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["SDATE"].ToString());
                        dateTimePicker4.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["EDATE"].ToString());

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

        public void SEARCHPLANWEEK()
        {
         
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT  [YEARS] AS '年度',[WEEKS]  AS '週次',[SDATE]  AS '開始日',[EDATE]  AS '結束日'");
                sbSql.AppendFormat(@"  ,[TD001]  AS '單別',[TD002]  AS '單號',[TD003]  AS '序號'");
                sbSql.AppendFormat(@"  ,[TD004]  AS '品號',[TD005]  AS '品名',[TD006]  AS '規格',[TD008]  AS '數量',[TD009]  AS '單位'");
                sbSql.AppendFormat(@"  ,[TD013] AS '日期' ,[MC004] AS '標準批量' ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCPLANWEEK]");
                sbSql.AppendFormat(@"  WHERE [YEARS]='{0}' AND [WEEKS]='{1}'",numericUpDown1.Value.ToString(),numericUpDown2.Value.ToString());
                sbSql.AppendFormat(@"  ORDER BY TD001,TD002,TD003");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    labelget.Text = "找不到資料";
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        //labelget.Text = "有 " + ds.Tables["TEMPds1"].Rows.Count.ToString() + " 筆";
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];

                        dataGridView2.AutoResizeColumns();
                        dataGridView2.FirstDisplayedScrollingRowIndex = dataGridView2.RowCount - 1;


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

        public void SEARCHCOOKIES()
        {
            string MB003 = null;
            string[] sArray = null;
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

            dtTemp.Clear();

            for (int i = 0; i < ds2.Tables["TEMPds2"].Rows.Count; i++)
            {

                COPNum = Convert.ToDecimal(ds2.Tables["TEMPds2"].Rows[i]["數量"].ToString());
                MB003 = ds2.Tables["TEMPds2"].Rows[i]["規格"].ToString();
                sArray = MB003.Split('g');
                //TOTALCOPNum = Convert.ToDecimal(Convert.ToDecimal(sArray[0].ToString())* COPNum);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MD003,INVMB.MB002,CASE WHEN ISNULL(INVMB.MB003,'')=''  THEN '1' ELSE INVMB.MB003 END AS MB003  ,MD004,MD006,[PROCESSNUM],[PROCESSTIME]  ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  LEFT JOIN   [TKMOC].[dbo].[ERPINVMB] ON [ERPINVMB].[MB001]=INVMB.MB001");
                //sbSql.AppendFormat(@"  WHERE MD003=INVMB.MB001  AND MD003 LIKE '3%' AND INVMB.MB002 NOT LIKE '%水麵%'  AND  INVMB.MB002 NOT LIKE '%餅麩%'  ");
                sbSql.AppendFormat(@"  WHERE MD003=INVMB.MB001  AND MD003 LIKE '3%'   ");
                sbSql.AppendFormat(@"  AND MD001='{0}'", ds2.Tables["TEMPds2"].Rows[i]["品號"].ToString());
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds3.Clear();
                adapter.Fill(ds3, "TEMPds3");
                sqlConn.Close();

                if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                {

                    foreach (DataRow od2 in ds3.Tables["TEMPds3"].Rows)
                    {
                        DataRow row = dtTemp.NewRow();
                        //row["MD001"] = od2["MC001"].ToString();
                        row["日期"] = ds2.Tables["TEMPds2"].Rows[i]["日期"].ToString();
                        row["品號"] = od2["MD003"].ToString();
                        row["品名"] = od2["MB002"].ToString();
                        if(!string.IsNullOrEmpty(od2["MB003"].ToString()))
                        {
                            COOKIES = Convert.ToDecimal(Regex.Replace(od2["MB003"].ToString(), "[^0-9]", ""));
                        }
                        else
                        {
                            COOKIES = 1;
                        }
                        if (!string.IsNullOrEmpty(od2["MD006"].ToString()))
                        {
                            TOTALCOPNum = Convert.ToDecimal(Convert.ToDecimal(od2["MD006"].ToString()) * 1000 * COPNum);
                        }
                        else
                        {
                            TOTALCOPNum = 1;
                        }
                        if (!string.IsNullOrEmpty(od2["MB003"].ToString()))
                        {
                            BATCH = Convert.ToDecimal(ds2.Tables["TEMPds2"].Rows[i]["標準批量"].ToString());
                        }
                        else
                        {
                            BATCH = 1;
                        }
                       
                        row["數量"] = Convert.ToInt32(TOTALCOPNum / COOKIES / BATCH);
                        
                        if (!string.IsNullOrEmpty(od2["PROCESSNUM"].ToString()))
                        {
                            if(Convert.ToDecimal(od2["PROCESSNUM"].ToString())>0)
                            {
                                MOCBATCH = Convert.ToDecimal(od2["PROCESSNUM"].ToString());
                            }
                            else
                            {
                                MOCBATCH = 1;
                            }

                        }
                        else
                        {
                            MOCBATCH = 1;
                        }
                        row["桶數"] = Convert.ToInt32(TOTALCOPNum / COOKIES / BATCH/ MOCBATCH);
                        row["標準批量"] = od2["PROCESSNUM"].ToString();
                        row["標準時間"] = od2["PROCESSTIME"].ToString();
                        dtTemp.Rows.Add(row);
                    }

                }

            }

            ////分組並計算

            //var Query = from p in dtTemp.AsEnumerable()
            //            orderby p.Field<string>("MD003")
            //            group p by new { MD003 = p.Field<string>("MD003"), UNIT = p.Field<string>("UNIT") } into g
            //            select new
            //            {
            //                //MD003 = g.Key,
            //                MD003 = g.Key.MD003,
            //                NUM = g.Sum(p => Convert.ToDouble(p.Field<string>("NUM"))),
            //                UNIT = g.Key.UNIT
            //            };


            //if (Query.Count() >= 1)
            //{
            //    foreach (var c in Query)
            //    {
            //        sbSql.Clear();
            //        sbSqlQuery.Clear();

            //        //sbSql.AppendFormat(@"  SELECT TOP 1 MB001,MB002,MB003  FROM [TK].dbo.INVMB WITH (NOLOCK)  WHERE   MB001='{0}'  ", c.MD003.ToString());
            //        sbSql.AppendFormat(@"  SELECT TOP 1 MB001,MB002,MB003,ISNULL(MC004,0) AS MC004 ,(CASE WHEN ISNULL(MC001,'')<>'' THEN CEILING({0}/MC004) ELSE CEILING({0}) END) AS NN", c.NUM.ToString());
            //        sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20001' AND LA001=MB001) AS NN1");
            //        sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20002' AND LA001=MB001) AS NN2");

            //        sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB WITH (NOLOCK)  ");
            //        sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC WITH (NOLOCK)  ON MC001=MB001");
            //        sbSql.AppendFormat(@"  WHERE    MB001='{0}'  ", c.MD003.ToString());

            //        adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

            //        sqlCmdBuilder = new SqlCommandBuilder(adapter);
            //        sqlConn.Open();
            //        ds3.Clear();
            //        adapter.Fill(ds3, "TEMPds3");
            //        sqlConn.Close();

            //        if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
            //        {
            //            DataRow row = dtTemp2.NewRow();
            //            row["品號"] = c.MD003;
            //            row["品名"] = ds3.Tables["TEMPds3"].Rows[0]["MB002"].ToString();                        
            //            row["規格"] = ds3.Tables["TEMPds3"].Rows[0]["MB003"].ToString();
            //            row["預計用量"] = Convert.ToDouble(c.NUM);
            //            row["單位"] = c.UNIT;
            //            COOKIES =Convert.ToDouble (Regex.Replace(ds3.Tables["TEMPds3"].Rows[0]["MB003"].ToString(), "[^0-9]", ""));
            //            row["需求片數"] = (Convert.ToDouble(c.NUM*1000/ COOKIES));
            //            row["生產批量"] = ds3.Tables["TEMPds3"].Rows[0]["MC004"].ToString();
            //            row["預計生產批量"] = ds3.Tables["TEMPds3"].Rows[0]["NN"].ToString();
            //            row["成品庫存"] = ds3.Tables["TEMPds3"].Rows[0]["NN1"].ToString();
            //            row["外銷庫存"] = ds3.Tables["TEMPds3"].Rows[0]["NN2"].ToString();
            //            dtTemp2.Rows.Add(row);
            //        }
            //    }
            //}


            //dataGridView1.DataSource = dtQuery.ToList();
            //label14.Text = "有 " + dtTemp2.Rows.Count.ToString() + " 筆";
            //dataGridView3.Rows.Clear();

            dataGridView3.DataSource = dtTemp;
            dataGridView3.AutoResizeColumns();
        }

        public void ExcelExportCOP()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = ds.Tables["TEMPds1"];

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dt.Rows.Count - 1;
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {

                if (j <= k)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString());
                    ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));
                    ws.GetRow(j + 1).CreateCell(11).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString()));
                    ws.GetRow(j + 1).CreateCell(12).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString()));
                    ws.GetRow(j + 1).CreateCell(13).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString()));

                    j++;
                }

            }



            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\預計訂單{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }

        public void ExcelExportBOM()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = dtTemp;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dt.Rows.Count - 1;
            foreach (DataGridViewRow dr in this.dataGridView3.Rows)
            {
                if (j <= k)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));

                    j++;
                }

            }



            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\預計用量{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }

        public void ExcelExportPLAN()
        {
            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = ds2.Tables["TEMPds2"];

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dt.Rows.Count - 1;
            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                if (j <= k)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString());
                    ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));
                    ws.GetRow(j + 1).CreateCell(11).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString());
                    ws.GetRow(j + 1).CreateCell(12).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString());
                    ws.GetRow(j + 1).CreateCell(13).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString()));

                    j++;
                }

            }



            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\預計計劃{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }

        }

        public void DELMOCPLANWEEK()
        {
            DialogResult dialogResult = MessageBox.Show("確定要刪除?", "?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE  [TKMOC].[dbo].[MOCPLANWEEK]");
                    sbSql.AppendFormat("  WHERE [YEARS]='{0}' AND [WEEKS]='{1}' AND [TD001]='{2}' AND [TD002]='{3}' AND [TD003]='{4}'", numericUpDown1.Value.ToString(), numericUpDown2.Value.ToString(), TD001, TD002, TD003);
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
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    TD001= row.Cells["單別"].Value.ToString();
                    TD002 = row.Cells["單號"].Value.ToString();
                    TD003 = row.Cells["序號"].Value.ToString();

                }
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
            SEARCHPLANWEEK();
            SEARCHCOOKIES();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADDTOMOCPLANWEEK();
            button5.PerformClick();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SEARCHPLANWEEK();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCHPLANWEEK();
            SEARCHCOOKIES();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExcelExportBOM();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            ExcelExportCOP();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            ExcelExportPLAN();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DELMOCPLANWEEK();
            SEARCHPLANWEEK();
            SEARCHCOOKIES();
        }

        #endregion

       
    }
}
