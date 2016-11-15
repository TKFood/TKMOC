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
    public partial class frmCOPMOC : Form
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
       
        string tablename = null;
        decimal COPNum = 0;
        decimal TOTALCOPNum = 0;
        double BOMNum = 0;
        double FinalNum = 0;
        decimal COOKIES = 1;
        decimal BATCH = 1;
        Thread TD;

        public frmCOPMOC()
        {
            InitializeComponent();

            dtTemp.Columns.Add("DATE");
            dtTemp.Columns.Add("MD003");
            dtTemp.Columns.Add("MB002");
            dtTemp.Columns.Add("NUM");
            

            dtTemp2.Columns.Add("品號");
            dtTemp2.Columns.Add("品名");
            dtTemp2.Columns.Add("規格");
            dtTemp2.Columns.Add("預計用量");
            dtTemp2.Columns.Add("單位");
            dtTemp2.Columns.Add("需求片數");
            dtTemp2.Columns.Add("成品庫存");
            dtTemp2.Columns.Add("外銷庫存");
            dtTemp2.Columns.Add("生產批量");
            dtTemp2.Columns.Add("預計生產批量");
        }


        #region FUNCTION
        public void Search()
        {
            StringBuilder TD001 = new StringBuilder();
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

            dtTemp.Clear();
            dtTemp2.Clear();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@"  SELECT 客戶,日期,品號,品名,規格,CONVERT(INT,SUM(訂單數量)) AS 訂單數量,單位");
                sbSql.Append(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20001' AND LA001=品號) AS '成品倉庫存'");
                sbSql.Append(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20002' AND LA001=品號) AS '外銷倉庫存'");
                sbSql.Append(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(TA015-TA017-TA018),0)) FROM [TK].dbo.MOCTA  WITH (NOLOCK) WHERE TA011 NOT IN ('Y','y') AND TA006=品號 ) AS '未完成的製令' ");
                sbSql.Append(@"  ,(SELECT CONVERT(INT,ISNULL(MC004,0))  FROM [TK].dbo.BOMMC WHERE MC001=品號) AS 標準批量");
                sbSql.Append(@"  FROM (");
                sbSql.Append(@"  SELECT  TC053  AS '客戶' ,TD013 AS '日期',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格'");
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
                //sbSql.Append(@"  AND ( TD004 LIKE '40109916000740%'  ) ");
                sbSql.Append(@"  ) AS TEMP");
                sbSql.Append(@"  GROUP  BY 客戶,日期,品號,品名,規格,單位");
                sbSql.Append(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    label14.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        label14.Text = "有 " + ds.Tables["TEMPds1"].Rows.Count.ToString() + " 筆";
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();

                        SEARCHCOOKIES();
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

            for (int i = 0; i < ds.Tables["TEMPds1"].Rows.Count; i++)
            {

                COPNum = Convert.ToDecimal(ds.Tables["TEMPds1"].Rows[i]["訂單數量"].ToString());
                MB003 = ds.Tables["TEMPds1"].Rows[i]["規格"].ToString();
                sArray = MB003.Split('g');
                //TOTALCOPNum = Convert.ToDecimal(Convert.ToDecimal(sArray[0].ToString())* COPNum);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MD003,MB002,CASE WHEN ISNULL(MB003,'')=''  THEN '1' ELSE MB003 END AS MB003,MD004,MD006 ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE MD003=MB001");
                sbSql.AppendFormat(@"  AND MD003 LIKE '3%' AND MB002 NOT LIKE '%水麵%'   AND  MB002 NOT LIKE '%餅麩%'");
                sbSql.AppendFormat(@"  AND MD001='{0}'" , ds.Tables["TEMPds1"].Rows[i]["品號"].ToString());
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();

                if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                {
                    foreach (DataRow od2 in ds2.Tables["TEMPds2"].Rows)
                    {
                        DataRow row = dtTemp.NewRow();
                        //row["MD001"] = od2["MC001"].ToString();
                        row["DATE"] = ds.Tables["TEMPds1"].Rows[i]["日期"].ToString();
                        row["MD003"] = od2["MD003"].ToString();
                        row["MB002"] = od2["MB002"].ToString();
                        COOKIES = Convert.ToDecimal(Regex.Replace(od2["MB003"].ToString(), "[^0-9]", ""));
                        TOTALCOPNum = Convert.ToDecimal(Convert.ToDecimal(od2["MD006"].ToString())*1000 * COPNum);
                        BATCH= Convert.ToDecimal(ds.Tables["TEMPds1"].Rows[i]["標準批量"].ToString());
                        row["NUM"] = Convert.ToInt32(TOTALCOPNum/ COOKIES/BATCH);
                        
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
            dataGridView2.Rows.Clear();
            dataGridView2.DataSource = dtTemp;
            dataGridView2.AutoResizeColumns();
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
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString()));
                    ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));

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
            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                if(j<=k)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));

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

        private void showwaitfrm()
        {
            PleaseWait objPleaseWait = new PleaseWait();
            objPleaseWait.ShowDialog();
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            //TD = new Thread(showwaitfrm);
            //TD.Start();
            //Thread.Sleep(2000);   //此行可以不需要，主要用於等待主窗體填充數據
            Search();
            //TD.Abort(); //主窗體加載完成數據後，線程結束，關閉等待窗體。

        }

        private void button3_Click(object sender, EventArgs e)
        {

            ExcelExportBOM();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExportCOP();
        }

        #endregion
    }
}
