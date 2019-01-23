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
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using System.Collections;

namespace TKMOC
{
    public partial class frmCOPMOCCHECK : Form
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
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();
        SqlDataAdapter adapter8 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder8 = new SqlCommandBuilder();
        SqlDataAdapter adapter9 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder9 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();
        DataSet ds9 = new DataSet();
        DataTable dt = new DataTable();
        DataTable dtTemp = new DataTable();
        DataTable dtTemp2 = new DataTable();
        int result;
        string tablename = null;
        decimal COPNum = 0;


        public frmCOPMOCCHECK()
        {
            InitializeComponent();

            NEWdtTemp();
            NEWdtTemp2();
        }

        #region FUNCTION
        public void NEWdtTemp()
        {
            dtTemp.Columns.Add("商品");

            dtTemp.Columns.Add("品號");
            dtTemp.Columns.Add("品名");
            //dtTemp.Columns.Add("數量");
            dtTemp.Columns.Add("單位");

            DataColumn colDecimal = new DataColumn("數量");
            colDecimal.DataType = System.Type.GetType("System.Decimal");
            dtTemp.Columns.Add(colDecimal);
        }

        public void NEWdtTemp2()
        {
            dtTemp2.Columns.Add("品號");
            dtTemp2.Columns.Add("品名");
            //dtTemp.Columns.Add("數量");
            dtTemp2.Columns.Add("規格");
            dtTemp2.Columns.Add("庫存單位");

            DataColumn colDecimal = new DataColumn("需求數量");
            colDecimal.DataType = System.Type.GetType("System.Decimal");
            dtTemp2.Columns.Add(colDecimal);

            DataColumn colDecimal2 = new DataColumn("庫存量");
            colDecimal2.DataType = System.Type.GetType("System.Decimal");
            dtTemp2.Columns.Add(colDecimal2);

            DataColumn colDecimal3 = new DataColumn("需求量比較");
            colDecimal2.DataType = System.Type.GetType("System.Decimal");
            dtTemp2.Columns.Add(colDecimal3);

            //DataColumn colDecimal4 = new DataColumn("預計採購量");
            //colDecimal2.DataType = System.Type.GetType("System.Decimal");
            //dtTemp2.Columns.Add(colDecimal4);
        }


        public void Search()
        {
           

            StringBuilder TD001 = new StringBuilder();
            StringBuilder TC027 = new StringBuilder();
            StringBuilder QUERY1 = new StringBuilder();

            if (checkBox1.Checked == true)
            {
                TD001.AppendFormat("'A221',");
            }
            if (checkBox2.Checked == true)
            {
                TD001.AppendFormat("'A222',");
            }
            if (checkBox3.Checked == true)
            {
                TD001.AppendFormat("'A228',");
            }
            if (checkBox4.Checked == true)
            {
                TD001.AppendFormat("'A225',");
            }
            if (checkBox5.Checked == true)
            {
                TD001.AppendFormat("'A226',");
            }
            if (checkBox6.Checked == true)
            {
                TD001.AppendFormat("'A227',");
            }
            if (checkBox7.Checked == true)
            {
                TD001.AppendFormat("'A223',");
            }
            if (checkBox8.Checked == true)
            {
                TD001.AppendFormat("'A229',");
            }
            TD001.AppendFormat("''");

            if (comboBox1.Text.ToString().Equals("已確認"))
            {
                TC027.AppendFormat(" 'Y',");
            }
            else if (comboBox1.Text.ToString().Equals("未確認"))
            {
                TC027.AppendFormat(" 'N',");
            }
            else if (comboBox1.Text.ToString().Equals("全部"))
            {
                TC027.AppendFormat(" 'Y','N', ");
            }

            TC027.AppendFormat("''");
            
          

            dtTemp.Clear();
            dtTemp2.Clear();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT 日期,品號,品名,客戶,CONVERT(INT,SUM(訂單數量)) AS 訂單數量,單位 ,單別,單號,序號,規格  ");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20001' AND LA001=品號) AS '成品倉庫存'");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20002' AND LA001=品號) AS '外銷倉庫存'");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(TA015-TA017-TA018),0)) FROM [TK].dbo.MOCTA  WITH (NOLOCK) WHERE TA011 NOT IN ('Y','y') AND TA006=品號 ) AS '未完成的製令' ");
                sbSql.AppendFormat(@"  ,(SELECT CONVERT(INT,ISNULL(MC004,0))  FROM [TK].dbo.BOMMC WHERE MC001=品號) AS 標準批量");
                sbSql.AppendFormat(@"  FROM (");
                sbSql.AppendFormat(@"  SELECT   TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TC053  AS '客戶' ,TD013 AS '日期',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*MD004 END) AS '訂單數量'");
                sbSql.AppendFormat(@"  ,MB004 AS '單位'");
                sbSql.AppendFormat(@"  ,((TD008-TD009)+(TD024-TD025)) AS '訂單量'");
                sbSql.AppendFormat(@"  ,TD010 AS '訂單單位' ");
                sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(MD002,'')<>'' THEN MD002 ELSE TD010 END ) AS '換算單位'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MD003>0 THEN MD003 ELSE 1 END) AS '分子'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MD004>0 THEN MD004 ELSE (TD008-TD009) END ) AS '分母'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.AppendFormat(@"  WHERE TD004=MB001");
                sbSql.AppendFormat(@"  AND TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@"  AND (TD004 LIKE '401%' OR TD004 LIKE '402%' OR TD004 LIKE '403%' OR TD004 LIKE '404%' OR TD004 LIKE '405%' OR TD004 LIKE '406%' OR TD004 LIKE '407%'   ) ");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TC001 IN ({0}) ", TD001.ToString());
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@" AND TC027 IN ({0})  ", TC027.ToString());
                sbSql.AppendFormat(@"  {0}", QUERY1.ToString());
                //sbSql.AppendFormat(@"  AND ( TD004 LIKE '40102910540200%'  ) ");
                //sbSql.AppendFormat(@"  AND ( TD002='20181211001'  ) ");
                sbSql.AppendFormat(@") AS TEMP");
                sbSql.AppendFormat(@"  GROUP  BY 客戶,日期,品號,品名,規格,單位,單別,單號,序號");
                sbSql.AppendFormat(@"  ORDER BY 日期,客戶,品號,品名,規格,單位,單別,單號,序號 ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                //if (CHECKYN.Equals("N"))
                //{
                //    //建立一個DataGridView的Column物件及其內容
                //    DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
                //    dgvc.Width = 40;
                //    dgvc.Name = "選取";

                //    this.dataGridView1.Columns.Insert(0, dgvc);
                //    CHECKYN = "Y";
                //}


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    //dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        //dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        //dataGridView1.AutoResizeColumns();


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

        public void SETCOPTHGROUPBY()
        {
            try
            {
                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    var query = from t in ds.Tables["TEMPds1"].AsEnumerable()
                                group t by new { t1 = t.Field<string>("品號") } into m
                                select new
                                {
                                    MB001 = m.Key.t1,
                                    SUM = m.Sum(n => n.Field<int>("訂單數量"))
                                };
                    if (query.ToList().Count > 0)
                    {
                        string MB001 = null;
                        string MB003 = null;
                        string[] sArray = null;
                        connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                        sqlConn = new SqlConnection(connectionString);

                        dtTemp.Clear();

                        query.ToList().ForEach(q =>
                        {

                            COPNum = Convert.ToDecimal(q.SUM);
                            MB001 = q.MB001.ToString();

                            //TOTALCOPNum = Convert.ToDecimal(Convert.ToDecimal(sArray[0].ToString())* COPNum);

                            sbSql.Clear();
                            sbSqlQuery.Clear();

                            sbSql.AppendFormat(@"  WITH TEMPTABLE (MD001,MD003,MD004,MD006,MD007,MD008,MC004,NUM,LV) AS");
                            sbSql.AppendFormat(@"  (");
                            sbSql.AppendFormat(@"  SELECT  MD001,MD003,MD004,MD006,MD007,MD008,MC004,CONVERT(decimal(18,5),(MD006*(1+MD008)/MD007)/MC004) AS NUM,1 AS LV FROM [TK].dbo.VBOMMD WHERE  MD001='{0}'", MB001);
                            sbSql.AppendFormat(@"  UNION ALL");
                            sbSql.AppendFormat(@"  SELECT A.MD001,A.MD003,A.MD004,A.MD006,A.MD007,A.MD008,A.MC004,CONVERT(decimal(18,5),(A.MD006*(1+A.MD008)/A.MD007/A.MC004)*(B.NUM)) AS NUM,LV+1");
                            sbSql.AppendFormat(@"  FROM [TK].dbo.VBOMMD A");
                            sbSql.AppendFormat(@"  INNER JOIN TEMPTABLE B on A.MD001=B.MD003");
                            sbSql.AppendFormat(@"  )");
                            sbSql.AppendFormat(@"  SELECT MD001,MD003,MD004,MD006,MD007,MD008,MC004,NUM,LV,MB002");
                            sbSql.AppendFormat(@"  FROM TEMPTABLE ");
                            sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON MB001=MD003");

                            sbSql.AppendFormat(@"  WHERE  (MD003 LIKE '301%') ");
                            //sbSql.AppendFormat(@"  WHERE  MD003='203022061' ");
                            sbSql.AppendFormat(@"  ORDER BY LV,MD001,MD003");



                            adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                            sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                            sqlConn.Open();
                            ds2.Clear();
                            adapter2.Fill(ds2, "TEMPds2");
                            sqlConn.Close();

                            if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                            {

                                foreach (DataRow od2 in ds2.Tables["TEMPds2"].Rows)
                                {
                                    DataRow row = dtTemp.NewRow();
                                    //row["MD001"] = od2["MC001"].ToString();

                                    row["商品"] = MB001.ToString();
                                    row["品號"] = od2["MD003"].ToString();
                                    row["品名"] = od2["MB002"].ToString();
                                    row["數量"] = Convert.ToDecimal(COPNum) * Convert.ToDecimal(od2["NUM"].ToString());
                                    row["單位"] = od2["MD004"].ToString();

                                    dtTemp.Rows.Add(row);


                                }

                            }

                        }
                        );
                    }



                    //query.ToList().ForEach(q =>
                    //{
                    //    //MessageBox.Show(q.MB001 + "," );
                    //    MessageBox.Show(q.MB001 + "," + q.SUM);
                    //});
                }

                if (dtTemp.Rows.Count > 0)
                {
                    dataGridView3.DataSource = dtTemp;
                    dataGridView3.AutoResizeColumns();

                    SETMOCGROUPBY();
                }
            }
            catch
            {

            }
            finally
            {

            }


        }

        public void SETMOCGROUPBY()
        {
            DateTime dt = DateTime.Now;


            if (dtTemp.Rows.Count >= 1)
            {
                var query = from t in dtTemp.AsEnumerable()
                            group t by new { t1 = t.Field<string>("品號") } into m
                            orderby m.Key.t1
                            select new
                            {
                                MB001 = m.Key.t1,
                                SUM = m.Sum(n => n.Field<decimal>("數量"))

                            }
                            ;

                if (query.ToList().Count > 0)
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    dtTemp2.Clear();

                    query.ToList().ForEach(q =>
                    {

                        string LA001 = q.MB001;
                        sbSql.Clear();
                        sbSqlQuery.Clear();

                        sbSql.AppendFormat(@" SELECT MB001 AS '品號',ISNULL(SUM(LA011*LA005),0) AS '庫存量',MB002 AS '品名',MB003 AS '規格',MB004 AS '庫存單位'   ");                        sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB   ");
                        sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVLA ON LA001=MB001 AND  LA009 IN ('20005')");
                        sbSql.AppendFormat(@" WHERE  MB001='{0}' ", LA001);
                        sbSql.AppendFormat(@" GROUP BY MB001,LA001,MB002,MB003,MB004 ");
                        sbSql.AppendFormat(@"  ");


                        adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                        sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                        sqlConn.Open();
                        ds3.Clear();
                        adapter3.Fill(ds3, "TEMPds3");
                        sqlConn.Close();

                        if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                        {

                            foreach (DataRow od3 in ds3.Tables["TEMPds3"].Rows)
                            {
                                DataRow row = dtTemp2.NewRow();
                                //row["MD001"] = od2["MC001"].ToString();

                                row["品號"] = od3["品號"].ToString();
                                row["品名"] = od3["品名"].ToString();
                                row["規格"] = od3["規格"].ToString();
                                row["庫存單位"] = od3["庫存單位"].ToString();
                                row["需求數量"] = Convert.ToDecimal(q.SUM.ToString());
                                row["庫存量"] = Convert.ToDecimal(od3["庫存量"].ToString());
                                row["需求量比較"] = Convert.ToDecimal(od3["庫存量"].ToString()) - Convert.ToDecimal(q.SUM.ToString());
                   

                                dtTemp2.Rows.Add(row);
                            }

                        }



                    }
                    );
                }


            }
            //dataGridView2.DataSource = query.ToList();
            //dataGridView2.AutoResizeColumns();

            dataGridView2.DataSource = dtTemp2;
            dataGridView2.AutoResizeColumns();

            //根据列表中数据不同，显示不同颜色背景
            foreach (DataGridViewRow dgRow in dataGridView2.Rows)
            {
                //判断
                if (Convert.ToDecimal(dgRow.Cells[6].Value) < 0)
                {
                    //将这行的背景色设置成Pink
                    dgRow.DefaultCellStyle.BackColor = Color.Pink;
                }
            }

        }


        public void ExcelExport()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            ws = wb.CreateSheet("Sheet1");
            ws.CreateRow(0);
            //第一行為欄位名稱
            ws.GetRow(0).CreateCell(0).SetCellValue("品號");
            ws.GetRow(0).CreateCell(1).SetCellValue("品名");
            ws.GetRow(0).CreateCell(2).SetCellValue("規格");
            ws.GetRow(0).CreateCell(3).SetCellValue("庫存單位");
            ws.GetRow(0).CreateCell(4).SetCellValue("需求數量");
            ws.GetRow(0).CreateCell(5).SetCellValue("庫存量");
            ws.GetRow(0).CreateCell(6).SetCellValue("需求量比較");
            ws.GetRow(0).CreateCell(7).SetCellValue("預計採購量");




            int j = 0;
            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(dr.Cells[0].Value.ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(dr.Cells[1].Value.ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(dr.Cells[2].Value.ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(dr.Cells[3].Value.ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(dr.Cells[4].Value.ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(dr.Cells[5].Value.ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(dr.Cells[6].Value.ToString());
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(dr.Cells[7].Value.ToString());

                j++;
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
            filename.AppendFormat(@"c:\temp\資料{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                string MB001 = null;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    MB001 = row.Cells["品號"].Value.ToString();

                    //MessageBox.Show(MB001.ToString());
                    Search3(MB001.ToString());
                }
                else
                {
                    MB001 = null;

                }
            }
        }
        public void Search3(string MB001)
        {
        

            StringBuilder TD001 = new StringBuilder();
            StringBuilder TC027 = new StringBuilder();
            StringBuilder QUERY1 = new StringBuilder();

            if (checkBox1.Checked == true)
            {
                TD001.AppendFormat("'A221',");
            }
            if (checkBox2.Checked == true)
            {
                TD001.AppendFormat("'A222',");
            }
            if (checkBox3.Checked == true)
            {
                TD001.AppendFormat("'A228',");
            }
            if (checkBox4.Checked == true)
            {
                TD001.AppendFormat("'A225',");
            }
            if (checkBox5.Checked == true)
            {
                TD001.AppendFormat("'A226',");
            }
            if (checkBox6.Checked == true)
            {
                TD001.AppendFormat("'A227',");
            }
            if (checkBox7.Checked == true)
            {
                TD001.AppendFormat("'A223',");
            }
            if (checkBox8.Checked == true)
            {
                TD001.AppendFormat("'A229',");
            }
            TD001.AppendFormat("''");

            if (comboBox1.Text.ToString().Equals("已確認"))
            {
                TC027.AppendFormat(" 'Y',");
            }
            else if (comboBox1.Text.ToString().Equals("未確認"))
            {
                TC027.AppendFormat(" 'N',");
            }
            else if (comboBox1.Text.ToString().Equals("全部"))
            {
                TC027.AppendFormat(" 'Y','N', ");
            }
            TC027.AppendFormat("''");


            QUERY1.AppendFormat(" AND ((TD004 IN (SELECT [MD001] FROM [TK].[dbo].[VBOMMD] WHERE [MD003]='{0}')) OR (TD004 IN (SELECT MD001 FROM [TK].[dbo].[VBOMMD] WHERE MD003 IN (SELECT MD001 FROM [TK].[dbo].[VBOMMD] WHERE [MD003]='{0}' ))) OR (TD004 IN (SELECT MD001 FROM [TK].[dbo].[BOMMD]  WHERE MD003 IN ( SELECT MD001 FROM [TK].[dbo].[BOMMD]  WHERE MD003 IN ( SELECT MD001 FROM [TK].[dbo].[BOMMD] WHERE [MD003]='{0}' ))) ) ) ", MB001.ToString());


            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT 日期,品號,品名,客戶,CONVERT(INT,SUM(訂單數量)) AS 訂單數量,單位 ,單別,單號,序號,規格  ");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20001' AND LA001=品號) AS '成品倉庫存'");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20002' AND LA001=品號) AS '外銷倉庫存'");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(TA015-TA017-TA018),0)) FROM [TK].dbo.MOCTA  WITH (NOLOCK) WHERE TA011 NOT IN ('Y','y') AND TA006=品號 ) AS '未完成的製令' ");
                sbSql.AppendFormat(@"  ,(SELECT CONVERT(INT,ISNULL(MC004,0))  FROM [TK].dbo.BOMMC WHERE MC001=品號) AS 標準批量");
                sbSql.AppendFormat(@"  FROM (");
                sbSql.AppendFormat(@"  SELECT   TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TC053  AS '客戶' ,TD013 AS '日期',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*MD004 END) AS '訂單數量'");
                sbSql.AppendFormat(@"  ,MB004 AS '單位'");
                sbSql.AppendFormat(@"  ,((TD008-TD009)+(TD024-TD025)) AS '訂單量'");
                sbSql.AppendFormat(@"  ,TD010 AS '訂單單位' ");
                sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(MD002,'')<>'' THEN MD002 ELSE TD010 END ) AS '換算單位'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MD003>0 THEN MD003 ELSE 1 END) AS '分子'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MD004>0 THEN MD004 ELSE (TD008-TD009) END ) AS '分母'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.AppendFormat(@"  WHERE TD004=MB001");
                sbSql.AppendFormat(@"  AND TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@"  AND (TD004 LIKE '401%' OR TD004 LIKE '402%' OR TD004 LIKE '403%' OR TD004 LIKE '404%' OR TD004 LIKE '405%' OR TD004 LIKE '406%' OR TD004 LIKE '407%'   ) ");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TC001 IN ({0}) ", TD001.ToString());
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@" AND TC027 IN ({0})  ", TC027.ToString());
                sbSql.AppendFormat(@"  {0}", QUERY1.ToString());
                //sbSql.AppendFormat(@"  AND ( TD004 LIKE '40102910540200%'  ) ");
                //sbSql.AppendFormat(@"  AND ( TD002='20180708006'  ) ");
                sbSql.AppendFormat(@") AS TEMP");
                sbSql.AppendFormat(@"  GROUP  BY 客戶,日期,品號,品名,規格,單位,單別,單號,序號");
                sbSql.AppendFormat(@"  ORDER BY 日期,客戶,品號,品名,規格,單位,單別,單號,序號 ");
                sbSql.AppendFormat(@"  ");

                adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                sqlConn.Open();
                ds7.Clear();
                adapter7.Fill(ds7, "TEMPds7");
                sqlConn.Close();



                if (ds7.Tables["TEMPds7"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds7.Tables["TEMPds7"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds7.Tables["TEMPds7"];
                        dataGridView4.AutoResizeColumns();


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

        public void Search8()
        {

            StringBuilder TA013 = new StringBuilder();

            if (comboBox1.Text.ToString().Equals("已確認"))
            {
                TA013.AppendFormat(" 'Y'");
            }
            else if (comboBox1.Text.ToString().Equals("未確認"))
            {
                TA013.AppendFormat(" 'N'");
            }
            else if (comboBox1.Text.ToString().Equals("全部"))
            {
                TA013.AppendFormat("('Y','N' ");
            }
            
          
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

              
                sbSql.AppendFormat(@"  SELECT TB003 AS '品號',TB012 AS '品名',SUM(TB004-TB005) AS '總需求量',SUM(TB004) AS '製令量',SUM(TB005) AS '已領量',TB007 AS '單位'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}' ",dateTimePicker3.Value.ToString("yyyyMMdd"),dateTimePicker4.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TB003 LIKE '301%'");
                sbSql.AppendFormat(@"  AND TA013 IN ({0})", TA013.ToString());
                sbSql.AppendFormat(@"  GROUP BY TB003,TB012,TB007");
                sbSql.AppendFormat(@"  ORDER BY TB003");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter8 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder8 = new SqlCommandBuilder(adapter8);
                sqlConn.Open();
                ds8.Clear();
                adapter8.Fill(ds8, "TEMPds8");
                sqlConn.Close();




                if (ds8.Tables["TEMPds8"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds8.Tables["TEMPds8"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds8.Tables["TEMPds8"];
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
            SETCOPTHGROUPBY();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Search8();
        }
        #endregion


    }
}
