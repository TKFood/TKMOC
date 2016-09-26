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
        DataGridViewRow drMOCPRODUCTDAILYREPORT = new DataGridViewRow();
        string tablename = null;
        string ID;
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

                sbSql.Append(@"  SELECT  [PRODUCEDEP] AS '製造組',[PRODUCEDATE] AS '日期',[PRODUCEID] AS '製令單號',[PRODUCENAME] AS '品號/品名'");
                sbSql.Append(@" ,[PASTRYPREIN] AS '油酥預計投入量(kg)',[PASTRY] AS '油酥原料',[PASTRYRECYCLE] AS '油酥可回收餅麩'");
                sbSql.Append(@" ,[WATERFLOURPREIN] AS '水麵預計投入量(kg)',[WATERFLOUR] AS '水面原料',[WATERFLOURSIDE] AS '水面可回收邊料'");
                sbSql.Append(@" ,[WATERFLOURRECYCLE] AS '水面可回收餅麩',[PASTRYFLODTIME] AS '油酥、摺疊製造時間(分)',[PASTRYFLODNUM] AS '油酥、摺疊製造人數'");
                sbSql.Append(@" ,[WATERFLOURTIME] AS '水面製造時間(分)',[WATERFLOURNUM] AS '水面製造人數',[RECYCLEFLOUR] AS '可回收餅麩'");
                sbSql.Append(@" ,[KNIFENUM] AS '刀數',[WEIGHTBEFRORE] AS '烤前單片重量(g)',[WEIGHTAFTER] AS '烤後單片重量(g)'");
                sbSql.Append(@" ,[ROWNUM] AS '每排數量',[RECOOKTIME] AS '重烤重工時間',[NGTOTAL] AS '未熟總量(kg)',[NGCOOKTIME] AS '未熟烤焙時間(分)'");
                sbSql.Append(@" ,[PREOUT] AS '預計產出(kg)',[PACKAGETIME] AS '包裝時間(內包裝區/罐裝)(分)',[PACKAGENUM] AS '包裝人數'");
                sbSql.Append(@" ,[STIR] AS '攪拌',[SIDES] AS '成型邊料(kg)',[COOKIES] AS '餅麩(kg)',[COOK] AS '烤焙(kg)',[NGPACKAGE] AS '包裝不良餅乾(kg)'");
                sbSql.Append(@" ,[NGPACKAGECAN] AS '包裝(內袋(卷) 罐)',[CAN] AS '包裝投入(袋(卷),罐)',[WEIGHTCAN] AS '一箱裸餅重'");
                sbSql.Append(@" ,[WEIGHTCANBOXED] AS '一箱餅含袋重',[HLAFWEIGHT] AS '半成品入庫數(kg) (含袋重)',[REMARK] AS '備註'");
                sbSql.Append(@" ,[MANUTIME] AS '製造工時(分)',[PACKTIME] AS '包裝工時(分)',[WEIGHTBEFORECOOK] AS '烤前實際總投入 (kg)'");
                sbSql.Append(@" ,[WEIGHTAFTERCOOK] AS '烤後實際總投入 (kg)',[ACTUALOUT] AS '實際產出(kg)(裸餅)',[WEIGHTPACKAGE] AS '袋重(kg)'");
                sbSql.Append(@" ,[PACKLOST] AS '包裝損耗率',[HLAFLOST] AS '半成品產出效率',[REWORKPCT] AS '重工佔比',[TOTALTIME] AS '總工時(分)'");
                sbSql.Append(@" ,[STIRPCT] AS '攪拌成型製成率%',[EVARATE] AS '蒸發率',[MANULOST] AS '製成損失率',[PCT] AS '製成率'");
                sbSql.Append(@" ,[PRETIME] AS '前置時間',[STOPTIME] AS '停機時間',[ID]");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT]  WITH (NOLOCK)");
                sbSql.AppendFormat(@" WHERE [PRODUCEDATE] ='{0}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"));
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
        public void ExcelExport()
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
            int k = dataGridView1.Rows.Count;
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
                ws.GetRow(j + 1).CreateCell(9).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString()));
                ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));
                ws.GetRow(j + 1).CreateCell(11).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString()));
                ws.GetRow(j + 1).CreateCell(12).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString()));
                ws.GetRow(j + 1).CreateCell(13).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString()));
                ws.GetRow(j + 1).CreateCell(14).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString()));
                ws.GetRow(j + 1).CreateCell(15).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[15].ToString()));
                ws.GetRow(j + 1).CreateCell(16).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[16].ToString()));
                ws.GetRow(j + 1).CreateCell(17).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[17].ToString()));
                ws.GetRow(j + 1).CreateCell(18).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[18].ToString()));
                ws.GetRow(j + 1).CreateCell(19).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[19].ToString()));
                ws.GetRow(j + 1).CreateCell(20).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[20].ToString()));
                ws.GetRow(j + 1).CreateCell(21).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[21].ToString()));
                ws.GetRow(j + 1).CreateCell(22).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[22].ToString()));
                ws.GetRow(j + 1).CreateCell(23).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[23].ToString()));
                ws.GetRow(j + 1).CreateCell(24).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[24].ToString()));
                ws.GetRow(j + 1).CreateCell(25).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[25].ToString()));
                ws.GetRow(j + 1).CreateCell(26).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[26].ToString()));
                ws.GetRow(j + 1).CreateCell(27).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[27].ToString()));
                ws.GetRow(j + 1).CreateCell(28).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[28].ToString()));
                ws.GetRow(j + 1).CreateCell(29).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[29].ToString()));
                ws.GetRow(j + 1).CreateCell(30).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[30].ToString()));
                ws.GetRow(j + 1).CreateCell(31).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[31].ToString()));
                ws.GetRow(j + 1).CreateCell(32).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[32].ToString()));
                ws.GetRow(j + 1).CreateCell(33).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[33].ToString()));
                ws.GetRow(j + 1).CreateCell(34).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[34].ToString()));
                ws.GetRow(j + 1).CreateCell(35).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[35].ToString());
                ws.GetRow(j + 1).CreateCell(36).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[36].ToString()));
                ws.GetRow(j + 1).CreateCell(37).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[37].ToString()));
                ws.GetRow(j + 1).CreateCell(38).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[38].ToString()));
                ws.GetRow(j + 1).CreateCell(39).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[39].ToString()));
                ws.GetRow(j + 1).CreateCell(40).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[40].ToString()));
                ws.GetRow(j + 1).CreateCell(41).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[41].ToString()));
                ws.GetRow(j + 1).CreateCell(42).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[42].ToString()));
                ws.GetRow(j + 1).CreateCell(43).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[43].ToString()));
                ws.GetRow(j + 1).CreateCell(44).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[44].ToString()));
                ws.GetRow(j + 1).CreateCell(45).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[45].ToString()));
                ws.GetRow(j + 1).CreateCell(46).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[46].ToString()));
                ws.GetRow(j + 1).CreateCell(47).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[47].ToString());

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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count >= 1)
            {
                drMOCPRODUCTDAILYREPORT = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];

                ID = dataGridView1.CurrentRow.Cells["ID"].Value.ToString();
                textID.Text = ID;
                dateTimePicker2.Value = Convert.ToDateTime(drMOCPRODUCTDAILYREPORT.Cells["日期"].Value.ToString());
               
                comboBox1.Text= drMOCPRODUCTDAILYREPORT.Cells["製造組"].Value.ToString();
                textBox3.Text = drMOCPRODUCTDAILYREPORT.Cells["製令單號"].Value.ToString(); 
                textBox4.Text = drMOCPRODUCTDAILYREPORT.Cells["品號/品名"].Value.ToString();
                textBox5.Text = drMOCPRODUCTDAILYREPORT.Cells["油酥預計投入量(kg)"].Value.ToString();
                textBox6.Text = drMOCPRODUCTDAILYREPORT.Cells["油酥原料"].Value.ToString();
                textBox7.Text = drMOCPRODUCTDAILYREPORT.Cells["油酥可回收餅麩"].Value.ToString();
                textBox8.Text = drMOCPRODUCTDAILYREPORT.Cells["水麵預計投入量(kg)"].Value.ToString();
                textBox9.Text = drMOCPRODUCTDAILYREPORT.Cells["水面原料"].Value.ToString();
                textBox10.Text = drMOCPRODUCTDAILYREPORT.Cells["水面可回收邊料"].Value.ToString();
                textBox11.Text = drMOCPRODUCTDAILYREPORT.Cells["水面可回收餅麩"].Value.ToString();
                textBox12.Text = drMOCPRODUCTDAILYREPORT.Cells["油酥、摺疊製造時間(分)"].Value.ToString();
                textBox13.Text = drMOCPRODUCTDAILYREPORT.Cells["油酥、摺疊製造人數"].Value.ToString();
                textBox14.Text = drMOCPRODUCTDAILYREPORT.Cells["水面製造時間(分)"].Value.ToString();
                textBox15.Text = drMOCPRODUCTDAILYREPORT.Cells["水面製造人數"].Value.ToString();
                textBox16.Text = drMOCPRODUCTDAILYREPORT.Cells["可回收餅麩"].Value.ToString();
                textBox17.Text = drMOCPRODUCTDAILYREPORT.Cells["刀數"].Value.ToString();
                textBox18.Text = drMOCPRODUCTDAILYREPORT.Cells["烤前單片重量(g)"].Value.ToString();
                textBox19.Text = drMOCPRODUCTDAILYREPORT.Cells["烤後單片重量(g)"].Value.ToString();
                textBox20.Text = drMOCPRODUCTDAILYREPORT.Cells["每排數量"].Value.ToString();
                textBox21.Text = drMOCPRODUCTDAILYREPORT.Cells["重烤重工時間"].Value.ToString();
                textBox22.Text = drMOCPRODUCTDAILYREPORT.Cells["未熟總量(kg)"].Value.ToString();
                textBox23.Text = drMOCPRODUCTDAILYREPORT.Cells["未熟烤焙時間(分)"].Value.ToString();
                textBox24.Text = drMOCPRODUCTDAILYREPORT.Cells["預計產出(kg)"].Value.ToString();
                textBox25.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝時間(內包裝區/罐裝)(分)"].Value.ToString();
                textBox26.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝人數"].Value.ToString();
                textBox27.Text = drMOCPRODUCTDAILYREPORT.Cells["攪拌"].Value.ToString();
                textBox28.Text = drMOCPRODUCTDAILYREPORT.Cells["成型邊料(kg)"].Value.ToString();
                textBox29.Text = drMOCPRODUCTDAILYREPORT.Cells["餅麩(kg)"].Value.ToString();
                textBox30.Text = drMOCPRODUCTDAILYREPORT.Cells["烤焙(kg)"].Value.ToString();
                textBox31.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝不良餅乾(kg)"].Value.ToString();
                textBox32.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝(內袋(卷) 罐)"].Value.ToString();
                textBox33.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝投入(袋(卷),罐)"].Value.ToString();
                textBox34.Text = drMOCPRODUCTDAILYREPORT.Cells["一箱裸餅重"].Value.ToString();
                textBox35.Text = drMOCPRODUCTDAILYREPORT.Cells["一箱餅含袋重"].Value.ToString();
                textBox36.Text = drMOCPRODUCTDAILYREPORT.Cells["半成品入庫數(kg) (含袋重)"].Value.ToString();
                textBox37.Text = drMOCPRODUCTDAILYREPORT.Cells["備註"].Value.ToString();
                textBox38.Text = drMOCPRODUCTDAILYREPORT.Cells["製造工時(分)"].Value.ToString();
                textBox39.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝工時(分)"].Value.ToString();
                textBox40.Text = drMOCPRODUCTDAILYREPORT.Cells["烤前實際總投入 (kg)"].Value.ToString();
                textBox41.Text = drMOCPRODUCTDAILYREPORT.Cells["烤後實際總投入 (kg)"].Value.ToString();
                textBox42.Text = drMOCPRODUCTDAILYREPORT.Cells["實際產出(kg)(裸餅)"].Value.ToString();
                textBox43.Text = drMOCPRODUCTDAILYREPORT.Cells["袋重(kg)"].Value.ToString();
                textBox44.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝損耗率"].Value.ToString();
                textBox45.Text = drMOCPRODUCTDAILYREPORT.Cells["半成品產出效率"].Value.ToString();
                textBox46.Text = drMOCPRODUCTDAILYREPORT.Cells["重工佔比"].Value.ToString();
                textBox47.Text = drMOCPRODUCTDAILYREPORT.Cells["總工時(分)"].Value.ToString();
                textBox48.Text = drMOCPRODUCTDAILYREPORT.Cells["攪拌成型製成率%"].Value.ToString();
                textBox49.Text = drMOCPRODUCTDAILYREPORT.Cells["蒸發率"].Value.ToString();
                textBox50.Text = drMOCPRODUCTDAILYREPORT.Cells["製成損失率"].Value.ToString();
                textBox51.Text = drMOCPRODUCTDAILYREPORT.Cells["製成率"].Value.ToString();
                textBox52.Text = drMOCPRODUCTDAILYREPORT.Cells["前置時間"].Value.ToString();
                textBox53.Text = drMOCPRODUCTDAILYREPORT.Cells["停機時間"].Value.ToString();

                //numericUpDown1.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["預計投入量(kg)"].Value.ToString());



            }

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
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] ");
                //sbSql.AppendFormat("  SET [PRODUCEDATE]='{1}',[PRODUCEID]='{2}',[PRODUCENAME]='{3}',[PREIN]='{4}',[PROCESSTIME]='{5}',[PRECESSPEOPEL]='{6}',[MATERIAL]='{7}',[RECYCLEMATERIAL]='{8}',[MATERIALTIME]='{9}',[MATERIALPEOPLE]='{10}',[WMATERIAL]='{11}',[WRECYCLESIDE]='{12}',[WRECYCLECOOKIES]='{13}',[PACKAGETIME]='{14}',[PACKAGEPEOPLE]='{15}',[TRECYCLE]='{16}',[NGTATOL]='{17}',[NGTIME]='{18}',[WEIGHTBEFORE]='{19}',[WEIGHTAFTER]='{20}',[COOKIENUM]='{21}',[BLADENUM]='{22}',[INBEFORE]='{23}',[INAFTER]='{24}',[PREOUT]='{25}',[NGSTIR]='{26}',[NGSIDE]='{27}',[NGCOOKIES]='{28}',[NGBAKE]='{29}',[NGNOGOOD]='{30}',[NGNOCAN]='{31}',[PACKAGEWEIGHT]='{32}',[PACKAGEIN]='{33}',[ACTUALOUT]='{34}',[HALFOUT]='{35}',[REMARK]='{36}',[MANUTIME]='{37}',[STIRPCT]='{38}',[EVARATE]='{39}',[LOSTRATE]='{40}',[HALFRATE]='{41}',[PACKAGETOTALTIME]='{42}',[WEIGHTTOTAL]='{43}',[PACKAGELOSTRATE]='{44}',[TOTALRATE]='{45}',[CANCOOKIESWEIGHT]='{46}',[CANWEIGHT] ='{47}' WHERE [ID]='{0}' ", textID.Text.ToString(), dateTimePicker2.Value.ToString("yyyy/MM/dd"), textBox3.Text.ToString(), textBox4.Text.ToString(), numericUpDown1.Value, numericUpDown2.Value, numericUpDown3.Value, numericUpDown4.Value, numericUpDown5.Value, numericUpDown6.Value, numericUpDown7.Value, numericUpDown8.Value, numericUpDown9.Value, numericUpDown10.Value, numericUpDown11.Value, numericUpDown12.Value, numericUpDown13.Value, numericUpDown14.Value, numericUpDown15.Value, numericUpDown16.Value, numericUpDown17.Value, numericUpDown18.Value, numericUpDown19.Value, numericUpDown20.Value, numericUpDown21.Value, numericUpDown22.Value, numericUpDown23.Value, numericUpDown24.Value, numericUpDown25.Value, numericUpDown26.Value, numericUpDown27.Value, numericUpDown28.Value, numericUpDown29.Value, numericUpDown30.Value, numericUpDown31.Value, numericUpDown32.Value, textBox5.Text.ToString(), numericUpDown33.Value, numericUpDown34.Value, numericUpDown35.Value, numericUpDown36.Value, numericUpDown37.Value, numericUpDown38.Value, numericUpDown39.Value, numericUpDown40.Value, numericUpDown41.Value, numericUpDown42.Value, numericUpDown43.Value);
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
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] ");
                sbSql.Append(" ( )  ");
                //sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}','{30}','{31}','{32}','{33}','{34}','{35}','{36}','{37}','{38}','{39}','{40}','{41}','{42}','{43}','{44}','{45}','{46}','{47}') ", Guid.NewGuid(), dateTimePicker2.Value.ToString("yyyy/MM/dd"),textBox3.Text.ToString(),textBox4.Text.ToString(),numericUpDown1.Value, numericUpDown2.Value, numericUpDown3.Value, numericUpDown4.Value, numericUpDown5.Value, numericUpDown6.Value, numericUpDown7.Value, numericUpDown8.Value, numericUpDown9.Value, numericUpDown10.Value, numericUpDown11.Value, numericUpDown12.Value, numericUpDown13.Value, numericUpDown14.Value, numericUpDown15.Value, numericUpDown16.Value, numericUpDown17.Value, numericUpDown18.Value, numericUpDown19.Value, numericUpDown20.Value, numericUpDown21.Value, numericUpDown22.Value, numericUpDown23.Value, numericUpDown24.Value, numericUpDown25.Value, numericUpDown26.Value, numericUpDown27.Value, numericUpDown28.Value, numericUpDown29.Value, numericUpDown30.Value, numericUpDown31.Value, numericUpDown32.Value,textBox5.Text.ToString(), numericUpDown33.Value, numericUpDown34.Value, numericUpDown35.Value, numericUpDown36.Value, numericUpDown37.Value, numericUpDown38.Value, numericUpDown39.Value, numericUpDown40.Value, numericUpDown41.Value, numericUpDown42.Value, numericUpDown43.Value);

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
        
        public void SetADD()
        {
            textBox3.Text = null;
            textBox4.Text = null;
            textID.Text = null;
            //
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
            textBox17.Text = null;
            textBox18.Text = null;
            textBox19.Text = null;
            textBox20.Text = null;
            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
            textBox25.Text = null;
            textBox26.Text = null;
            textBox27.Text = null;
            textBox28.Text = null;
            textBox29.Text = null;
            textBox30.Text = null;
            textBox31.Text = null;
            textBox32.Text = null;
            textBox33.Text = null;
            textBox34.Text = null;
            textBox35.Text = null;
            textBox36.Text = null;
            textBox37.Text = null;
            textBox38.Text = null;
            textBox39.Text = null;
            textBox40.Text = null;
            textBox41.Text = null;
            textBox42.Text = null;
            textBox43.Text = null;
            textBox44.Text = null;
            textBox45.Text = null;
            textBox46.Text = null;
            textBox47.Text = null;
            textBox48.Text = null;
            textBox49.Text = null;
            textBox50.Text = null;
            textBox51.Text = null;
            textBox52.Text = null;
            textBox53.Text = null;



            dateTimePicker2.Value = DateTime.Now;
            //
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            //textID.ReadOnly = false;
           
            dateTimePicker2.Enabled = true;
        }

        public void SetUPDATE()
        {
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            //textID.ReadOnly = false;
            
            dateTimePicker2.Enabled = true;
        }
        public void SetFINISH()
        {
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
            textBox5.ReadOnly = true;
            textID.ReadOnly = true;
           
        }

       

        public void PRINTDOC()
        {
            // 首先把建立的範本檔案讀入MemoryStream
            //首先把建立的範本檔案讀入MemoryStream
            System.IO.MemoryStream _memoryStream = new System.IO.MemoryStream(Properties.Resources.生產日報);

            //建立一個Document物件
            //並傳入MemoryStream
            Aspose.Words.Document doc = new Aspose.Words.Document(_memoryStream);

            //新增一個DataTable
            DataTable table = new DataTable();
            //建立Column
            table.Columns.Add("P11");
            table.Columns.Add("P12");
            table.Columns.Add("P21");
            table.Columns.Add("P22");


            //[APPLYUNIT] AS '申請單位',[APPDATE] AS '申請日期',[EQUIPMENTID] AS '機台編號' 
            //,[EQUIPMENTNAME] AS '設備名稱',[FINDEMP] AS '發現者',[APPLYEMP] AS '申請人' ,[ERROR] AS '異常情形'
            //,[STATUS] AS '原因及處理方式',[REMARK] AS '備註',[MAINEMP] AS '維修者',[MAINDATE] AS '維修時間'
            //透過建立的DataTable物件來New一個儲存資料的Row
            DataRow row = table.NewRow();
            //這些Row具有上面所建立相同的Column欄位
            //因此可以直接指定欄位名稱將資料填入裡面       
            //DateTime dt = Convert.ToDateTime(drMAINAPPLY.Cells["申請日期"].Value.ToString());
            row["P11"] ="A";
            row["P12"] = "B";
            row["P21"] = "C";
            row["P22"] = "D";

            //把所建立的資料行加入Table的Row清單內
            table.Rows.Add(row);


            //將DataTable傳入Document的MailMerge.Execute()方法
            //doc.MailMerge.Execute(table);
            //清空所有未被合併的功能變數
            doc.MailMerge.DeleteFields();

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            //將檔案儲存至c:\
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\生產日報{0}.doc", DateTime.Now.ToString("yyyyMMdd"));
            doc.Save(filename.ToString());

            MessageBox.Show("匯出完成-文件放在-" + filename.ToString());
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

            SetFINISH();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExcelExport();
            ExcelExport();
            //PRINTDOC();
        }















        #endregion
    }
}
