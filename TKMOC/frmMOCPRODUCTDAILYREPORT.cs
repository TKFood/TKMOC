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
using TKITDLL;

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
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        DataGridViewRow drMOCPRODUCTDAILYREPORT = new DataGridViewRow();
        string tablename = null;
        string ID;
        int result;
        string SOURCEID;
        Thread TD;

        public frmMOCPRODUCTDAILYREPORT()
        {
            InitializeComponent();
            tableLayoutPanel1.AutoScroll = true;
            tableLayoutPanel1.AutoScrollMinSize = new Size(1000, 700);
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

        #region FUNCTION
        public void Search()
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT  [PRODUCETYPE] AS '成品/半成品' ,[PRODUCEDEP] AS '製造組',[PRODUCEDATE] AS '日期',[PRODUCEMB001] AS '品號',[PRODUCENAME] AS '品名'");
                sbSql.AppendFormat(@" ,[PASTRYPREIN] AS '油酥預計投入量(kg)',[PASTRY] AS '油酥原料',[PASTRYRECYCLE] AS '油酥可回收餅麩'");
                sbSql.AppendFormat(@" ,[WATERFLOURPREIN] AS '水麵預計投入量(kg)',[WATERFLOUR] AS '水面原料',[WATERFLOURSIDE] AS '水面可回收邊料'");
                sbSql.AppendFormat(@" ,[WATERFLOURRECYCLE] AS '水面可回收餅麩',[PASTRYFLODTIME] AS '油酥、摺疊製造時間(分)',[PASTRYFLODNUM] AS '油酥、摺疊製造人數'");
                sbSql.AppendFormat(@" ,[WATERFLOURTIME] AS '水面製造時間(分)',[WATERFLOURNUM] AS '水面製造人數',[RECYCLEFLOUR] AS '今日產生可回收餅麩'");
                sbSql.AppendFormat(@" ,[KNIFENUM] AS '刀數',[WEIGHTBEFRORE] AS '烤前單片重量(g)',[WEIGHTAFTER] AS '烤後單片重量(g)'");
                sbSql.AppendFormat(@" ,[ROWNUM] AS '每排數量',[NGTOTAL] AS '未熟總量(kg)',[NGCOOKTIME] AS '未熟烤焙時間(分)',[RECOOKTIME] AS '重烤重工時間'");

                sbSql.AppendFormat(@" ,[PREOUT] AS '預計產出(kg)',[PACKAGETIME] AS '包裝時間(內包裝區/罐裝)(分)',[PACKAGENUM] AS '包裝人數'");
                sbSql.AppendFormat(@" ,[STIR] AS '攪拌',[SIDES] AS '成型邊料(kg)',[COOKIES] AS '餅麩(kg)',[COOK] AS '篩選餅乾區不良烤焙(kg)'");
                sbSql.AppendFormat(@" ,[OUTCOOKIES] AS '篩選餅乾區餅乾屑(kg)' ,[CLEANCOOKIES] AS '清掃廢料(kg)'  ");
                sbSql.AppendFormat(@" ,[NGPACKAGE] AS '包裝不良餅乾(kg)',[NGPACKAGECAN] AS '包裝(內袋(卷) 罐)',[CAN] AS '包裝投入(袋(卷),罐)',[WEIGHTCAN] AS '一箱裸餅重'");
                sbSql.AppendFormat(@" ,[WEIGHTCANBOXED] AS '一箱餅含袋重',[HLAFWEIGHT] AS '半成品入庫數(kg) (含袋重)',[REMARK] AS '備註'");
                sbSql.AppendFormat(@" ,[MANUTIME] AS '製造工時(分)',[PACKTIME] AS '包裝工時(分)',[WEIGHTBEFORECOOK] AS '預計投入量'");
                sbSql.AppendFormat(@" ,[WEIGHTAFTERCOOK] AS '烤後實際總投入 (kg)',[ACTUALOUT] AS '實際產出(kg)(裸餅)',[WEIGHTPACKAGE] AS '袋重(kg)'");
                sbSql.AppendFormat(@" ,[PACKLOST] AS '包裝損耗率',[HLAFLOST] AS '半成品產出效率',[REWORKPCT] AS '重工佔比',[TOTALTIME] AS '總工時(分)'");
                sbSql.AppendFormat(@" ,[STIRPCT] AS '攪拌成型製成率%',[EVARATE] AS '蒸發率',[MANULOST] AS '製成損失率',[PCT] AS '製成率'");
                sbSql.AppendFormat(@" ,[PRETIME] AS '前置時間',[STOPTIME] AS '停機時間'");
                sbSql.AppendFormat(@" ,[PREWEIGT] AS '容量規格',[PRECAN] AS '預計包罐數',[ACTUALCAN] AS '實際包罐數',[TOTALPCT] AS '總製成率'");
                sbSql.AppendFormat(@" ,[CANPCT] AS '總包罐製成率',TRYCAN AS '預計試吃包罐數',ACTUALTRYCAN  AS '實際試吃包罐數' ");
                sbSql.AppendFormat(@" ,[ID]");
                sbSql.AppendFormat(@" FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT]  WITH (NOLOCK)");
                sbSql.AppendFormat(@" WHERE [PRODUCEDATE] ='{0}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"));
                //sbSql.AppendFormat(@" WHERE [ID] ='{0}'", ID);
                sbSql.Append(@" ORDER BY [ID]  ");
                sbSql.Append(@"   ");


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
                        LOADDATA1();
                        
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

        public void LOADDATA1()
        {
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.HeaderCell.Style.Font = new Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Pixel);
                
            }
           
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoResizeColumns();
        }
        public void SearchMOCPRODUCTDAILYREPORTPROCESSID()
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [PROCESSID] AS '製令' FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORTPROCESSID] WHERE [SOURCEID]='{0}'", SOURCEID.ToString());
               
                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();

                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns();
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
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
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
                ws.GetRow(j + 1).CreateCell(35).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[35].ToString()));
                ws.GetRow(j + 1).CreateCell(36).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[36].ToString()));
                ws.GetRow(j + 1).CreateCell(37).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[37].ToString());
               
                ws.GetRow(j + 1).CreateCell(38).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[38].ToString()));
                ws.GetRow(j + 1).CreateCell(39).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[39].ToString()));
                ws.GetRow(j + 1).CreateCell(40).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[40].ToString()));
                ws.GetRow(j + 1).CreateCell(41).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[41].ToString()));
                ws.GetRow(j + 1).CreateCell(42).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[42].ToString()));
                ws.GetRow(j + 1).CreateCell(43).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[43].ToString()));
                ws.GetRow(j + 1).CreateCell(44).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[44].ToString()));
                ws.GetRow(j + 1).CreateCell(45).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[45].ToString()));
                ws.GetRow(j + 1).CreateCell(46).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[46].ToString()));
                ws.GetRow(j + 1).CreateCell(47).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[47].ToString()));
                ws.GetRow(j + 1).CreateCell(48).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[48].ToString()));
                ws.GetRow(j + 1).CreateCell(49).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[49].ToString()));
                ws.GetRow(j + 1).CreateCell(50).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[50].ToString()));
                ws.GetRow(j + 1).CreateCell(51).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[51].ToString()));
                ws.GetRow(j + 1).CreateCell(52).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[52].ToString()));
                ws.GetRow(j + 1).CreateCell(53).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[53].ToString()));
                ws.GetRow(j + 1).CreateCell(54).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[54].ToString()));
                ws.GetRow(j + 1).CreateCell(55).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[55].ToString()));
                ws.GetRow(j + 1).CreateCell(56).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[56].ToString()));
                ws.GetRow(j + 1).CreateCell(57).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[57].ToString()));
                ws.GetRow(j + 1).CreateCell(58).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[58].ToString()));
                ws.GetRow(j + 1).CreateCell(59).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[59].ToString()));
                ws.GetRow(j + 1).CreateCell(60).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[60].ToString()));
                ws.GetRow(j + 1).CreateCell(61).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[61].ToString());
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
            filename.AppendFormat(@"c:\temp\生產日報{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
                comboBox2.Text = drMOCPRODUCTDAILYREPORT.Cells["成品/半成品"].Value.ToString();
                comboBox1.Text= drMOCPRODUCTDAILYREPORT.Cells["製造組"].Value.ToString();
                textBox3.Text = drMOCPRODUCTDAILYREPORT.Cells["品號"].Value.ToString(); 
                textBox4.Text = drMOCPRODUCTDAILYREPORT.Cells["品名"].Value.ToString();
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
                textBox16.Text = drMOCPRODUCTDAILYREPORT.Cells["今日產生可回收餅麩"].Value.ToString();
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
                textBox30.Text = drMOCPRODUCTDAILYREPORT.Cells["篩選餅乾區不良烤焙(kg)"].Value.ToString();
                textBox31.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝不良餅乾(kg)"].Value.ToString();
                textBox32.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝(內袋(卷) 罐)"].Value.ToString();
                textBox33.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝投入(袋(卷),罐)"].Value.ToString();
                textBox34.Text = drMOCPRODUCTDAILYREPORT.Cells["一箱裸餅重"].Value.ToString();
                textBox35.Text = drMOCPRODUCTDAILYREPORT.Cells["一箱餅含袋重"].Value.ToString();
                textBox36.Text = drMOCPRODUCTDAILYREPORT.Cells["半成品入庫數(kg) (含袋重)"].Value.ToString();
                textBox37.Text = drMOCPRODUCTDAILYREPORT.Cells["備註"].Value.ToString();
                textBox38.Text = drMOCPRODUCTDAILYREPORT.Cells["製造工時(分)"].Value.ToString();
                textBox39.Text = drMOCPRODUCTDAILYREPORT.Cells["包裝工時(分)"].Value.ToString();
                textBox40.Text = drMOCPRODUCTDAILYREPORT.Cells["預計投入量"].Value.ToString();
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
                textBox54.Text = drMOCPRODUCTDAILYREPORT.Cells["容量規格"].Value.ToString();
                textBox55.Text = drMOCPRODUCTDAILYREPORT.Cells["預計包罐數"].Value.ToString();
                textBox56.Text = drMOCPRODUCTDAILYREPORT.Cells["實際包罐數"].Value.ToString();
                textBox57.Text = drMOCPRODUCTDAILYREPORT.Cells["總製成率"].Value.ToString();
                textBox58.Text = drMOCPRODUCTDAILYREPORT.Cells["總包罐製成率"].Value.ToString();
                textBox61.Text = drMOCPRODUCTDAILYREPORT.Cells["預計試吃包罐數"].Value.ToString();
                textBox62.Text = drMOCPRODUCTDAILYREPORT.Cells["實際試吃包罐數"].Value.ToString();
                textBox308.Text = drMOCPRODUCTDAILYREPORT.Cells["篩選餅乾區餅乾屑(kg)"].Value.ToString();
                textBox309.Text = drMOCPRODUCTDAILYREPORT.Cells["清掃廢料(kg)"].Value.ToString();
                //numericUpDown1.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["預計投入量(kg)"].Value.ToString());

                SOURCEID = drMOCPRODUCTDAILYREPORT.Cells["ID"].Value.ToString();
                SearchMOCPRODUCTDAILYREPORTPROCESSID();
            }

        }
        public void UPDATE()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] ");
                sbSql.AppendFormat(" SET [PRODUCETYPE]='{1}',[PRODUCEDEP]='{2}',[PRODUCEDATE]='{3}',[PRODUCEMB001]='{4}',[PRODUCENAME]='{5}',[PASTRYPREIN]='{6}',[PASTRY]='{7}',[PASTRYRECYCLE]='{8}',[WATERFLOURPREIN]='{9}',[WATERFLOUR]='{10}',[WATERFLOURSIDE]='{11}',[WATERFLOURRECYCLE]='{12}',[PASTRYFLODTIME]='{13}',[PASTRYFLODNUM]='{14}',[WATERFLOURTIME]='{15}',[WATERFLOURNUM]='{16}',[RECYCLEFLOUR]='{17}',[KNIFENUM]='{18}',[WEIGHTBEFRORE]='{19}',[WEIGHTAFTER]='{20}',[ROWNUM]='{21}',[RECOOKTIME]='{22}',[NGTOTAL]='{23}',[NGCOOKTIME]='{24}',[PREOUT]='{25}',[PACKAGETIME]='{26}',[PACKAGENUM]='{27}',[STIR]='{28}',[SIDES]='{29}',[COOKIES]='{30}',[COOK]='{31}',[NGPACKAGE]='{32}',[NGPACKAGECAN]='{33}',[CAN]='{34}',[WEIGHTCAN]='{35}',[WEIGHTCANBOXED]='{36}',[HLAFWEIGHT]='{37}',[REMARK]='{38}',[MANUTIME]='{39}',[PACKTIME]='{40}',[WEIGHTBEFORECOOK]='{41}',[WEIGHTAFTERCOOK]='{42}',[ACTUALOUT]='{43}',[WEIGHTPACKAGE]='{44}',[PACKLOST]='{45}',[HLAFLOST]='{46}',[REWORKPCT]='{47}',[TOTALTIME]='{48}',[STIRPCT]='{49}',[EVARATE]='{50}',[MANULOST]='{51}',[PCT]='{52}',[PRETIME]='{53}',[STOPTIME]='{54}',[PREWEIGT]='{55}',[PRECAN]='{56}',[ACTUALCAN]='{57}',[TOTALPCT]='{58}',[CANPCT]='{59}',[TRYCAN]='{60}',[ACTUALTRYCAN]='{61}',[OUTCOOKIES]='{62}',[CLEANCOOKIES]='{63}' WHERE [ID]='{0}'  ", textID.Text.ToString(), comboBox2.Text.ToString(), comboBox1.Text.ToString(), dateTimePicker2.Value.ToString("yyyy/MM/dd"),textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString(), textBox14.Text.ToString(), textBox15.Text.ToString(), textBox16.Text.ToString(), textBox17.Text.ToString(), textBox18.Text.ToString(), textBox19.Text.ToString(), textBox20.Text.ToString(), textBox21.Text.ToString(), textBox22.Text.ToString(), textBox23.Text.ToString(), textBox24.Text.ToString(), textBox25.Text.ToString(), textBox26.Text.ToString(), textBox27.Text.ToString(), textBox28.Text.ToString(), textBox29.Text.ToString(), textBox30.Text.ToString(), textBox31.Text.ToString(), textBox32.Text.ToString(), textBox33.Text.ToString(), textBox34.Text.ToString(), textBox35.Text.ToString(), textBox36.Text.ToString(), textBox37.Text.ToString(), textBox38.Text.ToString(), textBox39.Text.ToString(), textBox40.Text.ToString(), textBox41.Text.ToString(), textBox42.Text.ToString(), textBox43.Text.ToString(), textBox44.Text.ToString(), textBox45.Text.ToString(), textBox46.Text.ToString(), textBox47.Text.ToString(), textBox48.Text.ToString(), textBox49.Text.ToString(), textBox50.Text.ToString(), textBox51.Text.ToString(), textBox52.Text.ToString(), textBox53.Text.ToString(), textBox54.Text.ToString(), textBox55.Text.ToString(), textBox56.Text.ToString(), textBox57.Text.ToString(), textBox58.Text.ToString(), textBox61.Text.ToString(), textBox62.Text.ToString(), textBox308.Text.ToString(), textBox309.Text.ToString());
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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] ");
                sbSql.Append(" ( [ID],[PRODUCETYPE],[PRODUCEDEP],[PRODUCEDATE],[PRODUCEMB001],[PRODUCENAME],[PASTRYPREIN],[PASTRY],[PASTRYRECYCLE],[WATERFLOURPREIN],[WATERFLOUR],[WATERFLOURSIDE],[WATERFLOURRECYCLE],[PASTRYFLODTIME],[PASTRYFLODNUM],[WATERFLOURTIME],[WATERFLOURNUM],[RECYCLEFLOUR],[KNIFENUM],[WEIGHTBEFRORE],[WEIGHTAFTER],[ROWNUM],[RECOOKTIME],[NGTOTAL],[NGCOOKTIME],[PREOUT],[PACKAGETIME],[PACKAGENUM],[STIR],[SIDES],[COOKIES],[COOK],[NGPACKAGE],[NGPACKAGECAN],[CAN],[WEIGHTCAN],[WEIGHTCANBOXED],[HLAFWEIGHT],[REMARK],[MANUTIME],[PACKTIME],[WEIGHTBEFORECOOK],[WEIGHTAFTERCOOK],[ACTUALOUT],[WEIGHTPACKAGE],[PACKLOST],[HLAFLOST],[REWORKPCT],[TOTALTIME],[STIRPCT],[EVARATE],[MANULOST],[PCT],[PRETIME],[STOPTIME],[PREWEIGT],[PRECAN],[ACTUALCAN],[TOTALPCT] ,[CANPCT],[TRYCAN],[ACTUALTRYCAN],[OUTCOOKIES],[CLEANCOOKIES])  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}','{30}','{31}','{32}','{33}','{34}','{35}','{36}','{37}','{38}','{39}','{40}','{41}','{42}','{43}','{44}','{45}','{46}','{47}','{48}','{49}','{50}','{51}','{52}','{53}','{54}','{55}','{56}','{57}','{58}','{59}','{60}','{61}','{62}','{63}') ", Guid.NewGuid(), comboBox2.Text.ToString(), comboBox1.Text.ToString(), dateTimePicker2.Value.ToString("yyyy/MM/dd"), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString(), textBox14.Text.ToString(), textBox15.Text.ToString(), textBox16.Text.ToString(), textBox17.Text.ToString(), textBox18.Text.ToString(), textBox19.Text.ToString(), textBox20.Text.ToString(), textBox21.Text.ToString(), textBox22.Text.ToString(), textBox23.Text.ToString(), textBox24.Text.ToString(), textBox25.Text.ToString(), textBox26.Text.ToString(), textBox27.Text.ToString(), textBox28.Text.ToString(), textBox29.Text.ToString(), textBox30.Text.ToString(), textBox31.Text.ToString(), textBox32.Text.ToString(), textBox33.Text.ToString(), textBox34.Text.ToString(), textBox35.Text.ToString(), textBox36.Text.ToString(), textBox37.Text.ToString(), textBox38.Text.ToString(), textBox39.Text.ToString(), textBox40.Text.ToString(), textBox41.Text.ToString(), textBox42.Text.ToString(), textBox43.Text.ToString(), textBox44.Text.ToString(), textBox45.Text.ToString(), textBox46.Text.ToString(), textBox47.Text.ToString(), textBox48.Text.ToString(), textBox49.Text.ToString(), textBox50.Text.ToString(), textBox51.Text.ToString(), textBox52.Text.ToString(), textBox53.Text.ToString(), textBox54.Text.ToString(), textBox55.Text.ToString(), textBox56.Text.ToString(), textBox57.Text.ToString(), textBox58.Text.ToString(), textBox61.Text.ToString(), textBox62.Text.ToString(), textBox308.Text.ToString(), textBox309.Text.ToString());

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
            textBox54.Text = null;
            textBox55.Text = null;
            textBox56.Text = null;
            textBox57.Text = null;
            textBox58.Text = null;
            textBox59.Text = null;
            textBox61.Text = null;
            textBox62.Text = null;
            textBox308.Text = "0";
            textBox309.Text = "0";

            textID.ReadOnly = true;
            dateTimePicker2.Value = dateTimePicker1.Value;
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
            textBox17.ReadOnly = false;
            textBox18.ReadOnly = false;
            textBox19.ReadOnly = false;
            textBox20.ReadOnly = false;
            textBox21.ReadOnly = false;
            textBox22.ReadOnly = false;
            textBox23.ReadOnly = false;
            textBox24.ReadOnly = false;
            textBox25.ReadOnly = false;
            textBox26.ReadOnly = false;
            textBox27.ReadOnly = false;
            textBox28.ReadOnly = false;
            textBox29.ReadOnly = false;
            textBox30.ReadOnly = false;
            textBox31.ReadOnly = false;
            textBox32.ReadOnly = false;
            textBox33.ReadOnly = false;
            textBox34.ReadOnly = false;
            textBox35.ReadOnly = false;
            textBox36.ReadOnly = false;
            textBox37.ReadOnly = false;
            textBox38.ReadOnly = false;
            textBox39.ReadOnly = false;
            textBox40.ReadOnly = false;
            textBox41.ReadOnly = false;
            textBox42.ReadOnly = false;
            textBox43.ReadOnly = false;
            textBox44.ReadOnly = false;
            textBox45.ReadOnly = false;
            textBox46.ReadOnly = false;
            textBox47.ReadOnly = false;
            textBox48.ReadOnly = false;
            textBox49.ReadOnly = false;
            textBox50.ReadOnly = false;
            textBox51.ReadOnly = false;
            textBox52.ReadOnly = false;
            textBox53.ReadOnly = false;
            textBox54.ReadOnly = false;
            textBox55.ReadOnly = false;
            textBox56.ReadOnly = false;
            textBox57.ReadOnly = false;
            textBox58.ReadOnly = false;
            textBox59.ReadOnly = false;
            textBox61.ReadOnly = false;
            textBox62.ReadOnly = false;
            textBox308.ReadOnly = false;
            textBox309.ReadOnly = false;
            //textID.ReadOnly = false;

            dateTimePicker2.Enabled = true;
            dataGridView2.DataSource = null;
        }

        public void SetUPDATE()
        {
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
            textBox17.ReadOnly = false;
            textBox18.ReadOnly = false;
            textBox19.ReadOnly = false;
            textBox20.ReadOnly = false;
            textBox21.ReadOnly = false;
            textBox22.ReadOnly = false;
            textBox23.ReadOnly = false;
            textBox24.ReadOnly = false;
            textBox25.ReadOnly = false;
            textBox26.ReadOnly = false;
            textBox27.ReadOnly = false;
            textBox28.ReadOnly = false;
            textBox29.ReadOnly = false;
            textBox30.ReadOnly = false;
            textBox31.ReadOnly = false;
            textBox32.ReadOnly = false;
            textBox33.ReadOnly = false;
            textBox34.ReadOnly = false;
            textBox35.ReadOnly = false;
            textBox36.ReadOnly = false;
            textBox37.ReadOnly = false;
            textBox38.ReadOnly = false;
            textBox39.ReadOnly = false;
            textBox40.ReadOnly = false;
            textBox41.ReadOnly = false;
            textBox42.ReadOnly = false;
            textBox43.ReadOnly = false;
            textBox44.ReadOnly = false;
            textBox45.ReadOnly = false;
            textBox46.ReadOnly = false;
            textBox47.ReadOnly = false;
            textBox48.ReadOnly = false;
            textBox49.ReadOnly = false;
            textBox50.ReadOnly = false;
            textBox51.ReadOnly = false;
            textBox52.ReadOnly = false;
            textBox53.ReadOnly = false;
            textBox54.ReadOnly = false;
            textBox55.ReadOnly = false;
            textBox56.ReadOnly = false;
            textBox57.ReadOnly = false;
            textBox58.ReadOnly = false;
            textBox59.ReadOnly = false;
            textBox61.ReadOnly = false;
            textBox62.ReadOnly = false;
            textBox308.ReadOnly = false;
            textBox309.ReadOnly = false;
            //textID.ReadOnly = false;

            dateTimePicker2.Enabled = true;
        }
        public void SetFINISH()
        {
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
            textBox17.ReadOnly = true;
            textBox18.ReadOnly = true;
            textBox19.ReadOnly = true;
            textBox20.ReadOnly = true;
            textBox21.ReadOnly = true;
            textBox22.ReadOnly = true;
            textBox23.ReadOnly = true;
            textBox24.ReadOnly = true;
            textBox25.ReadOnly = true;
            textBox26.ReadOnly = true;
            textBox27.ReadOnly = true;
            textBox28.ReadOnly = true;
            textBox29.ReadOnly = true;
            textBox30.ReadOnly = true;
            textBox31.ReadOnly = true;
            textBox32.ReadOnly = true;
            textBox33.ReadOnly = true;
            textBox34.ReadOnly = true;
            textBox35.ReadOnly = true;
            textBox36.ReadOnly = true;
            textBox37.ReadOnly = true;
            textBox38.ReadOnly = true;
            textBox39.ReadOnly = true;
            textBox40.ReadOnly = true;
            textBox41.ReadOnly = true;
            textBox42.ReadOnly = true;
            textBox43.ReadOnly = true;
            textBox44.ReadOnly = true;
            textBox45.ReadOnly = true;
            textBox46.ReadOnly = true;
            textBox47.ReadOnly = true;
            textBox48.ReadOnly = true;
            textBox49.ReadOnly = true;
            textBox50.ReadOnly = true;
            textBox51.ReadOnly = true;
            textBox52.ReadOnly = true;
            textBox53.ReadOnly = true;
            textBox54.ReadOnly = true;
            textBox55.ReadOnly = true;
            textBox56.ReadOnly = true;
            textBox57.ReadOnly = true;
            textBox58.ReadOnly = true;
            textBox59.ReadOnly = true;
            textBox61.ReadOnly = true;
            textBox62.ReadOnly = true;
            textBox308.ReadOnly = false;
            textBox309.ReadOnly = true;
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

        public void CalMANUTIME()
        {
            if(!string.IsNullOrEmpty(textBox12.Text.ToString())&& !string.IsNullOrEmpty(textBox13.Text.ToString())&& !string.IsNullOrEmpty(textBox14.Text.ToString())&& !string.IsNullOrEmpty(textBox15.Text.ToString())&& !string.IsNullOrEmpty(textBox21.Text.ToString())&& !string.IsNullOrEmpty(textBox52.Text.ToString()))
            {
                textBox38.Text = Math.Round((Convert.ToDecimal(textBox12.Text.ToString()) * Convert.ToDecimal(textBox13.Text.ToString()) + Convert.ToDecimal(textBox14.Text.ToString()) * Convert.ToDecimal(textBox15.Text.ToString()) + Convert.ToDecimal(textBox21.Text.ToString()) + Convert.ToDecimal(textBox52.Text.ToString())),3).ToString();
            }
          
        }
        public void CalPACKTIME()
        {
            if (!string.IsNullOrEmpty(textBox25.Text.ToString()) && !string.IsNullOrEmpty(textBox26.Text.ToString()))
            {
                textBox39.Text = Math.Round((Convert.ToDecimal(textBox25.Text.ToString()) * Convert.ToDecimal(textBox26.Text.ToString())),3).ToString();
            }
            
        }
        public void CalWEIGHTBEFORECOOK()
        {
            if (!string.IsNullOrEmpty(textBox5.Text.ToString()) && !string.IsNullOrEmpty(textBox8.Text.ToString()) )
            {
                textBox40.Text = Math.Round((Convert.ToDecimal(textBox5.Text.ToString())+ (Convert.ToDecimal(textBox8.Text.ToString()))), 3).ToString();
            }
           
        }
        public void CalWEIGHTAFTERCOOK()
        {

            if (!string.IsNullOrEmpty(textBox19.Text.ToString()) && !string.IsNullOrEmpty(textBox20.Text.ToString()) && !string.IsNullOrEmpty(textBox17.Text.ToString()))
            {
                textBox41.Text = Math.Round((Convert.ToDecimal(textBox19.Text.ToString()) * (Convert.ToDecimal(textBox20.Text.ToString())) * (Convert.ToDecimal(textBox17.Text.ToString()))), 3).ToString();
            }

        }
        public void CalACTUALOUT()
        {
            if(!string.IsNullOrEmpty(textBox36.Text.ToString()) && !string.IsNullOrEmpty(textBox43.Text.ToString()))
            {
                textBox42.Text = Math.Round((Convert.ToDecimal(textBox36.Text.ToString())- Convert.ToDecimal(textBox43.Text.ToString())),3).ToString();
            }
            
        }
        public void CalWEIGHTPACKAGEE()
        {
            if (!string.IsNullOrEmpty(textBox36.Text.ToString()) && !string.IsNullOrEmpty(textBox34.Text.ToString()) && !string.IsNullOrEmpty(textBox35.Text.ToString()))
            {
                if (Convert.ToDecimal(textBox36.Text.ToString()) > 0 && Convert.ToDecimal(textBox34.Text.ToString()) > 0 && Convert.ToDecimal(textBox35.Text.ToString()) > 0)
                {
                    textBox43.Text = Math.Round((Convert.ToDecimal(textBox36.Text.ToString()) - (Convert.ToDecimal(textBox36.Text.ToString()) * Convert.ToDecimal(textBox34.Text.ToString()) / Convert.ToDecimal(textBox35.Text.ToString()))), 3).ToString();
                }
                    
            }
                
        }
        public void CalPACKLOST()
        {
            if(!string.IsNullOrEmpty(textBox32.Text.ToString()) && !string.IsNullOrEmpty(textBox33.Text.ToString()))
            {
                if (Convert.ToDecimal(textBox32.Text.ToString()) > 0 && Convert.ToDecimal(textBox33.Text.ToString()) > 0)
                {
                    textBox44.Text = Math.Round((Convert.ToDecimal(textBox32.Text.ToString()) / Convert.ToDecimal(textBox33.Text.ToString()) * 100), 3).ToString();
                }                    
            }
            
        }
        public void CalHLAFLOST()
        {
            if (!string.IsNullOrEmpty(textBox42.Text.ToString()) && !string.IsNullOrEmpty(textBox31.Text.ToString()) && !string.IsNullOrEmpty(textBox24.Text.ToString()))
            {
                if (Convert.ToDecimal(textBox42.Text.ToString()) > 0 && Convert.ToDecimal(textBox31.Text.ToString()) > 0 && Convert.ToDecimal(textBox24.Text.ToString()) > 0)
                {
                    textBox45.Text = Math.Round(((Convert.ToDecimal(textBox42.Text.ToString()) + Convert.ToDecimal(textBox31.Text.ToString())) / Convert.ToDecimal(textBox24.Text.ToString()) * 100), 3).ToString();
                }
                    
            }
                
        }
        public void CalREWORKPCT()
        {
            if (!string.IsNullOrEmpty(textBox21.Text.ToString()) && !string.IsNullOrEmpty(textBox47.Text.ToString()))
            {
                if (Convert.ToDecimal(textBox21.Text.ToString()) > 0 && Convert.ToDecimal(textBox47.Text.ToString()) > 0)
                {
                    textBox46.Text = Math.Round((Convert.ToDecimal(textBox21.Text.ToString()) / Convert.ToDecimal(textBox47.Text.ToString()) * 100), 3).ToString();
                }
                    
            }
            
        }
        public void CalTOTALTIME()
        {
            if (!string.IsNullOrEmpty(textBox38.Text.ToString()) && !string.IsNullOrEmpty(textBox39.Text.ToString()))
            {
                textBox47.Text = Math.Round((Convert.ToDecimal(textBox38.Text.ToString())+ Convert.ToDecimal(textBox39.Text.ToString())),3).ToString();
            }
            
        }
        public void CalSTIRPCT()
        {
            if (!string.IsNullOrEmpty(textBox6.Text.ToString()) && !string.IsNullOrEmpty(textBox9.Text.ToString()) && !string.IsNullOrEmpty(textBox27.Text.ToString()) && !string.IsNullOrEmpty(textBox28.Text.ToString()))
            {
                //if (Convert.ToDecimal(textBox6.Text.ToString()) > 0 && Convert.ToDecimal(textBox9.Text.ToString()) > 0 && Convert.ToDecimal(textBox27.Text.ToString()) > 0 && Convert.ToDecimal(textBox28.Text.ToString()) > 0)
                if (Convert.ToDecimal(textBox6.Text.ToString()) > 0 && Convert.ToDecimal(textBox9.Text.ToString()) > 0 )                    
                {
                    textBox48.Text = Math.Round(((Convert.ToDecimal(textBox6.Text.ToString()) + Convert.ToDecimal(textBox9.Text.ToString()) - Convert.ToDecimal(textBox27.Text.ToString()) - Convert.ToDecimal(textBox28.Text.ToString())) / (Convert.ToDecimal(textBox6.Text.ToString()) + Convert.ToDecimal(textBox9.Text.ToString())) * 100), 3).ToString();
                }
                    
            }
                
        }
        public void CalEVARATE()
        {
            if (!string.IsNullOrEmpty(textBox18.Text.ToString()) && !string.IsNullOrEmpty(textBox19.Text.ToString()))
            {
                if(Convert.ToDecimal(textBox18.Text.ToString())>0&& Convert.ToDecimal(textBox19.Text.ToString()) > 0)
                {
                    textBox49.Text = Math.Round(((Convert.ToDecimal(textBox18.Text.ToString()) - Convert.ToDecimal(textBox19.Text.ToString())) / Convert.ToDecimal(textBox18.Text.ToString()) * 100), 3).ToString();
                }
                
            }
           
        }
        public void CalMANULOST()
        {
            if (!string.IsNullOrEmpty(textBox42.Text.ToString()) && !string.IsNullOrEmpty(textBox30.Text.ToString()) && !string.IsNullOrEmpty(textBox31.Text.ToString()))
            {
                if (Convert.ToDecimal(textBox42.Text.ToString()) > 0 && Convert.ToDecimal(textBox30.Text.ToString()) >= 0 && Convert.ToDecimal(textBox31.Text.ToString()) >= 0)
                {
                    textBox50.Text = Math.Round(((((Convert.ToDecimal(textBox42.Text.ToString()) + Convert.ToDecimal(textBox30.Text.ToString()) + Convert.ToDecimal(textBox31.Text.ToString())) - Convert.ToDecimal(textBox42.Text.ToString())) / Convert.ToDecimal(textBox42.Text.ToString())) * 100), 3).ToString();
                }
                    
            }
                 
        }
        public void CalPCT()
        {
            if (!string.IsNullOrEmpty(textBox42.Text.ToString()) && !string.IsNullOrEmpty(textBox30.Text.ToString()) && !string.IsNullOrEmpty(textBox31.Text.ToString()))
            {
                if (Convert.ToDecimal(textBox42.Text.ToString()) > 0 && Convert.ToDecimal(textBox30.Text.ToString()) >= 0 && Convert.ToDecimal(textBox31.Text.ToString()) >= 0)
                {
                    textBox51.Text = Math.Round(((1 - ((Convert.ToDecimal(textBox30.Text.ToString()) + Convert.ToDecimal(textBox31.Text.ToString())) / Convert.ToDecimal(textBox42.Text.ToString()))) * 100), 3).ToString();
                }
                    
            }
            
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            CalMANUTIME();
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            CalMANUTIME();
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            CalMANUTIME();
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            CalMANUTIME();
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            CalMANUTIME();
            CalREWORKPCT();
        }

        private void textBox52_TextChanged(object sender, EventArgs e)
        {
            CalMANUTIME();
        }
        private void textBox25_TextChanged(object sender, EventArgs e)
        {
            CalPACKTIME();
        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {
            CalPACKTIME();
        }
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            CALTOTALIN();
            CalSTIRPCT();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            CALTOTALIN();
            CalSTIRPCT();
        }

        private void textBox27_TextChanged(object sender, EventArgs e)
        {
            
            CalSTIRPCT();
        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {            
            CalSTIRPCT();
        }
        private void textBox17_TextChanged(object sender, EventArgs e)
        {
            CalWEIGHTAFTERCOOK();
            //CalWEIGHTBEFORECOOK();
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            CalWEIGHTAFTERCOOK();
            CalEVARATE();

        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            CalWEIGHTAFTERCOOK();
            //CalWEIGHTBEFORECOOK();
        }
        private void textBox36_TextChanged(object sender, EventArgs e)
        {
            CalACTUALOUT();
            CalWEIGHTPACKAGEE();
        }

        private void textBox43_TextChanged(object sender, EventArgs e)
        {
            CalACTUALOUT();
        }
        private void textBox34_TextChanged(object sender, EventArgs e)
        {
            CalWEIGHTPACKAGEE();
        }

        private void textBox35_TextChanged(object sender, EventArgs e)
        {
            CalWEIGHTPACKAGEE();
        }
        private void textBox32_TextChanged(object sender, EventArgs e)
        {
            CalPACKLOST();
        }

        private void textBox33_TextChanged(object sender, EventArgs e)
        {
            CalPACKLOST();
        }
        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            CalHLAFLOST();
        }

        private void textBox31_TextChanged(object sender, EventArgs e)
        {
            CalHLAFLOST();
            CalMANULOST();
            CalPCT();
        }
        private void textBox38_TextChanged(object sender, EventArgs e)
        {
            CalTOTALTIME();
        }

        private void textBox39_TextChanged(object sender, EventArgs e)
        {
            CalTOTALTIME();
        }
        private void textBox47_TextChanged(object sender, EventArgs e)
        {
            CalREWORKPCT();
        }
         private void textBox40_TextChanged(object sender, EventArgs e)
        {
           
            CalWEIGHTBEFORECOOK();
        }

        private void textBox41_TextChanged(object sender, EventArgs e)
        {
            //CalEVARATE();
        }
        private void textBox42_TextChanged(object sender, EventArgs e)
        {
            CalMANULOST();
            CalPCT();
            CalHLAFLOST();
        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {
            CalMANULOST();
            CalPCT();
        }
        

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox4.Text = FINDMB002(textBox3.Text.ToString());
        }

        public string FINDMB002(string MB001)
        {
            DataSet ds = new DataSet();
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
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
                return ds.Tables["TEMPds1"].Rows[0]["MB002"].ToString();
            }
            else
            {
                return "";
            }


        }
        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text.ToString().Equals("成品"))
            {
                textBox34.Text = "1";
                textBox35.Text = "1";
                textBox34.ReadOnly = true;
                textBox35.ReadOnly = true;
            }
            else if (comboBox2.Text.ToString().Equals("半成品"))
            {
                textBox34.Text = null;
                textBox35.Text = null;
                textBox34.ReadOnly = false;
                textBox35.ReadOnly = false;
            }
        }
    
        public void CALPRECAN()
        {
            if (!string.IsNullOrEmpty(textBox36.Text.ToString()) && !string.IsNullOrEmpty(textBox54.Text.ToString()))
            {
                if (Convert.ToDecimal(textBox36.Text.ToString()) > 0 && Convert.ToDecimal(textBox54.Text.ToString()) > 0)
                {
                    textBox55.Text = Math.Round(((Convert.ToDecimal(textBox36.Text.ToString())) / (Convert.ToDecimal(textBox54.Text.ToString()))), 0).ToString();
                }

            }
        }
        private void textBox54_TextChanged(object sender, EventArgs e)
        {
            CALPRECAN();
        }
        public void CALTOTALPCT()
        {
            if (!string.IsNullOrEmpty(textBox48.Text.ToString()) && !string.IsNullOrEmpty(textBox51.Text.ToString()))
            {
                if (Convert.ToDecimal(textBox48.Text.ToString()) > 0 && Convert.ToDecimal(textBox51.Text.ToString()) > 0)
                {
                    textBox57.Text = Math.Round(((Convert.ToDecimal(textBox48.Text.ToString())) * (Convert.ToDecimal(textBox51.Text.ToString())))/100, 2).ToString();
                }

            }
        }
        private void textBox48_TextChanged(object sender, EventArgs e)
        {
            CALTOTALPCT();
        }

        private void textBox51_TextChanged(object sender, EventArgs e)
        {
            CALTOTALPCT();
        }
        public void CALCANPCT()
        {
            if (!string.IsNullOrEmpty(textBox55.Text.ToString()) && !string.IsNullOrEmpty(textBox56.Text.ToString()))
            {
                if (Convert.ToDecimal(textBox55.Text.ToString()) > 0 && Convert.ToDecimal(textBox56.Text.ToString()) > 0)
                {
                    textBox58.Text = Math.Round(((Convert.ToDecimal(textBox56.Text.ToString())) /(Convert.ToDecimal(textBox55.Text.ToString()))) * 100, 2).ToString();
                }

            }
        }
        private void textBox55_TextChanged(object sender, EventArgs e)
        {
            CALCANPCT();
        }

        private void textBox56_TextChanged(object sender, EventArgs e)
        {
            CALCANPCT();
        }

        public void CHECKNULL()
        {
            if(string.IsNullOrEmpty(textBox5.Text))
            {
                textBox5.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox6.Text))
            {
                textBox6.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox7.Text))
            {
                textBox7.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox8.Text))
            {
                textBox8.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox9.Text))
            {
                textBox9.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox10.Text))
            {
                textBox10.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox11.Text))
            {
                textBox11.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox12.Text))
            {
                textBox12.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox13.Text))
            {
                textBox13.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox14.Text))
            {
                textBox14.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox15.Text))
            {
                textBox15.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox16.Text))
            {
                textBox16.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox17.Text))
            {
                textBox17.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox18.Text))
            {
                textBox18.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox19.Text))
            {
                textBox19.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox20.Text))
            {
                textBox20.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox21.Text))
            {
                textBox21.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox22.Text))
            {
                textBox22.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox23.Text))
            {
                textBox23.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox24.Text))
            {
                textBox24.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox25.Text))
            {
                textBox25.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox26.Text))
            {
                textBox26.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox27.Text))
            {
                textBox27.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox28.Text))
            {
                textBox28.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox29.Text))
            {
                textBox29.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox30.Text))
            {
                textBox30.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox31.Text))
            {
                textBox31.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox32.Text))
            {
                textBox32.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox33.Text))
            {
                textBox33.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox34.Text))
            {
                textBox34.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox35.Text))
            {
                textBox35.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox36.Text))
            {
                textBox36.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox37.Text))
            {
                textBox37.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox38.Text))
            {
                textBox38.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox39.Text))
            {
                textBox39.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox40.Text))
            {
                textBox40.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox41.Text))
            {
                textBox41.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox42.Text))
            {
                textBox42.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox43.Text))
            {
                textBox43.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox44.Text))
            {
                textBox44.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox45.Text))
            {
                textBox45.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox46.Text))
            {
                textBox46.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox47.Text))
            {
                textBox47.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox48.Text))
            {
                textBox48.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox49.Text))
            {
                textBox49.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox50.Text))
            {
                textBox50.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox51.Text))
            {
                textBox51.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox52.Text))
            {
                textBox52.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox53.Text))
            {
                textBox53.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox54.Text))
            {
                textBox54.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox55.Text))
            {
                textBox55.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox56.Text))
            {
                textBox56.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox57.Text))
            {
                textBox57.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox58.Text))
            {
                textBox58.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox59.Text))
            {
                textBox59.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox61.Text))
            {
                textBox61.Text = "0";
            }
            if (string.IsNullOrEmpty(textBox62.Text))
            {
                textBox62.Text = "0";
            }

            
        }
        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            //CalWEIGHTBEFORECOOK();
            CalEVARATE();
        }

        public void UPDATEWEIGHTAFTER()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("   UPDATE [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT]");
                sbSql.AppendFormat("   SET   [WEIGHTAFTERCOOK]=[KNIFENUM]*[ROWNUM]*[WEIGHTAFTER]");
                sbSql.AppendFormat("   WHERE  [WEIGHTAFTERCOOK]<>[KNIFENUM]*[ROWNUM]*[WEIGHTAFTER]");
                sbSql.AppendFormat("   ");
                sbSql.AppendFormat("   UPDATE  [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT]");
                sbSql.AppendFormat("   SET [WEIGHTBEFORECOOK]=[PASTRYPREIN] +[WATERFLOURPREIN] ");
                sbSql.AppendFormat("   WHERE [WEIGHTBEFORECOOK]<>[PASTRYPREIN] +[WATERFLOURPREIN]");
                sbSql.AppendFormat("   ");
                sbSql.AppendFormat("   UPDATE  [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT]");
                sbSql.AppendFormat("   SET [TOTALIN]=[PASTRY]+[PASTRYRECYCLE]+[WATERFLOUR] +[WATERFLOURSIDE]+[WATERFLOURRECYCLE]");
                sbSql.AppendFormat("   WHERE [TOTALIN]<>[PASTRY]+[PASTRYRECYCLE]+[WATERFLOUR] +[WATERFLOURSIDE]+[WATERFLOURRECYCLE]");
                sbSql.AppendFormat("   ");
                sbSql.AppendFormat("   UPDATE [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT]");
                sbSql.AppendFormat("   SET [EVARATE]=(([WEIGHTBEFRORE]-[WEIGHTAFTER])/[WEIGHTBEFRORE])*100");
                sbSql.AppendFormat("   WHERE [EVARATE]<>(([WEIGHTBEFRORE]-[WEIGHTAFTER])/[WEIGHTBEFRORE])*100");
                sbSql.AppendFormat("   ");


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

        private void textBox59_TextChanged(object sender, EventArgs e)
        {
            CALTOTALIN();
        }

        public void CALTOTALIN()
        {
            if (!string.IsNullOrEmpty(textBox6.Text.ToString()) && !string.IsNullOrEmpty(textBox7.Text.ToString()) && !string.IsNullOrEmpty(textBox9.Text.ToString()) && !string.IsNullOrEmpty(textBox10.Text.ToString()) && !string.IsNullOrEmpty(textBox11.Text.ToString()))
            {
                textBox59.Text = Math.Round((Convert.ToDecimal(textBox6.Text.ToString()))+ (Convert.ToDecimal(textBox7.Text.ToString()))+ (Convert.ToDecimal(textBox9.Text.ToString()))+ (Convert.ToDecimal(textBox10.Text.ToString()))+ (Convert.ToDecimal(textBox11.Text.ToString())), 3).ToString();
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            CALTOTALIN();
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            CALTOTALIN();
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            CALTOTALIN();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            CalWEIGHTBEFORECOOK();
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            CalWEIGHTBEFORECOOK();
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
            CHECKNULL();
            if (!string.IsNullOrEmpty(textID.Text.ToString()))
            {
                UPDATE();
            }
            else
            {
                ADD();
            }

            UPDATEWEIGHTAFTER();
            SetFINISH();
            Search();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExcelExport();           
            //PRINTDOC();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            frmMOCPRODUCTDAILYREPORTPROCESSID objMOCPRODUCTDAILYREPORTPROCESSID = new frmMOCPRODUCTDAILYREPORTPROCESSID(SOURCEID);
            objMOCPRODUCTDAILYREPORTPROCESSID.ShowDialog();
            Search();

        }
        private void button7_Click(object sender, EventArgs e)
        {
            SetFINISH();
        }








        #endregion


    }
}
