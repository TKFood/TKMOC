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

                sbSql.Append(@" SELECT [PRODUCEDATE] AS '日期',[PRODUCEID]  AS '製令單號',[PRODUCENAME]  AS '品號/品名' ");
                sbSql.Append(@" ,[PREIN]  AS '預計投入量(kg)',[PROCESSTIME]  AS '製造時間(分)',[PRECESSPEOPEL]  AS '製造人數'");
                sbSql.Append(@" ,[MATERIAL]  AS '油酥原料(1)',[RECYCLEMATERIAL]  AS '油酥可回收餅麩(2)',[MATERIALTIME]  AS '水面製造時間(分)(3)'");
                sbSql.Append(@" ,[MATERIALPEOPLE]  AS '人數(4)',[WMATERIAL]  AS '水面原料(5)',[WRECYCLESIDE]  AS '水面可回收邊料(6)'");
                sbSql.Append(@" ,[WRECYCLECOOKIES]  AS '水面可回收餅麩(7)',[PACKAGETIME]  AS '包裝時間(內包裝區+罐裝)(分)(8)'");
                sbSql.Append(@" ,[PACKAGEPEOPLE]  AS '包裝人數(9)',[TRECYCLE]  AS '可回收餅麩(10)=(2)+(7)(kg)',[NGTATOL]  AS '未熟總量(11)(kg)'");
                sbSql.Append(@" ,[NGTIME]  AS '未熟烤焙時間(12)(分)',[WEIGHTBEFORE]  AS '烤前單片重量(g)',[WEIGHTAFTER]  AS '烤後單片重量(g)'");
                sbSql.Append(@" ,[COOKIENUM]  AS '每排數量',[BLADENUM]  AS '刀數',[INBEFORE]  AS '烤前實際總投入(kg)',[INAFTER]  AS '烤後實際總投入(kg)'");
                sbSql.Append(@" ,[PREOUT]  AS '預計產出(kg)',[NGSTIR]  AS '製造課不良-攪拌(kg)',[NGSIDE]  AS '製造課不良-成型邊料(kg)'");
                sbSql.Append(@" ,[NGCOOKIES]  AS '製造課不良-餅麩(kg)',[NGBAKE]  AS '製造課不良-烤焙(kg)',[NGNOGOOD]  AS '包裝課不良-包裝不良餅乾(kg)'");
                sbSql.Append(@" ,[NGNOCAN]  AS '包裝課不良-包裝(內袋(卷) 罐)',[PACKAGEWEIGHT]  AS '袋重(kg)',[PACKAGEIN]  AS '包裝投入(袋(卷),罐)'");
                sbSql.Append(@" ,[ACTUALOUT]  AS '實際產出(kg)(裸餅)',[HALFOUT]  AS '半成品入庫數(kg) (含袋重)',[REMARK]  AS '備註'");
                sbSql.Append(@" ,[MANUTIME]  AS '製造工時',[STIRPCT]  AS '攪拌成型製成率%',[EVARATE]  AS '蒸發率'");
                sbSql.Append(@" ,[LOSTRATE]  AS '製成損失率',[HALFRATE]  AS '半成品產出效率',[PACKAGETOTALTIME]  AS '包裝工時'");
                sbSql.Append(@",[WEIGHTTOTAL]  AS '袋重',[PACKAGELOSTRATE]  AS '包裝損耗率',[TOTALRATE]  AS '製成率' ");
                sbSql.Append(@" ,[CANCOOKIESWEIGHT]  AS '一箱裸餅重',[CANWEIGHT]   AS '一箱餅含袋重',[ID]");
                sbSql.Append(@" ");
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count >= 1)
            {
                ID = dataGridView1.CurrentRow.Cells["ID"].Value.ToString();
                drMOCPRODUCTDAILYREPORT = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];

                textBox3.Text = drMOCPRODUCTDAILYREPORT.Cells["製令單號"].Value.ToString(); 
                textBox4.Text = drMOCPRODUCTDAILYREPORT.Cells["品號/品名"].Value.ToString();
                textBox5.Text = drMOCPRODUCTDAILYREPORT.Cells["備註"].Value.ToString();
                textID.Text = ID;
                numericUpDown1.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["預計投入量(kg)"].Value.ToString());
                numericUpDown2.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["製造時間(分)"].Value.ToString());
                numericUpDown3.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["製造人數"].Value.ToString());
                numericUpDown4.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["油酥原料(1)"].Value.ToString());
                numericUpDown5.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["油酥可回收餅麩(2)"].Value.ToString());
                numericUpDown6.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["水面製造時間(分)(3)"].Value.ToString());
                numericUpDown7.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["人數(4)"].Value.ToString());
                numericUpDown8.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["水面原料(5)"].Value.ToString());
                numericUpDown9.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["水面可回收邊料(6)"].Value.ToString());
                numericUpDown10.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["水面可回收餅麩(7)"].Value.ToString());
                numericUpDown11.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["包裝時間(內包裝區+罐裝)(分)(8)"].Value.ToString());
                numericUpDown12.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["包裝人數(9)"].Value.ToString());
                numericUpDown13.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["可回收餅麩(10)=(2)+(7)(kg)"].Value.ToString());
                numericUpDown14.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["未熟總量(11)(kg)"].Value.ToString());
                numericUpDown15.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["未熟烤焙時間(12)(分)"].Value.ToString());
                numericUpDown16.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["烤前單片重量(g)"].Value.ToString());
                numericUpDown17.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["烤後單片重量(g)"].Value.ToString());
                numericUpDown18.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["每排數量"].Value.ToString());
                numericUpDown19.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["刀數"].Value.ToString());
                numericUpDown20.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["烤前實際總投入(kg)"].Value.ToString());
                numericUpDown21.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["烤後實際總投入(kg)"].Value.ToString());
                numericUpDown22.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["預計產出(kg)"].Value.ToString());
                numericUpDown23.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["製造課不良-攪拌(kg)"].Value.ToString());
                numericUpDown24.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["製造課不良-成型邊料(kg)"].Value.ToString());
                numericUpDown25.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["製造課不良-餅麩(kg)"].Value.ToString());
                numericUpDown26.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["製造課不良-烤焙(kg)"].Value.ToString());
                numericUpDown27.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["包裝課不良-包裝不良餅乾(kg)"].Value.ToString());
                numericUpDown28.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["包裝課不良-包裝(內袋(卷) 罐)"].Value.ToString());
                numericUpDown29.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["袋重(kg)"].Value.ToString());
                numericUpDown30.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["包裝投入(袋(卷),罐)"].Value.ToString());
                numericUpDown31.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["實際產出(kg)(裸餅)"].Value.ToString());
                numericUpDown32.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["半成品入庫數(kg) (含袋重)"].Value.ToString());
                numericUpDown33.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["製造工時"].Value.ToString());
                numericUpDown34.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["攪拌成型製成率%"].Value.ToString());
                numericUpDown35.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["蒸發率"].Value.ToString());
                numericUpDown36.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["製成損失率"].Value.ToString());
                numericUpDown37.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["半成品產出效率"].Value.ToString());
                numericUpDown38.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["包裝工時"].Value.ToString());
                numericUpDown39.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["袋重"].Value.ToString());
                numericUpDown40.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["包裝損耗率"].Value.ToString());
                numericUpDown41.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["製成率"].Value.ToString());
                numericUpDown42.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["一箱裸餅重"].Value.ToString());
                numericUpDown43.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["一箱餅含袋重"].Value.ToString());
                
                dateTimePicker2.Value = Convert.ToDateTime(drMOCPRODUCTDAILYREPORT.Cells["日期"].Value.ToString()); 
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
                sbSql.AppendFormat("  SET [PRODUCEDATE]='{1}',[PRODUCEID]='{2}',[PRODUCENAME]='{3}',[PREIN]='{4}',[PROCESSTIME]='{5}',[PRECESSPEOPEL]='{6}',[MATERIAL]='{7}',[RECYCLEMATERIAL]='{8}',[MATERIALTIME]='{9}',[MATERIALPEOPLE]='{10}',[WMATERIAL]='{11}',[WRECYCLESIDE]='{12}',[WRECYCLECOOKIES]='{13}',[PACKAGETIME]='{14}',[PACKAGEPEOPLE]='{15}',[TRECYCLE]='{16}',[NGTATOL]='{17}',[NGTIME]='{18}',[WEIGHTBEFORE]='{19}',[WEIGHTAFTER]='{20}',[COOKIENUM]='{21}',[BLADENUM]='{22}',[INBEFORE]='{23}',[INAFTER]='{24}',[PREOUT]='{25}',[NGSTIR]='{26}',[NGSIDE]='{27}',[NGCOOKIES]='{28}',[NGBAKE]='{29}',[NGNOGOOD]='{30}',[NGNOCAN]='{31}',[PACKAGEWEIGHT]='{32}',[PACKAGEIN]='{33}',[ACTUALOUT]='{34}',[HALFOUT]='{35}',[REMARK]='{36}',[MANUTIME]='{37}',[STIRPCT]='{38}',[EVARATE]='{39}',[LOSTRATE]='{40}',[HALFRATE]='{41}',[PACKAGETOTALTIME]='{42}',[WEIGHTTOTAL]='{43}',[PACKAGELOSTRATE]='{44}',[TOTALRATE]='{45}',[CANCOOKIESWEIGHT]='{46}',[CANWEIGHT] ='{47}' WHERE [ID]='{0}' ", textID.Text.ToString(), dateTimePicker2.Value.ToString("yyyy/MM/dd"), textBox3.Text.ToString(), textBox4.Text.ToString(), numericUpDown1.Value, numericUpDown2.Value, numericUpDown3.Value, numericUpDown4.Value, numericUpDown5.Value, numericUpDown6.Value, numericUpDown7.Value, numericUpDown8.Value, numericUpDown9.Value, numericUpDown10.Value, numericUpDown11.Value, numericUpDown12.Value, numericUpDown13.Value, numericUpDown14.Value, numericUpDown15.Value, numericUpDown16.Value, numericUpDown17.Value, numericUpDown18.Value, numericUpDown19.Value, numericUpDown20.Value, numericUpDown21.Value, numericUpDown22.Value, numericUpDown23.Value, numericUpDown24.Value, numericUpDown25.Value, numericUpDown26.Value, numericUpDown27.Value, numericUpDown28.Value, numericUpDown29.Value, numericUpDown30.Value, numericUpDown31.Value, numericUpDown32.Value, textBox5.Text.ToString(), numericUpDown33.Value, numericUpDown34.Value, numericUpDown35.Value, numericUpDown36.Value, numericUpDown37.Value, numericUpDown38.Value, numericUpDown39.Value, numericUpDown40.Value, numericUpDown41.Value, numericUpDown42.Value, numericUpDown43.Value);
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
                sbSql.Append(" ([ID],[PRODUCEDATE],[PRODUCEID],[PRODUCENAME],[PREIN],[PROCESSTIME],[PRECESSPEOPEL],[MATERIAL],[RECYCLEMATERIAL],[MATERIALTIME],[MATERIALPEOPLE],[WMATERIAL],[WRECYCLESIDE],[WRECYCLECOOKIES],[PACKAGETIME],[PACKAGEPEOPLE],[TRECYCLE],[NGTATOL],[NGTIME],[WEIGHTBEFORE],[WEIGHTAFTER],[COOKIENUM],[BLADENUM],[INBEFORE],[INAFTER],[PREOUT],[NGSTIR],[NGSIDE],[NGCOOKIES],[NGBAKE],[NGNOGOOD],[NGNOCAN],[PACKAGEWEIGHT],[PACKAGEIN],[ACTUALOUT],[HALFOUT],[REMARK],[MANUTIME],[STIRPCT],[EVARATE],[LOSTRATE],[HALFRATE],[PACKAGETOTALTIME],[WEIGHTTOTAL],[PACKAGELOSTRATE],[TOTALRATE],[CANCOOKIESWEIGHT],[CANWEIGHT]  )  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}','{30}','{31}','{32}','{33}','{34}','{35}','{36}','{37}','{38}','{39}','{40}','{41}','{42}','{43}','{44}','{45}','{46}','{47}') ", Guid.NewGuid(), dateTimePicker2.Value.ToString("yyyy/MM/dd"),textBox3.Text.ToString(),textBox4.Text.ToString(),numericUpDown1.Value, numericUpDown2.Value, numericUpDown3.Value, numericUpDown4.Value, numericUpDown5.Value, numericUpDown6.Value, numericUpDown7.Value, numericUpDown8.Value, numericUpDown9.Value, numericUpDown10.Value, numericUpDown11.Value, numericUpDown12.Value, numericUpDown13.Value, numericUpDown14.Value, numericUpDown15.Value, numericUpDown16.Value, numericUpDown17.Value, numericUpDown18.Value, numericUpDown19.Value, numericUpDown20.Value, numericUpDown21.Value, numericUpDown22.Value, numericUpDown23.Value, numericUpDown24.Value, numericUpDown25.Value, numericUpDown26.Value, numericUpDown27.Value, numericUpDown28.Value, numericUpDown29.Value, numericUpDown30.Value, numericUpDown31.Value, numericUpDown32.Value,textBox5.Text.ToString(), numericUpDown33.Value, numericUpDown34.Value, numericUpDown35.Value, numericUpDown36.Value, numericUpDown37.Value, numericUpDown38.Value, numericUpDown39.Value, numericUpDown40.Value, numericUpDown41.Value, numericUpDown42.Value, numericUpDown43.Value);

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
            textBox5.Text = null;
            textID.Text = null;
            numericUpDown1.Value = 0;
            numericUpDown2.Value = 0;
            numericUpDown3.Value = 0;
            numericUpDown4.Value = 0;
            numericUpDown5.Value = 0;
            numericUpDown6.Value = 0;
            numericUpDown7.Value = 0;
            numericUpDown8.Value = 0;
            numericUpDown9.Value = 0;
            numericUpDown10.Value = 0;
            numericUpDown11.Value = 0;
            numericUpDown12.Value = 0;
            numericUpDown13.Value = 0;
            numericUpDown14.Value = 0;
            numericUpDown15.Value = 0;
            numericUpDown16.Value = 0;
            numericUpDown17.Value = 0;
            numericUpDown18.Value = 0;
            numericUpDown19.Value = 0;
            numericUpDown20.Value = 0;
            numericUpDown21.Value = 0;
            numericUpDown22.Value = 0;
            numericUpDown23.Value = 0;
            numericUpDown24.Value = 0;
            numericUpDown25.Value = 0;
            numericUpDown26.Value = 0;
            numericUpDown27.Value = 0;
            numericUpDown28.Value = 0;
            numericUpDown29.Value = 0;
            numericUpDown30.Value = 0;
            numericUpDown31.Value = 0;
            numericUpDown32.Value = 0;
            numericUpDown33.Value = 0;
            numericUpDown34.Value = 0;
            numericUpDown35.Value = 0;
            numericUpDown36.Value = 0;
            numericUpDown37.Value = 0;
            numericUpDown38.Value = 0;
            numericUpDown39.Value = 0;
            numericUpDown40.Value = 0;
            numericUpDown41.Value = 0;
            numericUpDown42.Value = 0;
            numericUpDown43.Value = 0;

            dateTimePicker2.Value = DateTime.Now;
            //
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            //textID.ReadOnly = false;
            numericUpDown1.Enabled = true;
            numericUpDown2.Enabled = true;
            numericUpDown3.Enabled = true;
            numericUpDown4.Enabled = true;
            numericUpDown5.Enabled = true;
            numericUpDown6.Enabled = true;
            numericUpDown7.Enabled = true;
            numericUpDown8.Enabled = true;
            numericUpDown9.Enabled = true;
            numericUpDown10.Enabled = true;
            numericUpDown11.Enabled = true;
            numericUpDown12.Enabled = true;
            numericUpDown13.Enabled = true;
            numericUpDown14.Enabled = true;
            numericUpDown15.Enabled = true;
            numericUpDown16.Enabled = true;
            numericUpDown17.Enabled = true;
            numericUpDown18.Enabled = true;
            numericUpDown19.Enabled = true;
            numericUpDown20.Enabled = true;
            numericUpDown21.Enabled = true;
            numericUpDown22.Enabled = true;
            numericUpDown23.Enabled = true;
            numericUpDown24.Enabled = true;
            numericUpDown25.Enabled = true;
            numericUpDown26.Enabled = true;
            numericUpDown27.Enabled = true;
            numericUpDown28.Enabled = true;
            numericUpDown29.Enabled = true;
            numericUpDown30.Enabled = true;
            numericUpDown31.Enabled = true;
            numericUpDown32.Enabled = true;
            numericUpDown33.Enabled = true;
            numericUpDown34.Enabled = true;
            numericUpDown35.Enabled = true;
            numericUpDown36.Enabled = true;
            numericUpDown37.Enabled = true;
            numericUpDown38.Enabled = true;
            numericUpDown39.Enabled = true;
            numericUpDown40.Enabled = true;
            numericUpDown41.Enabled = true;
            numericUpDown42.Enabled = true;
            numericUpDown43.Enabled = true;
            dateTimePicker2.Enabled = true;
        }

        public void SetUPDATE()
        {
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            //textID.ReadOnly = false;
            numericUpDown1.Enabled = true;
            numericUpDown2.Enabled = true;
            numericUpDown3.Enabled = true;
            numericUpDown4.Enabled = true;
            numericUpDown5.Enabled = true;
            numericUpDown6.Enabled = true;
            numericUpDown7.Enabled = true;
            numericUpDown8.Enabled = true;
            numericUpDown9.Enabled = true;
            numericUpDown10.Enabled = true;
            numericUpDown11.Enabled = true;
            numericUpDown12.Enabled = true;
            numericUpDown13.Enabled = true;
            numericUpDown14.Enabled = true;
            numericUpDown15.Enabled = true;
            numericUpDown16.Enabled = true;
            numericUpDown17.Enabled = true;
            numericUpDown18.Enabled = true;
            numericUpDown19.Enabled = true;
            numericUpDown20.Enabled = true;
            numericUpDown21.Enabled = true;
            numericUpDown22.Enabled = true;
            numericUpDown23.Enabled = true;
            numericUpDown24.Enabled = true;
            numericUpDown25.Enabled = true;
            numericUpDown26.Enabled = true;
            numericUpDown27.Enabled = true;
            numericUpDown28.Enabled = true;
            numericUpDown29.Enabled = true;
            numericUpDown30.Enabled = true;
            numericUpDown31.Enabled = true;
            numericUpDown32.Enabled = true;
            numericUpDown33.Enabled = true;
            numericUpDown34.Enabled = true;
            numericUpDown35.Enabled = true;
            numericUpDown36.Enabled = true;
            numericUpDown37.Enabled = true;
            numericUpDown38.Enabled = true;
            numericUpDown39.Enabled = true;
            numericUpDown40.Enabled = true;
            numericUpDown41.Enabled = true;
            numericUpDown42.Enabled = true;
            numericUpDown43.Enabled = true;
            dateTimePicker2.Enabled = true;
        }
        public void SetFINISH()
        {
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
            textBox5.ReadOnly = true;
            textID.ReadOnly = true;
            numericUpDown1.Enabled = false;
            numericUpDown2.Enabled = false;
            numericUpDown3.Enabled = false;
            numericUpDown4.Enabled = false;
            numericUpDown5.Enabled = false;
            numericUpDown6.Enabled = false;
            numericUpDown7.Enabled = false;
            numericUpDown8.Enabled = false;
            numericUpDown9.Enabled = false;
            numericUpDown10.Enabled = false;
            numericUpDown11.Enabled = false;
            numericUpDown12.Enabled = false;
            numericUpDown13.Enabled = false;
            numericUpDown14.Enabled = false;
            numericUpDown15.Enabled = false;
            numericUpDown16.Enabled = false;
            numericUpDown17.Enabled = false;
            numericUpDown18.Enabled = false;
            numericUpDown19.Enabled = false;
            numericUpDown20.Enabled = false;
            numericUpDown21.Enabled = false;
            numericUpDown22.Enabled = false;
            numericUpDown23.Enabled = false;
            numericUpDown24.Enabled = false;
            numericUpDown25.Enabled = false;
            numericUpDown26.Enabled = false;
            numericUpDown27.Enabled = false;
            numericUpDown28.Enabled = false;
            numericUpDown29.Enabled = false;
            numericUpDown30.Enabled = false;
            numericUpDown31.Enabled = false;
            numericUpDown32.Enabled = false;
            numericUpDown33.Enabled = false;
            numericUpDown34.Enabled = false;
            numericUpDown35.Enabled = false;
            numericUpDown36.Enabled = false;
            numericUpDown37.Enabled = false;
            numericUpDown38.Enabled = false;
            numericUpDown39.Enabled = false;
            numericUpDown40.Enabled = false;
            numericUpDown41.Enabled = false;
            numericUpDown42.Enabled = false;
            numericUpDown43.Enabled = false;
            dateTimePicker2.Enabled = false;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            CalINBEFORE();
        }
        private void numericUpDown23_ValueChanged(object sender, EventArgs e)
        {
            CalINBEFORE();
            CalSTIRPCT();
        }

        private void numericUpDown24_ValueChanged(object sender, EventArgs e)
        {
            CalINBEFORE();
            CalSTIRPCT();
        }
        public void CalINBEFORE()
        {
            numericUpDown20.Value = numericUpDown1.Value - numericUpDown23.Value - numericUpDown24.Value;
        }

        public void CalINAFTERE()
        {
            numericUpDown21.Value = (numericUpDown17.Value * numericUpDown18.Value* numericUpDown19.Value)/1000;
        }
        private void numericUpDown17_ValueChanged(object sender, EventArgs e)
        {
            CalINAFTERE();
        }

        private void numericUpDown18_ValueChanged(object sender, EventArgs e)
        {
            CalINAFTERE();
        }

        private void numericUpDown19_ValueChanged(object sender, EventArgs e)
        {
            CalINAFTERE();
        }
        public void CalACTUALOUT()
        {
            numericUpDown31.Value = numericUpDown32.Value - numericUpDown29.Value;
        }
        private void numericUpDown29_ValueChanged(object sender, EventArgs e)
        {
            CalACTUALOUT();
        }

        private void numericUpDown32_ValueChanged(object sender, EventArgs e)
        {
            CalACTUALOUT();
            CalWEIGHTTOTAL();
        }

        public void CalMANUTIME()
        {
            numericUpDown33.Value = (numericUpDown2.Value* numericUpDown3.Value) + (numericUpDown6.Value * numericUpDown7.Value);
        }
        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            CalMANUTIME();
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            CalMANUTIME();
        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            CalMANUTIME();
        }

        private void numericUpDown7_ValueChanged(object sender, EventArgs e)
        {
            CalMANUTIME();
        }
        public void CalSTIRPCT()
        {
            numericUpDown34.Value = (numericUpDown4.Value + numericUpDown8.Value - numericUpDown23.Value - numericUpDown24.Value) / (numericUpDown4.Value + numericUpDown8.Value) * 100;
        }
        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            CalSTIRPCT();
        }

        private void numericUpDown8_ValueChanged(object sender, EventArgs e)
        {
            CalSTIRPCT();
        }
        public void CalEVARATE()
        {
            numericUpDown35.Value = (numericUpDown20.Value - numericUpDown21.Value) / numericUpDown20.Value * 100;
        }
        private void numericUpDown20_ValueChanged(object sender, EventArgs e)
        {
            CalEVARATE();
        }

        private void numericUpDown21_ValueChanged(object sender, EventArgs e)
        {
            CalEVARATE();
            CalHALFRATE();
        }

        public void CalLOSTRATE()
        {
            numericUpDown36.Value = (numericUpDown31.Value + numericUpDown26.Value + numericUpDown27.Value) - numericUpDown31.Value / numericUpDown31.Value * 100;
        }
        private void numericUpDown26_ValueChanged(object sender, EventArgs e)
        {
            CalLOSTRATE();
            CalHALFRATE();
            CalTOTALRATE();
        }

        private void numericUpDown27_ValueChanged(object sender, EventArgs e)
        {
            CalLOSTRATE();
            CalTOTALRATE();
        }

        private void numericUpDown31_ValueChanged(object sender, EventArgs e)
        {
            CalLOSTRATE();
            CalHALFRATE();
            CalTOTALRATE();
        }
        public void CalHALFRATE()
        {
            numericUpDown37.Value = (numericUpDown31.Value + numericUpDown26.Value) / numericUpDown21.Value * 100;
        }
        public void CalPACKAGETOTALTIME()
        {
            numericUpDown38.Value = numericUpDown11.Value * numericUpDown12.Value;
        }
        private void numericUpDown11_ValueChanged(object sender, EventArgs e)
        {
            CalPACKAGETOTALTIME();
        }

        private void numericUpDown12_ValueChanged(object sender, EventArgs e)
        {
            CalPACKAGETOTALTIME();
        }
        public void CalWEIGHTTOTAL()
        {
            numericUpDown39.Value = numericUpDown32.Value - (numericUpDown32.Value / numericUpDown42.Value * numericUpDown43.Value);
        }

        private void numericUpDown42_ValueChanged(object sender, EventArgs e)
        {
            CalWEIGHTTOTAL();
        }

        private void numericUpDown43_ValueChanged(object sender, EventArgs e)
        {
            CalWEIGHTTOTAL();
        }
        public void CalPACKAGELOSTRATE()
        {
            numericUpDown40.Value = numericUpDown28.Value / numericUpDown30.Value * 100;
        }
        private void numericUpDown28_ValueChanged(object sender, EventArgs e)
        {
            CalPACKAGELOSTRATE();
        }

        private void numericUpDown30_ValueChanged(object sender, EventArgs e)
        {
            CalPACKAGELOSTRATE();
        }
        public void CalTOTALRATE()
        {
            numericUpDown41.Value = 1 - ((numericUpDown26.Value + numericUpDown27.Value) / numericUpDown31.Value * 100);
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














        #endregion


    }
}
