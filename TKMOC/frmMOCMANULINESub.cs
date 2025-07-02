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
using TKITDLL;

namespace TKMOC
{
    public partial class frmMOCMANULINESub : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter20 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder20 = new SqlCommandBuilder();
        DataSet ds20 = new DataSet();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();

        string EDITID;
        int result;
        decimal BOXNUMERB;

        public frmMOCMANULINESub()
        {
            InitializeComponent();
        }

        public frmMOCMANULINESub(string ID)
        {
            EDITID = ID;
            InitializeComponent();

            SEARCHMOCMANULINE();
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


        public void SEARCHMOCMANULINE()
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

                sbSql.Clear();
                sbSqlQuery.Clear();



            
                sbSql.AppendFormat(@"  
                                     SELECT 
                                     [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'
                                     ,[MB003] AS '規格',ISNULL([BAR],0) AS '桶數',ISNULL([NUM],0) AS '數量',ISNULL([BOX],0)   AS '箱數'   ,ISNULL([PACKAGE],0)  AS '片數',[CLINET] AS '客戶'
                                     ,[MC004],CONVERT(varchar(100),[OUTDATE],112)  AS '交期',[TA029] AS '備註' ,[HALFPRO] AS '半成品數量'
                                     ,[MANUHOUR] AS 生產時間 
                                     ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'
                                    ,[MANUPRENUMS] AS '需多投數量做底'
                                     ,[ID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].[dbo].[BOMMC]
                                     WHERE [MB001]=[MC001]
                                     AND  [ID]='{0}'
                                    ", EDITID);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        textBox1.Text = ds1.Tables["TEMPds1"].Rows[0]["線別"].ToString();
                        
                        textBox3.Text = ds1.Tables["TEMPds1"].Rows[0]["品號"].ToString();
                        textBox4.Text = ds1.Tables["TEMPds1"].Rows[0]["品名"].ToString();
                        textBox5.Text = ds1.Tables["TEMPds1"].Rows[0]["規格"].ToString();
                        textBox6.Text = ds1.Tables["TEMPds1"].Rows[0]["桶數"].ToString();
                        textBox7.Text = ds1.Tables["TEMPds1"].Rows[0]["數量"].ToString();
                        textBox8.Text = ds1.Tables["TEMPds1"].Rows[0]["箱數"].ToString();
                        textBox9.Text = ds1.Tables["TEMPds1"].Rows[0]["片數"].ToString();
                        textBox10.Text = ds1.Tables["TEMPds1"].Rows[0]["客戶"].ToString();
                        textBox32.Text = ds1.Tables["TEMPds1"].Rows[0]["MC004"].ToString();
                        textBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["備註"].ToString();
                        textBox13.Text = ds1.Tables["TEMPds1"].Rows[0]["生產時間"].ToString();
                        textBox12.Text = ds1.Tables["TEMPds1"].Rows[0]["半成品數量"].ToString();
                        textBox40.Text = ds1.Tables["TEMPds1"].Rows[0]["訂單單別"].ToString();
                        textBox41.Text = ds1.Tables["TEMPds1"].Rows[0]["訂單號"].ToString();
                        textBox42.Text = ds1.Tables["TEMPds1"].Rows[0]["訂單序號"].ToString();
                        textBox99.Text = ds1.Tables["TEMPds1"].Rows[0]["需多投數量做底"].ToString();

                        string yy = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString().Substring(0, 4);
                        string MM = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString().Substring(4, 2);
                        string dd = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString().Substring(6, 2);

                        dateTimePicker1.Value = Convert.ToDateTime(yy+"/"+MM+"/"+dd);

                        if(!String.IsNullOrEmpty(ds1.Tables["TEMPds1"].Rows[0]["交期"].ToString()))
                        {
                            string OUTyy = ds1.Tables["TEMPds1"].Rows[0]["交期"].ToString().Substring(0, 4);
                            string OUTMM = ds1.Tables["TEMPds1"].Rows[0]["交期"].ToString().Substring(4, 2);
                            string OUTdd = ds1.Tables["TEMPds1"].Rows[0]["交期"].ToString().Substring(6, 2);

                            dateTimePicker2.Value = Convert.ToDateTime(OUTyy + "/" + OUTMM + "/" + OUTdd);
                        }
                        else
                        {
                            dateTimePicker2.Format = DateTimePickerFormat.Custom;
                            dateTimePicker2.CustomFormat = " ";
                        }

                        textBoxID.Text = ds1.Tables["TEMPds1"].Rows[0]["ID"].ToString();
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

        public void UPDATEMOCMANULINE()
        {
            StringBuilder SQL_EXE = new StringBuilder();
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
                SQL_EXE.AppendFormat(@"
                                    UPDATE [TKMOC].[dbo].[MOCMANULINE] 
                                    SET 
                                        [BAR] = @BAR,
                                        [NUM] = @NUM,
                                        [BOX] = @BOX,
                                        [PACKAGE] = @PACKAGE,
                                        [CLINET] = @CLINET,
                                        [MANUDATE] = @MANUDATE,
                                        [OUTDATE] = @OUTDATE,
                                        [TA029] = @TA029,
                                        [MANUHOUR] = @MANUHOUR,
                                        [HALFPRO] = @HALFPRO,
                                        [COPTD001] = @COPTD001,
                                        [COPTD002] = @COPTD002,
                                        [COPTD003] = @COPTD003,
                                        [MANUPRENUMS] = @MANUPRENUMS
                                    WHERE [ID] = @ID
                                    ");
              

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = SQL_EXE.ToString();
                cmd.Transaction = tran;

                // 加入參數
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@ID", textBoxID.Text);
                cmd.Parameters.AddWithValue("@BAR", textBox6.Text);
                cmd.Parameters.AddWithValue("@NUM", textBox7.Text);
                cmd.Parameters.AddWithValue("@BOX", textBox8.Text);
                cmd.Parameters.AddWithValue("@PACKAGE", textBox9.Text);
                cmd.Parameters.AddWithValue("@CLINET", textBox10.Text);
                cmd.Parameters.AddWithValue("@MANUDATE", dateTimePicker1.Value.ToString("yyyyMMdd"));
                cmd.Parameters.AddWithValue("@OUTDATE", dateTimePicker2.Value.ToString("yyyyMMdd"));
                cmd.Parameters.AddWithValue("@TA029", textBox2.Text);
                cmd.Parameters.AddWithValue("@MANUHOUR", textBox13.Text);
                cmd.Parameters.AddWithValue("@HALFPRO", textBox12.Text);
                cmd.Parameters.AddWithValue("@COPTD001", textBox40.Text);
                cmd.Parameters.AddWithValue("@COPTD002", textBox41.Text);
                cmd.Parameters.AddWithValue("@COPTD003", textBox42.Text);
                cmd.Parameters.AddWithValue("@MANUPRENUMS", textBox99.Text);

                // 執行 SQL
                int result = cmd.ExecuteNonQuery();
                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    MessageBox.Show("更新失敗");
                }
                else
                {
                    tran.Commit();      //執行交易  
                    this.Close();
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
        public void CALPRODUCTDETAIL()
        {
            Decimal num1;
            Decimal num2;
            try
            {
                if (Decimal.TryParse(textBox7.Text, out num1) && Decimal.TryParse(textBox99.Text, out num1) && Decimal.TryParse(textBox32.Text, out num2))
                {
                    if(Convert.ToDecimal(textBox7.Text)>=0 & Convert.ToDecimal(textBox99.Text) >= 0 & Convert.ToDecimal(textBox32.Text)>0)
                    {
                        textBox6.Text = Math.Round((Convert.ToDecimal(textBox7.Text)+ Convert.ToDecimal(textBox99.Text)) / Convert.ToDecimal(textBox32.Text), 4).ToString();
                    }
                    else
                    {
                        textBox6.Text = "0";
                    }
                    
                }
                
                if (Decimal.TryParse(textBox9.Text, out num1) && Decimal.TryParse(textBox32.Text, out num2))
                {
                    if (Convert.ToDecimal(textBox9.Text) > 0 & Convert.ToDecimal(textBox32.Text) > 0 & BOXNUMERB>0)
                    {
                        textBox8.Text = Math.Round(Convert.ToDecimal(textBox9.Text) / Convert.ToDecimal(textBox32.Text) / BOXNUMERB, 4).ToString();
                    }
                    else
                    {
                        textBox8.Text = "0";
                    }
                    
                }

            }
            catch
            {
                //MessageBox.Show("請填數字");
            }
            finally
            {

            }

        }
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }

        private void textBox99_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001BOX();
            textBox11.Text = BOXNUMERB.ToString();

            CALPRODUCTDETAIL();
        }

        public void SEARCHMB001BOX()
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

                sbSql.Clear();
                sbSqlQuery.Clear();

             
                sbSql.AppendFormat(@" 
                                        SELECT  MD001,MD003,MB001,MB002,ISNULL(MD007,1) AS MD007,ISNULL(MD010,1) AS MD010,ISNULL(MD006,1) AS MD006
                                        FROM [TK].dbo.BOMMD,[TK].dbo.INVMB
                                        WHERE MD003=MB001
                                        AND MB002 LIKE '%箱%'
                                        AND MD003 LIKE '2%'
                                        AND MD001 LIKE '{0}%'
                                    ", textBox3.Text);

                adapter20 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder20 = new SqlCommandBuilder(adapter20);
                sqlConn.Open();
                ds20.Clear();
                adapter20.Fill(ds20, "TEMPds20");
                sqlConn.Close();


                if (ds20.Tables["TEMPds20"].Rows.Count == 0)
                {
                    BOXNUMERB = 1;
                }
                else
                {
                    if (ds20.Tables["TEMPds20"].Rows.Count >= 1)
                    {
                        if(Convert.ToDecimal(ds20.Tables["TEMPds20"].Rows[0]["MD007"].ToString())>0 & Convert.ToDecimal(ds20.Tables["TEMPds20"].Rows[0]["MD006"].ToString())>0)
                        {
                            BOXNUMERB = (Convert.ToDecimal(ds20.Tables["TEMPds20"].Rows[0]["MD007"].ToString()) / Convert.ToDecimal(ds20.Tables["TEMPds20"].Rows[0]["MD006"].ToString()));
                        }
                        else
                        {
                            BOXNUMERB = 1;
                        }
                        
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

        public void UPDATE_SPECIAL_MODIFY(string ID,string MB001,string NUM,string MANUDATE,string MANU,string COPTD001, string COPTD002, string COPTD003)
        {
            DataTable DT = FIND_MOCMANULINEBATCHMODIFYS();
            if(DT!=null && DT.Rows.Count>=1)
            {
                string CHECKMB001 = "";
                foreach(DataRow DR in DT.Rows)
                {
                    CHECKMB001 = DR["MB001"].ToString();

                    //40806040000021 可可小布雪180g
                    if (MB001.Equals(CHECKMB001))
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

                            sbSql.AppendFormat(@" 
                                        UPDATE [TKMOC].[dbo].[MOCMANULINE]
                                        SET 
                                        [NUM] = ROUND(TEMP.CALNUMS * {1}, 3)
                                        ,[BAR]=TEMP.CALNUMS * {1}/TEMP.BASEBARS
                                        FROM 
                                        (
                                            SELECT 
                                                MC1.MC001 AS MC1MC001
                                                ,MC1.MC004 AS MC1MC004
                                                ,MD1.MD003 AS MD1MD003
                                                ,MD1.MD006 AS MD1MD006
                                                ,MD1.MD007 AS MD1MD007
                                                ,MD1.MD008 AS MD1MD008
                                                ,((1.0 / MC1.MC004) * MD1.MD006 / MD1.MD007 * (1 + MD1.MD008)) AS CALNUMS
                                                ,MC2.MC004 AS 'BASEBARS'
                                            FROM [TK].dbo.BOMMC MC1
                                            JOIN [TK].dbo.BOMMD MD1 ON MC1.MC001 = MD1.MD001
	                                        JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                                            WHERE MC1.MC001 = '{0}'
                                            AND MD1.MD003 LIKE '3%'
                                        ) AS TEMP
                                        WHERE TEMP.MD1MD003 = [TKMOC].[dbo].[MOCMANULINE].[MB001]
                                        AND [MANU] = '{3}'
                                        AND CONVERT(nvarchar, [MANUDATE], 112) = '{2}'
                                        AND [COPTD001]='{4}' AND [COPTD002]='{5}' AND [COPTD003]='{6}'

                                        UPDATE [TKMOC].[dbo].[MOCMANULINE]
                                        SET 
                                        [NUM] = ROUND(TEMP.CALNUMS2 * {1}, 3)
                                        ,[BAR]=TEMP.CALNUMS2 *  {1}/TEMP.BASEBARS
                                        FROM 
                                         (SELECT 
	                                        MC1.MC001 AS 'MC1MC001'
	                                        ,MC1.MC004  AS 'MC1MC004'
	                                        ,MD1.MD003 AS 'MD1MD003'
	                                        ,MD1.MD006 AS 'MD1MD006'
	                                        ,MD1.MD007 AS 'MD1MD007'
	                                        ,MD1.MD008 AS 'MD1MD008'
	                                        ,((1/MC1.MC004)*MD1.MD006/MD1.MD007*(1+MD1.MD008)) AS 'CALNUMS'
	                                        ,MC2.MC004  AS 'MC2MC004'
	                                        ,MD2.MD003 AS 'MD2MD003'
	                                        ,MD2.MD006 AS 'MD2MD006'
	                                        ,MD2.MD007 AS 'MD2MD007'
	                                        ,MD2.MD008 AS 'MD2MD008'
	                                        ,(((1/MC1.MC004)*MD1.MD006/MD1.MD007*(1+MD1.MD008))/MC2.MC004*MD2.MD006/MD2.MD007*(1+MD2.MD008)) AS 'CALNUMS2'
                                            ,MC3.MC004 AS 'BASEBARS'
	                                        FROM [TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
	                                        JOIN [TK].dbo.BOMMC MC2 ON MD1.MD003=MC2.MC001
	                                        JOIN [TK].dbo.BOMMD MD2 ON MD1.MD003=MD2.MD001
                                            JOIN [TK].dbo.BOMMC MC3 ON MC3.MC001=MD2.MD003
	                                        WHERE MC1.MC001=MD1.MD001
	                                        AND MC1.MC001 ='{0}'
	                                        AND MD1.MD003 LIKE '3%'
                                        ) AS TEMP
                                        WHERE TEMP.MD2MD003=[MOCMANULINE].[MB001]
                                        AND [MANU] = '{3}'
                                        AND CONVERT(nvarchar,[MANUDATE],112)='{2}'     
                                        AND [COPTD001]='{4}' AND [COPTD002]='{5}' AND [COPTD003]='{6}'               
                                        ", CHECKMB001, NUM, MANUDATE, MANU, COPTD001, COPTD002, COPTD003);


                            cmd.Connection = sqlConn;
                            cmd.CommandTimeout = 60;
                            cmd.CommandText = sbSql.ToString();
                            cmd.Transaction = tran;
                            result = cmd.ExecuteNonQuery();

                            if (result == 0)
                            {
                                tran.Rollback();    //交易取消
                                MessageBox.Show("更新失敗");
                            }
                            else
                            {
                                tran.Commit();      //執行交易  
                                MessageBox.Show("已連動更新數量");

                                this.Close();
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
        }

        public DataTable FIND_MOCMANULINEBATCHMODIFYS()
        {
            DataSet DS1 = new DataSet();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

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


                sbSql.AppendFormat(@" 
                                    SELECT
                                    [MB001]
                                    ,[MB002]
                                    FROM [TKMOC].[dbo].[MOCMANULINEBATCHMODIFYS]
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                DS1.Clear();
                adapter1.Fill(DS1, "DS1");
                sqlConn.Close();


                if (DS1.Tables["DS1"].Rows.Count >= 1)
                {
                    return DS1.Tables["DS1"];
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
            finally
            {

            }

        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            UPDATEMOCMANULINE();

            //特殊規則-[TKMOC].[dbo].[MOCMANULINEBATCHMODIFYS]
            //限2層BOM可用批次修改
            //3層以上BOM要另外寫
            //小布雪專用
            UPDATE_SPECIAL_MODIFY(textBoxID.Text.Trim(), textBox3.Text.Trim(), textBox7.Text.Trim(), dateTimePicker1.Value.ToString("yyyyMMdd"),textBox1.Text.Trim(),textBox40.Text.Trim(), textBox41.Text.Trim(), textBox42.Text.Trim());

        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        #endregion

        
    }
}
