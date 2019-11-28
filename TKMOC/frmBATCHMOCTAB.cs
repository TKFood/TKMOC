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

namespace TKMOC
{
    public partial class frmBATCHMOCTAB : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();

        List<ADDITEM> ADDTARGET = new List<ADDITEM>();
        List<ADDITEM> FIND = new List<ADDITEM>();

        public class ADDITEM
        {
            public string MB001;
            public double NUM;

        }

        public frmBATCHMOCTAB()
        {
            InitializeComponent();
        }

        public void TEST()
        {
            ADDTARGET.Add(new ADDITEM { MB001 = "40101110430280", NUM =100 });

            SERACH(ADDTARGET[0].MB001, ADDTARGET[0].NUM);

            foreach (var find in FIND)
            {
                MessageBox.Show(find.MB001 + " " + find.NUM);
            }
        }

        public void SERACH(string MB001,double NUM)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  WITH NODE (MD001,MD003,LAYER,MD004,MC004,PREMC004,MD006,MD007,MD008,USEDNUM ,PREUSEDNUM) AS");
                sbSql.AppendFormat(@"  (");
                sbSql.AppendFormat(@"  SELECT MD001,MD003,0 ,[MD004],[MC004],[MC004] AS PREMC004,[MD006],[MD007],[MD008],CONVERT(DECIMAL(18,4),([MD006]/[MD007]/[MC004]*(1+MD008))),CONVERT(DECIMAL(18,4),1) AS PREUSEDNUM  FROM [TK].[dbo].[VBOMMD]");
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT TB1.MD001,TB2.MD003,TB2.LAYER+1,TB2.MD004,TB2.MC004,TB1.MC004,TB2.MD006,TB2.MD007,TB2.MD008,TB2.USEDNUM,CONVERT(DECIMAL(18,4),(TB1.[MD006]/TB1.[MD007]/TB1.[MC004]*(1+TB1.MD008))) AS PREUSEDNUM FROM [TK].[dbo].[VBOMMD] TB1");
                sbSql.AppendFormat(@"  INNER JOIN NODE TB2");
                sbSql.AppendFormat(@"  ON TB1.MD003 = TB2.MD001");
                sbSql.AppendFormat(@"  )");
                sbSql.AppendFormat(@"  SELECT MD001,MD003,LAYER,MD004,MC004,PREMC004,MD006,MD007,MD008,USEDNUM ,PREUSEDNUM ,USEDNUM*PREUSEDNUM AS TOTALUSED FROM NODE");
                sbSql.AppendFormat(@"  WHERE  MD001='{0}'",MB001);
                sbSql.AppendFormat(@"  ORDER BY LAYER ,MD001, MD003");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        foreach (DataRow od in ds1.Tables["ds1"].Rows)
                        {
                            FIND.Add(new ADDITEM { MB001 = od["MD003"].ToString(), NUM =Convert.ToDouble(od["TOTALUSED"].ToString()) });
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

    


        #region FUNCTION

        #endregion

        #region BUTTON
        private void button2_Click(object sender, EventArgs e)
        {
            TEST();
        }
        #endregion
    }
}
