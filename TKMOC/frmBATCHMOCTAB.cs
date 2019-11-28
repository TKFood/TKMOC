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

            //foreach (var find in FIND)
            //{
            //    MessageBox.Show(find.MB001 + " " + find.NUM);
            //}
        }

        public void SERACH(string MB001,double NUM)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MD003,ROUND({0}*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),3) AS MD004",NUM);
                sbSql.AppendFormat(@"  FROM [TK].dbo.[BOMMD],[TK].dbo.[INVMB]");
                sbSql.AppendFormat(@"  WHERE [BOMMD].MD003=[INVMB].MB001");
                sbSql.AppendFormat(@"  AND MD001='{0}'",MB001);
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
                            FIND.Add(new ADDITEM { MB001 = od["MD003"].ToString(), NUM =Convert.ToDouble(od["MD004"].ToString()) });
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

        public void SERACHFIND(string MB001, double NUM)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MD003,ROUND({0}*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),3) AS MD004", NUM);
                sbSql.AppendFormat(@"  FROM [TK].dbo.[BOMMD],[TK].dbo.[INVMB]");
                sbSql.AppendFormat(@"  WHERE [BOMMD].MD003=[INVMB].MB001");
                sbSql.AppendFormat(@"  AND MD001='{0}'", MB001);
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
                        ADDTARGET.Add(new ADDITEM { MB001 = MB001, NUM = NUM });

                        foreach (DataRow od in ds1.Tables["ds1"].Rows)
                        {
                            FIND.Add(new ADDITEM { MB001 = od["MD003"].ToString(), NUM = Convert.ToDouble(od["MD004"].ToString()) });
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
