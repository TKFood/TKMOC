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
    public partial class frmMOCMANULINE : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        StringBuilder sbSqlQuery2 = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();
        SqlDataAdapter adapter8 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder8 = new SqlCommandBuilder();
        SqlDataAdapter adapter9 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder9 = new SqlCommandBuilder();
        SqlDataAdapter adapter10 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder10 = new SqlCommandBuilder();
        SqlDataAdapter adapter11 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder11 = new SqlCommandBuilder();
        SqlDataAdapter adapter12 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder12 = new SqlCommandBuilder();
        SqlDataAdapter adapter13 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder13 = new SqlCommandBuilder();
        SqlDataAdapter adapter14 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder14 = new SqlCommandBuilder();
        SqlDataAdapter adapter15 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder15 = new SqlCommandBuilder();
        SqlDataAdapter adapter16 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder16 = new SqlCommandBuilder();
        SqlDataAdapter adapter17 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder17 = new SqlCommandBuilder();
        SqlDataAdapter adapter18 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder18 = new SqlCommandBuilder();
        SqlDataAdapter adapter19 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder19 = new SqlCommandBuilder();
        SqlDataAdapter adapter20= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder20 = new SqlCommandBuilder();
        SqlDataAdapter adapter21 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder21= new SqlCommandBuilder();
        SqlDataAdapter adapter22 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder22 = new SqlCommandBuilder();
        SqlDataAdapter adapter23 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder23 = new SqlCommandBuilder();
        SqlDataAdapter adapter24 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder24 = new SqlCommandBuilder();
        SqlDataAdapter adapter25 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder25 = new SqlCommandBuilder();
        SqlDataAdapter adapter26 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder26 = new SqlCommandBuilder();
        SqlDataAdapter adapter27 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder27 = new SqlCommandBuilder();
        SqlDataAdapter adapter28 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder28 = new SqlCommandBuilder();
        SqlDataAdapter adapter29 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder29 = new SqlCommandBuilder();
        SqlDataAdapter adapter30 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder30 = new SqlCommandBuilder();
        SqlDataAdapter adapter31= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder31 = new SqlCommandBuilder();
        SqlDataAdapter adapter32 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder32 = new SqlCommandBuilder();
        SqlDataAdapter adapter33 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder33 = new SqlCommandBuilder();
        SqlDataAdapter adapter34 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder34 = new SqlCommandBuilder();
        SqlDataAdapter adapter35 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder35 = new SqlCommandBuilder();
        SqlDataAdapter adapter36 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder36 = new SqlCommandBuilder();
        SqlDataAdapter adapter37 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder37 = new SqlCommandBuilder();

        SqlDataAdapter adapterCALENDAR = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderCALENDAR = new SqlCommandBuilder();




        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();
        DataSet ds9 = new DataSet();
        DataSet ds10= new DataSet();
        DataSet ds13 = new DataSet();
        DataSet ds14 = new DataSet();
        DataSet ds15 = new DataSet();
        DataSet ds16 = new DataSet();
        DataSet ds17 = new DataSet();
        DataSet ds18 = new DataSet();
        DataSet ds19 = new DataSet();
        DataSet ds20 = new DataSet();
        DataSet ds21 = new DataSet();
        DataSet ds22 = new DataSet();
        DataSet ds23 = new DataSet();
        DataSet ds24 = new DataSet();
        DataSet ds25 = new DataSet();
        DataSet ds26 = new DataSet();
        DataSet ds27 = new DataSet();
        DataSet ds28 = new DataSet();
        DataSet ds29 = new DataSet();
        DataSet ds30 = new DataSet();
        DataSet ds31 = new DataSet();
        DataSet ds32 = new DataSet();
        DataSet ds33 = new DataSet();
        DataSet ds34 = new DataSet();
        DataSet ds35 = new DataSet();
        DataSet ds36 = new DataSet();
        DataSet ds37 = new DataSet();


        DataSet dsCALENDAR = new DataSet();

        DataSet dsBOMMC = new DataSet();
        DataSet dsBOMMD = new DataSet();

        DataSet TEMPds = new DataSet();
        decimal SUM11 = 0;
        decimal SUM21 = 0;
        decimal SUM31 = 0;
        decimal SUM41 = 0;


        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string MANU= "新廠製二組";

        string ID1;
        DateTime dt1;
        string DELID1;
        string DELMOCTA001A;
        string DELMOCTA002A;
        string IN1="20001";
        string ID2;
        DateTime dt2;
        string DELID2;
        string DELMOCTA001B;
        string DELMOCTA002B;
        string IN2 = "20001";
        string ID3;
        DateTime dt3;
        string DELID3;
        string DELMOCTA001C;
        string DELMOCTA002C;
        string IN3 = "20001";
        string ID4;
        DateTime dt4;
        string DELID4;
        string DELMOCTA001D;
        string DELMOCTA002D;
        string IN4 = "20001";
        DateTime dt5;
        string DELID5;
        string DELMOCTA001E;
        string DELMOCTA002E;

        string ID6;
        DateTime dt6;
        string DELID6;
        string DELMOCTA001F;
        string DELMOCTA002F;
        string IN6 = "20021";

        string ID10;
        DateTime dt10;
        string DELID10;
        string DELMOCTA001J;
        string DELMOCTA002J;
        string IN10 = "20021";

        string TA001 = "A510";
        string TA002;
        string TA029;
        string MB001;
        string MB002;
        string MB003;
        string MB001B;
        string MB002B;
        string MB003B;
        string MB001C;
        string MB002C;
        string MB003C;
        string MB001D;
        string MB002D;
        string MB003D;
        string MB001E;
        string MB002E;
        string MB003E;
        string MB001F;
        string MB002F;
        string MB003F;
        decimal BAR;
        decimal BOX;
        decimal BAR2;
        decimal BAR3;
        decimal SUM1;
        decimal SUM2;
        decimal SUM3;
        decimal SUM4;
        decimal SUM5;

        string BOMVARSION;
        string UNIT;
        decimal BOMBAR;
        int BOXNUMERB;
        int MOCBOX;

        string SUBID;
        string SUBBAR;
        string SUBNUM;
        string SUBBOX;
        string SUBPACKAGE;
        string SUBID2;
        string SUBBAR2;
        string SUBNUM2;
        string SUBBOX2;
        string SUBPACKAGE2;
        string SUBID3;
        string SUBBAR3;
        string SUBNUM3;
        string SUBBOX3;
        string SUBPACKAGE3;
        string SUBID4;
        string SUBBAR4;
        string SUBNUM4;
        string SUBBOX4;
        string SUBPACKAGE4;
        string SUBID5;
        string SUBBAR5;
        string SUBNUM5;
        string SUBBOX5;
        string SUBPACKAGE5;

        string TA026;
        string TA027;
        string TA028;
        string TA026A;
        string TA027A;
        string TA028A;
        string TA026B;
        string TA027B;
        string TA028B;
        string TA026C;
        string TA027C;
        string TA028C;
        string TA026D;
        string TA027D;
        string TA028D;

        string DELMOCMANULINECOPID;
        string LIMITSERCHTD002;

        public class MOCTADATA
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_count;
            public string DataGroup;
            public string TA001;
            public string TA002;
            public string TA003;
            public string TA004;
            public string TA005;
            public string TA006;
            public string TA007;
            public string TA009;
            public string TA010;
            public string TA011;
            public string TA012;
            public string TA013;
            public string TA014;
            public string TA015;
            public string TA016;
            public string TA017;
            public string TA018;
            public string TA019;
            public string TA020;
            public string TA021;
            public string TA022;
            public string TA024;
            public string TA025;
            public string TA026;
            public string TA027;
            public string TA028;
            public string TA029;
            public string TA030;
            public string TA031;
            public string TA033;
            public string TA034;
            public string TA035;
            public string TA040;
            public string TA041;
            public string TA043;
            public string TA044;
            public string TA045;
            public string TA046;
            public string TA047;
            public string TA049;
            public string TA050;
            public string TA200;
        }

        public class MOCTBDATA
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_count;
            public string DataGroup;
            
        }

        Thread TD;
        public frmMOCMANULINE()
        {
            InitializeComponent();
            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();
            comboBox5load();
            comboBox6load();
            comboBox7load();
            comboBox7load();
            comboBox7load();
            comboBox7load();
            comboBox7load();
            comboBox8load();
            comboBox12load();

            comboBox13load();
            comboBox14load();
            comboBox15load();
            comboBox16load();
            comboBox17load();


            comboBox19load();

            SETIN();

            //SET CALENDAR
            SETCALENDAR();

            MANU = "新廠包裝線";
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

        private void frmMOCMANULINE_Load(object sender, EventArgs e)
        {
            dataGridView20.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView20.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

            ////建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            //Rectangle rect = dataGridView3.GetCellDisplayRectangle(0, -1, true);
            //rect.X = rect.Location.X + rect.Width / 4 - 18;
            //rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            //CheckBox cbHeader = new CheckBox();
            //cbHeader.Name = "checkboxHeader";
            //cbHeader.Size = new Size(18, 18);
            //cbHeader.Location = rect.Location;

            ////全选要设定的事件
            //cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            ////将 CheckBox 加入到 dataGridView
            //dataGridView3.Controls.Add(cbHeader);


            #endregion
        }
        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠製二組%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD002";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠包裝線%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD002";
            comboBox2.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox3load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠製一組%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "MD002";
            comboBox3.DisplayMember = "MD002";
            sqlConn.Close();


        }
        public void comboBox4load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠製三組(手工)%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "MD002";
            comboBox4.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox5load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'  ORDER BY MC001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "MC001";
            comboBox5.DisplayMember = "MC002";
            sqlConn.Close();

           

        }
        public void comboBox6load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'  ORDER BY MC001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            da.Fill(dt);
            comboBox6.DataSource = dt.DefaultView;
            comboBox6.ValueMember = "MC001";
            comboBox6.DisplayMember = "MC002";
            sqlConn.Close();

           
        }
        public void comboBox7load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'  ORDER BY MC001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            da.Fill(dt);
            comboBox7.DataSource = dt.DefaultView;
            comboBox7.ValueMember = "MC001";
            comboBox7.DisplayMember = "MC002";
            sqlConn.Close();

           


        }
        public void comboBox8load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'  ORDER BY MC001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            da.Fill(dt);
            comboBox8.DataSource = dt.DefaultView;
            comboBox8.ValueMember = "MC001";
            comboBox8.DisplayMember = "MC002";
            sqlConn.Close();

           


        }

        public void comboBox12load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MQ001,MQ002 FROM [TK].dbo.CMSMQ WHERE MQ003='22' ORDER BY MQ001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MQ001", typeof(string));
            dt.Columns.Add("MQ002", typeof(string));

            da.Fill(dt);
            comboBox12.DataSource = dt.DefaultView;
            comboBox12.ValueMember = "MQ001";
            comboBox12.DisplayMember = "MQ002";
            sqlConn.Close();




        }

        public void comboBox13load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠統百包裝線%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox13.DataSource = dt.DefaultView;
            comboBox13.ValueMember = "MD002";
            comboBox13.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox14load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '20021%'  ORDER BY MC001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            da.Fill(dt);
            comboBox14.DataSource = dt.DefaultView;
            comboBox14.ValueMember = "MC001";
            comboBox14.DisplayMember = "MC002";
            sqlConn.Close();


        }

        public void comboBox15load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox15.DataSource = dt.DefaultView;
            comboBox15.ValueMember = "MD002";
            comboBox15.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox16load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'  ORDER BY MC001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            da.Fill(dt);
            comboBox16.DataSource = dt.DefaultView;
            comboBox16.ValueMember = "MC001";
            comboBox16.DisplayMember = "MC002";
            sqlConn.Close();



        }

        public void comboBox17load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox17.DataSource = dt.DefaultView;
            comboBox17.ValueMember = "MD001";
            comboBox17.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox19load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox19.DataSource = dt.DefaultView;
            comboBox19.ValueMember = "MD002";
            comboBox19.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void SEARCHMOCMANULINE()
        {
            if(MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名' ");
                    sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BAR] AS '桶數',[NUM] AS '數量',[CLINET] AS '客戶',[OUTDATE] AS '交期',[TA029] AS '備註',[HALFPRO] AS '半成品數量'");
                    sbSql.AppendFormat(@"  ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'");
                    sbSql.AppendFormat(@"  ,[ID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  WHERE [MANU]='{0}' ", MANU);
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[SERNO]");
                    sbSql.AppendFormat(@"  ");

                    adapter1= new SqlDataAdapter(@"" + sbSql, sqlConn);

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
                            dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                            dataGridView1.AutoResizeColumns();
                            //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'");
                    sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BOX] AS '箱數',[PACKAGE] AS '包裝數',[CLINET] AS '客戶',[MANUHOUR] AS '生產時間',[OUTDATE] AS '交期',[TA029] AS '備註',[HALFPRO] AS '半成品數量'");
                    sbSql.AppendFormat(@"  ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'");
                    sbSql.AppendFormat(@"  ,[ID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  WHERE [MANU]='{0}' ", MANU);
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker3.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[SERNO]");
                    sbSql.AppendFormat(@"  ");

                    adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                    sqlConn.Open();
                    ds5.Clear();
                    adapter7.Fill(ds5, "TEMPds5");
                    sqlConn.Close();


                    if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds5.Tables["TEMPds5"].Rows.Count >= 1)
                        {
                            //dataGridView1.Rows.Clear();
                            dataGridView3.DataSource = ds5.Tables["TEMPds5"];
                            dataGridView3.AutoResizeColumns();
                            //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'");
                    sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BAR] AS '桶數',[NUM] AS '數量',[CLINET] AS '客戶',[OUTDATE] AS '交期',[TA029] AS '備註',[HALFPRO] AS '半成品數量'");
                    sbSql.AppendFormat(@"  ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'");
                    sbSql.AppendFormat(@"  ,[ID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  WHERE [MANU]='{0}' ", MANU);
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker6.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[SERNO]");
                    sbSql.AppendFormat(@"  ");

                    adapter9 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder9 = new SqlCommandBuilder(adapter9);
                    sqlConn.Open();
                    ds7.Clear();
                    adapter9.Fill(ds7, "TEMPds7");
                    sqlConn.Close();


                    if (ds7.Tables["TEMPds7"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds7.Tables["TEMPds7"].Rows.Count >= 1)
                        {
                            //dataGridView1.Rows.Clear();
                            dataGridView5.DataSource = ds7.Tables["TEMPds7"];
                            dataGridView5.AutoResizeColumns();
                            //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'");
                    sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BAR] AS '桶數',[NUM] AS '數量',[CLINET] AS '客戶',[OUTDATE] AS '交期',[TA029] AS '備註',[HALFPRO] AS '半成品數量'");
                    sbSql.AppendFormat(@"  ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'");
                    sbSql.AppendFormat(@"  ,[ID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  WHERE [MANU]='{0}' ", MANU);
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker8.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[SERNO]");
                    sbSql.AppendFormat(@"  ");

                    adapter10 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder10 = new SqlCommandBuilder(adapter10);
                    sqlConn.Open();
                    ds8.Clear();
                    adapter10.Fill(ds8, "TEMPds8");
                    sqlConn.Close();


                    if (ds8.Tables["TEMPds8"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds8.Tables["TEMPds8"].Rows.Count >= 1)
                        {
                            //dataGridView1.Rows.Clear();
                            dataGridView7.DataSource = ds8.Tables["TEMPds8"];
                            dataGridView7.AutoResizeColumns();
                            //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

            else if (MANU.Equals("新廠統百包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'");
                    sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BOX] AS '箱數',[PACKAGE] AS '包裝數',[CLINET] AS '客戶',[MANUHOUR] AS '生產時間',[OUTDATE] AS '交期',[TA029] AS '備註',[HALFPRO] AS '半成品數量'");
                    sbSql.AppendFormat(@"  ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'");
                    sbSql.AppendFormat(@"  ,[ID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  WHERE [MANU]='{0}' ", MANU);
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker17.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[SERNO]");
                    sbSql.AppendFormat(@"  ");

                    adapter23 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder23 = new SqlCommandBuilder(adapter23);
                    sqlConn.Open();
                    ds23.Clear();
                    adapter23.Fill(ds23, "TEMPds23");
                    sqlConn.Close();


                    if (ds23.Tables["TEMPds23"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds23.Tables["TEMPds23"].Rows.Count >= 1)
                        {
                            //dataGridView1.Rows.Clear();
                            dataGridView16.DataSource = ds23.Tables["TEMPds23"];
                            dataGridView16.AutoResizeColumns();
                            //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        }

        public void SEARCHMOCMANULINETEMP(string STATUS,string TD002)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                sbSqlQuery2.Clear();

                if (STATUS.Equals("否"))
                {
                    sbSqlQuery.AppendFormat(@" WHERE  [TID] IS NULL ");
                }
                else if (STATUS.Equals("是"))
                {
                    sbSqlQuery.AppendFormat(@"WHERE [TID] IS NOT NULL ");
                }
                else
                {
                    sbSqlQuery.AppendFormat(@" WHERE 1=1 ");
                }

                if(!string.IsNullOrEmpty(TD002))
                {
                    sbSqlQuery2.AppendFormat(@" AND   [MOCMANULINETEMP].[COPTD002] LIKE '%{0}%'", TD002);
                }
                else
                {
                    sbSqlQuery2.AppendFormat(@" ");
                }

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MOCMANULINETEMP].[MANU] AS '線別',CONVERT(varchar(100),[MOCMANULINETEMP].[MANUDATE],112) AS '生產日',[MOCMANULINETEMP].[MB001] AS '品號',[MOCMANULINETEMP].[MB002] AS '品名' ");
                sbSql.AppendFormat(@"  ,[MOCMANULINETEMP].[MB003] AS '規格',[MOCMANULINETEMP].[NUM] AS '數量',[MOCMANULINETEMP].[BAR] AS '桶數',[MOCMANULINETEMP].[PACKAGE] AS'包裝數',[MOCMANULINETEMP].[BOX] AS'箱數',[MOCMANULINETEMP].[CLINET] AS '客戶',[MOCMANULINETEMP].[OUTDATE] AS '交期',[MOCMANULINETEMP].[TA029] AS '備註',[MOCMANULINETEMP].[HALFPRO] AS '半成品數量'");
                sbSql.AppendFormat(@"  ,[MOCMANULINETEMP].[COPTD001] AS '訂單單別',[MOCMANULINETEMP].[COPTD002] AS '訂單號',[MOCMANULINETEMP].[COPTD003] AS '訂單序號'");
                sbSql.AppendFormat(@"  ,[MOCTA001] AS '製令',[MOCTA002] AS '製令號'");
                sbSql.AppendFormat(@"  ,[MOCMANULINETEMP].[ID],[MOCMANULINETEMP].[TID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINETEMP]");
                sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MOCMANULINE] ON [MOCMANULINE].ID=[MOCMANULINETEMP].[TID]");
                sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MOCMANULINERESULT] ON [MOCMANULINERESULT].[SID]=[MOCMANULINETEMP].[TID]");
                sbSql.AppendFormat(@"  {0}", sbSqlQuery.ToString());
                sbSql.AppendFormat(@"  {0}", sbSqlQuery2.ToString());
                sbSql.AppendFormat(@"  ORDER BY [MOCMANULINETEMP].[MANUDATE],[MOCMANULINETEMP].[SERNO]");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView20.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView20.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView20.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        public void SEARCHMOCMANULINETEMPTD002(string STATUS,string TD002)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if (STATUS.Equals("否"))
                {
                    sbSqlQuery.AppendFormat(@" WHERE  [TID] IS NULL ");
                    sbSqlQuery.AppendFormat(@" AND  [MOCMANULINETEMP].[COPTD002] LIKE '%{0}%'",TD002);
                }
                else if (STATUS.Equals("是"))
                {
                    sbSqlQuery.AppendFormat(@"WHERE [TID] IS NOT NULL ");
                    sbSqlQuery.AppendFormat(@" AND  [MOCMANULINETEMP].[COPTD002] LIKE '%{0}%'", TD002);
                }
                else
                {
                    sbSqlQuery.AppendFormat(@"  ");
                }


                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MOCMANULINETEMP].[MANU] AS '線別',CONVERT(varchar(100),[MOCMANULINETEMP].[MANUDATE],112) AS '生產日',[MOCMANULINETEMP].[MB001] AS '品號',[MOCMANULINETEMP].[MB002] AS '品名' ");
                sbSql.AppendFormat(@"  ,[MOCMANULINETEMP].[MB003] AS '規格',[MOCMANULINETEMP].[NUM] AS '數量',[MOCMANULINETEMP].[BAR] AS '桶數',[MOCMANULINETEMP].[PACKAGE] AS'包裝數',[MOCMANULINETEMP].[BOX] AS'箱數',[MOCMANULINETEMP].[CLINET] AS '客戶',[MOCMANULINETEMP].[OUTDATE] AS '交期',[MOCMANULINETEMP].[TA029] AS '備註',[MOCMANULINETEMP].[HALFPRO] AS '半成品數量'");
                sbSql.AppendFormat(@"  ,[MOCMANULINETEMP].[COPTD001] AS '訂單單別',[MOCMANULINETEMP].[COPTD002] AS '訂單號',[MOCMANULINETEMP].[COPTD003] AS '訂單序號'");
                sbSql.AppendFormat(@"  ,[MOCTA001] AS '製令',[MOCTA002] AS '製令號'");
                sbSql.AppendFormat(@"  ,[MOCMANULINETEMP].[ID],[MOCMANULINETEMP].[TID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINETEMP]");
                sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MOCMANULINE] ON [MOCMANULINE].ID=[MOCMANULINETEMP].[TID]");
                sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MOCMANULINERESULT] ON [MOCMANULINERESULT].[SID]=[MOCMANULINETEMP].[TID]");
                sbSql.AppendFormat(@"  {0}", sbSqlQuery.ToString());
                sbSql.AppendFormat(@"  ORDER BY [MOCMANULINETEMP].[MANUDATE],[MOCMANULINETEMP].[SERNO]");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView20.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView20.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView20.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
            SEARCHBOMMD();

            SEARCHMOCMANULINETEMPDATAS(textBox1.Text.Trim());
        }

        public void SEARCHMB001()
        {         

            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004,MB017            ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", textBox1.Text.Trim());
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                       
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            textBox2.Text = ds2.Tables["TEMPds2"].Rows[0]["MB002"].ToString();
                            textBox3.Text = ds2.Tables["TEMPds2"].Rows[0]["MB003"].ToString();
                            textBox32.Text = ds2.Tables["TEMPds2"].Rows[0]["MC004"].ToString();
                            //comboBox5.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            //label51.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
            else if (MANU.Equals("新廠包裝線"))
            {
 

                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", textBox7.Text.Trim());
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            textBox10.Text = ds2.Tables["TEMPds2"].Rows[0]["MB002"].ToString();
                            textBox11.Text = ds2.Tables["TEMPds2"].Rows[0]["MB003"].ToString();
                            textBox33.Text = ds2.Tables["TEMPds2"].Rows[0]["MC004"].ToString();
                            //comboBox6.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            //label52.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();

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

            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", textBox14.Text.Trim());
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            textBox17.Text = ds2.Tables["TEMPds2"].Rows[0]["MB002"].ToString();
                            textBox18.Text = ds2.Tables["TEMPds2"].Rows[0]["MB003"].ToString();
                            textBox34.Text = ds2.Tables["TEMPds2"].Rows[0]["MC004"].ToString();
                            //comboBox7.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            //label53.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", textBox20.Text.Trim());
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                       
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            textBox24.Text = ds2.Tables["TEMPds2"].Rows[0]["MB002"].ToString();
                            textBox25.Text = ds2.Tables["TEMPds2"].Rows[0]["MB003"].ToString();
                            textBox35.Text = ds2.Tables["TEMPds2"].Rows[0]["MC004"].ToString();
                            //comboBox8.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            //label54.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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

            else if (MANU.Equals("新廠統百包裝線"))
            {


                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", textBox56.Text.Trim());
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            textBox62.Text = ds2.Tables["TEMPds2"].Rows[0]["MB002"].ToString();
                            textBox63.Text = ds2.Tables["TEMPds2"].Rows[0]["MB003"].ToString();
                            textBox60.Text = ds2.Tables["TEMPds2"].Rows[0]["MC004"].ToString();
                            //comboBox6.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            //label52.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();

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

            else if (MANU.Equals("少量訂單"))
            {


                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", textBox731.Text.Trim());
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            textBox721.Text = ds2.Tables["TEMPds2"].Rows[0]["MB002"].ToString();
                            textBox732.Text = ds2.Tables["TEMPds2"].Rows[0]["MB003"].ToString();
                            textBox752.Text = ds2.Tables["TEMPds2"].Rows[0]["MC004"].ToString();
                            //comboBox6.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            //label52.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();

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

        }

        public void SETNULL1()
        {
            //textBox1.Text = null;

            textBox4.Text = null;
            textBox5.Text = null;

            textBox2.Text = null;
            textBox3.Text = null;
            textBox32.Text = null;

          

        }
       
        public void ADDMOCMANULINE()
        {
            Guid NEWGUID = new Guid();
            NEWGUID = Guid.NewGuid();

            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(" ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[TA029],[OUTDATE],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',N'{9}','{10}','{11}','{12}','{13}','{14}')", NEWGUID.ToString(), comboBox1.Text, dateTimePicker2.Value.ToString("yyyy/MM/dd"), textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox52.Text, dateTimePicker14.Value.ToString("yyyy/MM/dd"),textBox67.Text, textBox40.Text, textBox41.Text, textBox73.Text);
                    sbSql.AppendFormat(" ");
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
                        UPDATEMOCMANULINETEMP(NEWGUID, TEMPds);

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
            else if (MANU.Equals("新廠包裝線"))
            {
               

                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(" ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',N'{11}','{12}','{13}','{14}','{15}')", NEWGUID.ToString(), comboBox2.Text, dateTimePicker4.Value.ToString("yyyy/MM/dd"), textBox7.Text, textBox10.Text, textBox11.Text, textBox9.Text, textBox13.Text, textBox8.Text, textBox12.Text, dateTimePicker5.Value.ToString("yyyy/MM/dd"), textBox53.Text,textBox68.Text,textBox42.Text, textBox43.Text, textBox72.Text);
                    sbSql.AppendFormat(" ");
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
                        UPDATEMOCMANULINETEMP(NEWGUID, TEMPds);


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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(" ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[TA029],[OUTDATE],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',N'{9}','{10}','{11}','{12}','{13}','{14}')", NEWGUID.ToString(), comboBox3.Text, dateTimePicker7.Value.ToString("yyyy/MM/dd"), textBox14.Text, textBox17.Text, textBox18.Text, textBox15.Text, textBox19.Text, textBox16.Text, textBox54.Text, dateTimePicker15.Value.ToString("yyyy/MM/dd"),textBox69.Text, textBox44.Text, textBox45.Text, textBox74.Text);
                    sbSql.AppendFormat(" ");
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

                        UPDATEMOCMANULINETEMP(NEWGUID, TEMPds);

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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(" ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[TA029],[OUTDATE],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',N'{9}','{10}','{11}','{12}','{13}','{14}')", NEWGUID.ToString(), comboBox4.Text, dateTimePicker9.Value.ToString("yyyy/MM/dd"), textBox20.Text, textBox24.Text, textBox25.Text, textBox21.Text, textBox23.Text, textBox22.Text, textBox55.Text, dateTimePicker16.Value.ToString("yyyy/MM/dd"),textBox70.Text, textBox46.Text, textBox47.Text, textBox75.Text);
                    sbSql.AppendFormat(" ");
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

                        UPDATEMOCMANULINETEMP(NEWGUID, TEMPds);
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

            else if (MANU.Equals("新廠統百包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(" ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',N'{11}','{12}','{13}','{14}','{15}')", "NEWID()", comboBox13.Text, dateTimePicker18.Value.ToString("yyyy/MM/dd"), textBox56.Text, textBox62.Text, textBox63.Text, textBox57.Text, textBox58.Text, textBox59.Text, textBox61.Text, dateTimePicker19.Value.ToString("yyyy/MM/dd"), textBox64.Text,textBox71.Text, textBox65.Text, textBox66.Text, textBox76.Text);
                    sbSql.AppendFormat(" ");
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

            else if (MANU.Equals("少量訂單"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINETEMP]");
                    sbSql.AppendFormat(" ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[CLINET],[MANUHOUR],[BAR],[NUM],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',N'{11}','{12}','{13}','{14}','{15}','{16}','{17}')", "NEWID()", comboBox19.Text, dateTimePicker23.Value.ToString("yyyy/MM/dd"), textBox731.Text, textBox721.Text, textBox732.Text, textBox761.Text, textBox762.Text, textBox741.Text, textBox742.Text, textBox753.Text, textBox751.Text, dateTimePicker24.Value.ToString("yyyy/MM/dd"), textBox771.Text, textBox772.Text, textBox781.Text, textBox782.Text, textBox783.Text);
                    sbSql.AppendFormat(" ");
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

            SEARCHMOCMANULINE();
        }
        public void SETNULL2()
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox32.Text = "0";
            textBox52.Text = null;
            textBox67.Text = "0";
            textBox40.Text = null;
            textBox41.Text = null;
            textBox73.Text = null;

        }
        public void SETNULL3()
        {
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = "0";
            textBox33.Text = "0";
            textBox53.Text = null;
            textBox68.Text = "0";
            textBox42.Text = null;
            textBox43.Text = null;
            textBox72.Text = null;
        }
        public void SETNULL4()
        {
            textBox14.Text = null;
            textBox15.Text = null;
            textBox16.Text = null;
            textBox17.Text = null;
            textBox18.Text = null;
            textBox19.Text = null;
            textBox34.Text = "0";
            textBox54.Text = null;
            textBox69.Text = "0";
            textBox44.Text = null;
            textBox45.Text = null;
            textBox74.Text = null;
        }

        public void SETNULL6()
        {
            textBox20.Text = null;
            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
            textBox25.Text = null;         
           
            textBox35.Text = "0";
            textBox55.Text = null;
            textBox70.Text = "0";
            textBox46.Text = null;
            textBox47.Text = null;
            textBox75.Text = null;
        }

        public void SETNULL7()
        {
            textBox56.Text = null;
            textBox57.Text = null;
            textBox59.Text = null;
            textBox61.Text = null;
            textBox62.Text = null;
            textBox63.Text = null;
            textBox60.Text = "0";
            textBox58.Text = "0";
            textBox64.Text = null;
            textBox71.Text = "0";
            textBox56.Text = null;
            textBox66.Text = null;
            textBox76.Text = null;
        }

        public void SETNULL8()
        {
            textBox731.Text = null;
            textBox753.Text = null;
            textBox761.Text = null;
            textBox721.Text = null;
            textBox732.Text = null;
            textBox741.Text = null;
            textBox751.Text = null;
            textBox762.Text = "0";
            textBox752.Text = "0";
            textBox771.Text = null;
            textBox772.Text = "0";
            textBox781.Text = null;
            textBox782.Text = null;
            textBox783.Text = null;
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox40.Text = null;
            textBox41.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    ID1 = row.Cells["ID"].Value.ToString();
                    dt1=Convert.ToDateTime (row.Cells["生產日"].Value.ToString().Substring(0,4)+"/"+row.Cells["生產日"].Value.ToString().Substring(4, 2)+"/"+row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001= row.Cells["品號"].Value.ToString();
                    MB002 = row.Cells["品名"].Value.ToString();
                    MB003 = row.Cells["規格"].Value.ToString();
                    BAR = Convert.ToDecimal(row.Cells["桶數"].Value.ToString());
                    SUM1 = Convert.ToDecimal(row.Cells["數量"].Value.ToString());
                    TA029 = row.Cells["備註"].Value.ToString();
                    TA026A = row.Cells["訂單單別"].Value.ToString();
                    TA027A = row.Cells["訂單號"].Value.ToString();
                    TA028A = row.Cells["訂單序號"].Value.ToString();

                    SUBID = row.Cells["ID"].Value.ToString();
                    SUBBAR = row.Cells["桶數"].Value.ToString();
                    SUBNUM = row.Cells["數量"].Value.ToString();
                    SUBBOX= null;
                    SUBPACKAGE = null;

                    SEARCHMB017();
                    SEARCHMOCMANULINERESULT();

                    SEARCHMOCMANULINECOP(ID1);
                    SEARCHMOCMANULINEMERGERESLUTMOCTA(ID1.ToString());
                    //SEARCHMOCMANULINECOP();

                    ;
                }
                else
                {
                    ID1 = null;
                    SUBID = null;
                    SUBBAR = null;
                    SUBNUM = null;
                    SUBBOX = null;
                    SUBPACKAGE = null;

                    TA026A = null;
                    TA027A = null;
                    TA028A = null;

                }
            }
        }
        
        public void DELMOCMANULINE()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat("  WHERE ID='{0}'", ID1);
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

                        UPDATEMOCMANULINETEMPTONULL(ID1);
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

            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat("  WHERE ID='{0}'", ID2);
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

                        UPDATEMOCMANULINETEMPTONULL(ID2);
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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat("  WHERE ID='{0}'", ID3);
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

                        UPDATEMOCMANULINETEMPTONULL(ID3);
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat("  WHERE ID='{0}'", ID4);
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

                        UPDATEMOCMANULINETEMPTONULL(ID4);
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

            else if (MANU.Equals("新廠統百包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat("  WHERE ID='{0}'", ID6);
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

            else if (MANU.Equals("少量訂單"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINETEMP]");
                    sbSql.AppendFormat("  WHERE ID='{0}'", ID10);
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


            SEARCHMOCMANULINE();
        }

        public void ADDMOCMANULINERESULT()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(" ([SID],[MOCTA001],[MOCTA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", ID1, TA001, TA002);
                    sbSql.AppendFormat(" ");
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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(" ([SID],[MOCTA001],[MOCTA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", ID2, TA001, TA002);
                    sbSql.AppendFormat(" ");
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
                
            else if(MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(" ([SID],[MOCTA001],[MOCTA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", ID3, TA001, TA002);
                    sbSql.AppendFormat(" ");
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(" ([SID],[MOCTA001],[MOCTA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", ID4, TA001, TA002);
                    sbSql.AppendFormat(" ");
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

            else if (MANU.Equals("新廠統百包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(" ([SID],[MOCTA001],[MOCTA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", ID6, TA001, TA002);
                    sbSql.AppendFormat(" ");
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
        }

        public void ADDMOCTATB()
        {
            MOCTADATA MOCTA = new MOCTADATA();
            MOCTA = SETMOCTA();
            string MOCMB001 = null;
            decimal MOCTA004 = 0; ;
            string MOCTB009 = null;
       

            const int MaxLength = 100;

            if (MANU.Equals("新廠製二組"))
            {
                MOCMB001 = MB001;
                MOCTA004 = BAR;
                MOCTA.TA026 = TA026A;
                MOCTA.TA027 = TA027A;
                MOCTA.TA028 = TA028A;
                //MOCTB009 = textBox77.Text;

            }
            else if (MANU.Equals("新廠包裝線"))
            {
                MOCMB001 = MB001B;
                MOCTA004 = BOX;
                MOCTA.TA026 = TA026;
                MOCTA.TA027 = TA027;
                MOCTA.TA028 = TA028;
                //MOCTB009 = textBox78.Text;

            }
            else if (MANU.Equals("新廠製一組"))
            {
                MOCMB001 = MB001C;
                MOCTA004 = BAR2;
                MOCTA.TA026 = TA026B;
                MOCTA.TA027 = TA027B;
                MOCTA.TA028 = TA028B;
                //MOCTB009 = textBox79.Text;
            }
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                MOCMB001 = MB001D;
                MOCTA004 = BAR3;

                MOCTA.TA026 = TA026C;
                MOCTA.TA027 = TA027C;
                MOCTA.TA028 = TA028C;
                //MOCTB009 = textBox80.Text;
            }
            else if (MANU.Equals("水麵"))
            {
                MOCMB001 = MB001E;
                MOCTA004 = Convert.ToDecimal(textBox31.Text)/ BOMBAR;
                //MOCTB009 = textBox81.Text;
            }

            else if (MANU.Equals("新廠統百包裝線"))
            {
                MOCMB001 = MB001F;
                MOCTA004 = BOX;
                MOCTA.TA026 = TA026D;
                MOCTA.TA027 = TA027D;
                MOCTA.TA028 = TA028D;
            }
            try
            {
                //check TA002=2,TA040=2
                if (MOCTA.TA002.Substring(0,1).Equals("2")&& MOCTA.TA040.Substring(0, 1).Equals("2"))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTA]");
                    sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_count],[DataGroup],[TA001],[TA002],[TA003],[TA004],[TA005],[TA006],[TA007]");
                    sbSql.AppendFormat(" ,[TA009],[TA010],[TA011],[TA012],[TA013],[TA014],[TA015],[TA016],[TA017],[TA018]");
                    sbSql.AppendFormat(" ,[TA019],[TA020],[TA021],[TA022],[TA024],[TA025],[TA029],[TA030],[TA031],[TA034],[TA035]");
                    sbSql.AppendFormat(" ,[TA040],[TA041],[TA043],[TA044],[TA045],[TA046],[TA047],[TA049],[TA050],[TA200]");
                    sbSql.AppendFormat(" ,[TA026],[TA027],[TA028]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" VALUES");
                    sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA.TA003, MOCTA.TA004, MOCTA.TA005, MOCTA.TA006, MOCTA.TA007);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TA009, MOCTA.TA010, MOCTA.TA011, MOCTA.TA012, MOCTA.TA013, MOCTA.TA014, MOCTA.TA015, MOCTA.TA016, MOCTA.TA017, MOCTA.TA018);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}',N'{6}','{7}','{8}','{9}','{10}',", MOCTA.TA019, MOCTA.TA020, MOCTA.TA021, MOCTA.TA022, MOCTA.TA024, MOCTA.TA025, MOCTA.TA029, MOCTA.TA030, MOCTA.TA031, MOCTA.TA034, MOCTA.TA035);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TA040, MOCTA.TA041, MOCTA.TA043, MOCTA.TA044, MOCTA.TA045, MOCTA.TA046, MOCTA.TA047, MOCTA.TA049, MOCTA.TA050, MOCTA.TA200);
                    sbSql.AppendFormat(" '{0}','{1}','{2}'", MOCTA.TA026, MOCTA.TA027, MOCTA.TA028);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" INSERT INTO [TK].dbo.[MOCTB]");
                    sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_count],[DataGroup],[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007]");
                    sbSql.AppendFormat(" ,[TB009],[TB011],[TB012],[TB013],[TB014],[TB018],[TB019],[TB020],[TB022],[TB024]");
                    sbSql.AppendFormat(" ,[TB025],[TB026],[TB027],[TB029],[TB030],[TB031],[TB501],[TB554],[TB556],[TB560])");
                    sbSql.AppendFormat(" (SELECT ");
                    sbSql.AppendFormat(" '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE]", MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE);
                    sbSql.AppendFormat(" ,'{0}' [TRANS_NAME],{1} [sync_count],'{2}' [DataGroup],'{3}' [TB001],'{4}' [TB002],[BOMMD].MD003 [TB003],ROUND({5}*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),3) [TB004],0 [TB005],'****' [TB006],[INVMB].MB004  [TB007]", MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA004);
                    sbSql.AppendFormat(" ,[INVMB].MB017 [TB009],'1' [TB011],[INVMB].MB002 [TB012],[INVMB].MB003 [TB013],[BOMMD].MD001 [TB014],'N' [TB018],'0' [TB019],'0' [TB020],'2' [TB022],'0' [TB024]");
                    sbSql.AppendFormat(" ,'****' [TB025],'0' [TB026],'1' [TB027],'0' [TB029],'0' [TB030],'0' [TB031],'0' [TB501],'N' [TB554],'0' [TB556],'0' [TB560]");
                    sbSql.AppendFormat(" FROM [TK].dbo.[BOMMD],[TK].dbo.[INVMB]");
                    sbSql.AppendFormat(" WHERE [BOMMD].MD003=[INVMB].MB001");
                    sbSql.AppendFormat(" AND MD001='{0}' AND ISNULL(MD012,'')='' )", MOCMB001);
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
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
                

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }



        public MOCTADATA SETMOCTA()
        {
            if (MANU.Equals("新廠製二組"))
            {
                SEARCHBOMMC();

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "140020";
                MOCTA.USR_GROUP = "103000";
                //MOCTA.CREATE_DATE = dt1.ToString("yyyyMMdd");
                MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "140020";
                MOCTA.MODI_DATE = dt1.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt1.ToString("yyyyMMdd");
                MOCTA.TA004 = dt1.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt1.ToString("yyyyMMdd");
                MOCTA.TA010 = dt1.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt1.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                //MOCTA.TA014 = dt1.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BAR * BOMBAR).ToString();
                MOCTA.TA015 = SUM1.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN1;
                MOCTA.TA021 = "02";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002;
                MOCTA.TA035 = MB003;
                MOCTA.TA040 = dt1.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = "0";
                MOCTA.TA046 = "0";
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }

            else if (MANU.Equals("新廠包裝線"))
            {
                SEARCHBOMMC();
                

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "140020";
                MOCTA.USR_GROUP = "103000";
                //MOCTA.CREATE_DATE = dt2.ToString("yyyyMMdd");
                MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "140020";
                MOCTA.MODI_DATE = dt2.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt2.ToString("yyyyMMdd");
                MOCTA.TA004 = dt2.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001B;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt2.ToString("yyyyMMdd");
                MOCTA.TA010 = dt2.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt2.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                // MOCTA.TA014 = dt2.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BOX * BOMBAR).ToString();
                MOCTA.TA015 = SUM2.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN2;
                MOCTA.TA021 = "09";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002B;
                MOCTA.TA035 = MB003B;
                MOCTA.TA040 = dt2.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = "0";
                MOCTA.TA046 = "0";
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }

            else if (MANU.Equals("新廠製一組"))
            {
                SEARCHBOMMC();

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "140020";
                MOCTA.USR_GROUP = "103000";
                //MOCTA.CREATE_DATE = dt3.ToString("yyyyMMdd");
                MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "140020";
                MOCTA.MODI_DATE = dt3.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt3.ToString("yyyyMMdd");
                MOCTA.TA004 = dt3.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001C;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt3.ToString("yyyyMMdd");
                MOCTA.TA010 = dt3.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt3.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                //MOCTA.TA014 = dt3.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BAR2 * BOMBAR).ToString();
                MOCTA.TA015 = SUM3.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN3;
                MOCTA.TA021 = "03";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002C;
                MOCTA.TA035 = MB003C;
                MOCTA.TA040 = dt3.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = "0";
                MOCTA.TA046 = "0";
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                SEARCHBOMMC();

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "140020";
                MOCTA.USR_GROUP = "103000";
                //MOCTA.CREATE_DATE = dt4.ToString("yyyyMMdd");
                MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "140020";
                MOCTA.MODI_DATE = dt4.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt4.ToString("yyyyMMdd");
                MOCTA.TA004 = dt4.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001D;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt4.ToString("yyyyMMdd");
                MOCTA.TA010 = dt4.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt4.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                //MOCTA.TA014 = dt4.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BAR3 * BOMBAR).ToString();
                MOCTA.TA015 = SUM4.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN4;
                MOCTA.TA021 = "04";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002D;
                MOCTA.TA035 = MB003D;
                MOCTA.TA040 = dt4.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = "0";
                MOCTA.TA046 = "0";
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }
            else if (MANU.Equals("水麵"))
            {
                SEARCHBOMMC();

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "140020";
                MOCTA.USR_GROUP = "103000";
                //MOCTA.CREATE_DATE = dt5.ToString("yyyyMMdd");
                MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "140020";
                MOCTA.MODI_DATE = dt5.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt5.ToString("yyyyMMdd");
                MOCTA.TA004 = dt5.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001E;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt5.ToString("yyyyMMdd");
                MOCTA.TA010 = dt5.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt5.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                //MOCTA.TA014 = dt5.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                MOCTA.TA015 = textBox31.Text;
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = textBox36.Text;
                MOCTA.TA021 = textBox27.Text;
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = "";
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002E;
                MOCTA.TA035 = MB003E;
                MOCTA.TA040 = dt5.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = "0";
                MOCTA.TA046 = "0";
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";



                return MOCTA;
            }

            else if (MANU.Equals("新廠統百包裝線"))
            {
                SEARCHBOMMC();

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "140020";
                MOCTA.USR_GROUP = "103000";
                //MOCTA.CREATE_DATE = dt6.ToString("yyyyMMdd");
                MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "140020";
                MOCTA.MODI_DATE = dt6.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt6.ToString("yyyyMMdd");
                MOCTA.TA004 = dt6.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001F;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt6.ToString("yyyyMMdd");
                MOCTA.TA010 = dt6.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt6.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                // MOCTA.TA014 = dt2.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BOX * BOMBAR).ToString();
                MOCTA.TA015 = SUM5.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN6;
                MOCTA.TA021 = "10";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002F;
                MOCTA.TA035 = MB003F;
                MOCTA.TA040 = dt6.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = "0";
                MOCTA.TA046 = "0";
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }


            return null;
            
        }
        

        public void SEARCHBOMMC()
        {
            BOMVARSION = null;
            UNIT = null;
            BOMBAR = 0;

            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001B);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001C);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001D);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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
            else if (MANU.Equals("水麵"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001E);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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

            else if (MANU.Equals("新廠統百包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001F);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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

        }
        public void SEARCHMOCMANULINERESULT()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(@"  WHERE [SID]='{0}'", ID1);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                    sqlConn.Open();
                    ds3.Clear();
                    adapter3.Fill(ds3, "TEMPds3");
                    sqlConn.Close();


                    if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                    {
                        dataGridView2.DataSource = null;
                    }
                    else
                    {
                        if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                        {

                            dataGridView2.DataSource = ds3.Tables["TEMPds3"];
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
            else if  (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(@"  WHERE [SID]='{0}'", ID2);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter8 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder8 = new SqlCommandBuilder(adapter8);
                    sqlConn.Open();
                    ds6.Clear();
                    adapter8.Fill(ds6, "TEMPds6");
                    sqlConn.Close();


                    if (ds6.Tables["TEMPds6"].Rows.Count == 0)
                    {
                        dataGridView4.DataSource = null;
                    }
                    else
                    {
                        if (ds6.Tables["TEMPds6"].Rows.Count >= 1)
                        {

                            dataGridView4.DataSource = ds6.Tables["TEMPds6"];
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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(@"  WHERE [SID]='{0}'", ID3);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter11 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder11 = new SqlCommandBuilder(adapter11);
                    sqlConn.Open();
                    ds9.Clear();
                    adapter11.Fill(ds9, "TEMPds9");
                    sqlConn.Close();


                    if (ds9.Tables["TEMPds9"].Rows.Count == 0)
                    {
                        dataGridView6.DataSource = null;

                    }
                    else
                    {
                        if (ds9.Tables["TEMPds9"].Rows.Count >= 1)
                        {

                            dataGridView6.DataSource = ds9.Tables["TEMPds9"];
                            dataGridView6.AutoResizeColumns();
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(@"  WHERE [SID]='{0}'", ID4);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter12 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder12 = new SqlCommandBuilder(adapter12);
                    sqlConn.Open();
                    ds10.Clear();
                    adapter12.Fill(ds10, "TEMPds10");
                    sqlConn.Close();


                    if (ds10.Tables["TEMPds10"].Rows.Count == 0)
                    {
                        dataGridView8.DataSource = null;
                    }
                    else
                    {
                        if (ds10.Tables["TEMPds10"].Rows.Count >= 1)
                        {

                            dataGridView8.DataSource = ds10.Tables["TEMPds10"];
                            dataGridView8.AutoResizeColumns();
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

            else if (MANU.Equals("新廠統百包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(@"  WHERE [SID]='{0}'", ID6);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter25 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder25 = new SqlCommandBuilder(adapter25);
                    sqlConn.Open();
                    ds25.Clear();
                    adapter25.Fill(ds25, "TEMPds25");
                    sqlConn.Close();


                    if (ds25.Tables["TEMPds25"].Rows.Count == 0)
                    {
                        dataGridView17.DataSource = null;
                    }
                    else
                    {
                        if (ds25.Tables["TEMPds25"].Rows.Count >= 1)
                        {

                            dataGridView17.DataSource = ds25.Tables["TEMPds25"];
                            dataGridView17.AutoResizeColumns();
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

        }

        public string GETMAXTA002(string TA001)
        {
            string TA002;

            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt1.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

                        }
                        return null;
                    }

                }
                catch
                {
                    return null;
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt2.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

                        }
                        return null;
                    }

                }
                catch
                {
                    return null;
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt3.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

                        }
                        return null;
                    }

                }
                catch
                {
                    return null;
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt4.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

                        }
                        return null;
                    }

                }
                catch
                {
                    return null;
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            else if (MANU.Equals("水麵"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt5.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

                        }
                        return null;
                    }

                }
                catch
                {
                    return null;
                }
                finally
                {
                    sqlConn.Close();
                }
            }

            else if (MANU.Equals("新廠統百包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt6.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

                        }
                        return null;
                    }

                }
                catch
                {
                    return null;
                }
                finally
                {
                    sqlConn.Close();
                }
            }

            return null;

        }
        public string SETTA002(string TA002)
        {

            if (MANU.Equals("新廠製二組"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt1.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt1.ToString("yyyyMMdd") + temp.ToString();
                }
            }

            else if (MANU.Equals("新廠包裝線"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt2.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt2.ToString("yyyyMMdd") + temp.ToString();
                }
            }

            else if (MANU.Equals("新廠製一組"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt3.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt3.ToString("yyyyMMdd") + temp.ToString();
                }
            }
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt4.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt4.ToString("yyyyMMdd") + temp.ToString();
                }
            }
            else if (MANU.Equals("水麵"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt5.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt5.ToString("yyyyMMdd") + temp.ToString();
                }
            }

            else if (MANU.Equals("新廠統百包裝線"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt6.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt6.ToString("yyyyMMdd") + temp.ToString();
                }
            }

            return null;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                //MessageBox.Show("新廠製二組");
                MANU = "新廠製二組";
            }
            else if(tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                //MessageBox.Show("新廠製一組");
                MANU = "新廠製一組";
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                //MessageBox.Show("新廠製三組(手工)");
                MANU = "新廠製三組(手工)";
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                //MessageBox.Show("新廠包裝線");
                MANU = "新廠包裝線";
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
            {
                //MessageBox.Show("水麵");
                MANU = "水麵";
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage8"])
            {
                //MessageBox.Show("水麵");
                MANU = "新廠統百包裝線";
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage8"])
            {
                //MessageBox.Show("水麵");
                MANU = "新廠統百包裝線";
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage10"])
            {
                //MessageBox.Show("水麵");
                MANU = "少量訂單";
            }

            


        }
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();

            SEARCHMOCMANULINETEMPDATAS(textBox7.Text.Trim());


        }
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    ID2 = row.Cells["ID"].Value.ToString();
                    dt2 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001B = row.Cells["品號"].Value.ToString();
                    MB002B = row.Cells["品名"].Value.ToString();
                    MB003B = row.Cells["規格"].Value.ToString();
                    BOX = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    SUM2 = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    TA029 = row.Cells["備註"].Value.ToString();
                    TA026 = row.Cells["訂單單別"].Value.ToString();
                    TA027 = row.Cells["訂單號"].Value.ToString();
                    TA028 = row.Cells["訂單序號"].Value.ToString();
                                        
                    SUBID2 = row.Cells["ID"].Value.ToString();
                    SUBBAR2 = "";
                    SUBNUM2 = "";
                    SUBBOX2 = row.Cells["箱數"].Value.ToString();
                    SUBPACKAGE2 = row.Cells["包裝數"].Value.ToString();

                    SEARCHMOCMANULINERESULT();
                    SEARCHMOCMANULINEMERGERESLUTMOCTA(ID2.ToString());
                    //SEARCHMOCMANULINECOP();

                }
                else
                {
                    ID2 = null;
                    SUBID2 = null;
                    SUBBAR2 = null;
                    SUBNUM2 = null;
                    SUBBOX2= null;
                    SUBPACKAGE2 = null;
                    TA026 = null;
                    TA027 = null;
                    TA028 = null;

                }
            }
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
            SEARCHBOMMD();

            SEARCHMOCMANULINETEMPDATAS(textBox14.Text.Trim());
        }
        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            textBox44.Text = null;
            textBox45.Text = null;

            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    ID3 = row.Cells["ID"].Value.ToString();
                    dt3 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001C= row.Cells["品號"].Value.ToString();
                    MB002C = row.Cells["品名"].Value.ToString();
                    MB003C = row.Cells["規格"].Value.ToString();
                    BAR2 = Convert.ToDecimal(row.Cells["桶數"].Value.ToString());
                    SUM3 = Convert.ToDecimal(row.Cells["數量"].Value.ToString());
                    TA029 = row.Cells["備註"].Value.ToString();
                    TA026B = row.Cells["訂單單別"].Value.ToString();
                    TA027B = row.Cells["訂單號"].Value.ToString();
                    TA028B = row.Cells["訂單序號"].Value.ToString();

                    SUBID3 = row.Cells["ID"].Value.ToString();
                    SUBBAR3 = row.Cells["桶數"].Value.ToString();
                    SUBNUM3 = row.Cells["數量"].Value.ToString();
                    SUBBOX3 = null;
                    SUBPACKAGE3 = null;

                    SEARCHMOCMANULINERESULT();
                    SEARCHMOCMANULINEMERGERESLUTMOCTA(ID3.ToString());
                    //SEARCHMOCMANULINECOP();

                }
                else
                {
                    ID3 = null;
                    SUBID3 = null;
                    SUBBAR3 = null;
                    SUBNUM3 = null;
                    SUBBOX3 = null;
                    SUBPACKAGE3 = null;

                    TA026B = null;
                    TA027B = null;
                    TA028B = null;

                }
            }
        }
        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
            SEARCHBOMMD();

            SEARCHMOCMANULINETEMPDATAS(textBox20.Text.Trim());
        }
        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {

            textBox46.Text = null;
            textBox47.Text = null;

            if (dataGridView7.CurrentRow != null)
            {
                int rowindex = dataGridView7.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView7.Rows[rowindex];
                    ID4 = row.Cells["ID"].Value.ToString();
                    dt4 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001D = row.Cells["品號"].Value.ToString();
                    MB002D = row.Cells["品名"].Value.ToString();
                    MB003D = row.Cells["規格"].Value.ToString();
                    BAR3 = Convert.ToDecimal(row.Cells["桶數"].Value.ToString());
                    SUM4 = Convert.ToDecimal(row.Cells["數量"].Value.ToString());
                    TA029 = row.Cells["備註"].Value.ToString();

                    TA026C = row.Cells["訂單單別"].Value.ToString();
                    TA027C = row.Cells["訂單號"].Value.ToString();
                    TA028C = row.Cells["訂單序號"].Value.ToString();

                    SUBID4 = row.Cells["ID"].Value.ToString();
                    SUBBAR4 = row.Cells["桶數"].Value.ToString();
                    SUBNUM4 = row.Cells["數量"].Value.ToString();
                    SUBBOX4 = null;
                    SUBPACKAGE4 = null;

                    SEARCHMOCMANULINERESULT();
                    SEARCHMOCMANULINEMERGERESLUTMOCTA(ID4.ToString());
                    //SEARCHMOCMANULINECOP();

                }
                else
                {
                    ID4 = null;
                    SUBID4 = null;
                    SUBBAR4= null;
                    SUBNUM4 = null;
                    SUBBOX4 = null;
                    SUBPACKAGE4 = null;

                    TA026C = null;
                    TA027C = null;
                    TA028C = null;

                }
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
                    DELID1 = row.Cells["SID"].Value.ToString();
                    DELMOCTA001A= row.Cells["製令"].Value.ToString();
                    DELMOCTA002A = row.Cells["單號"].Value.ToString();

                }
                else
                {
                    DELID1 = null;

                }
            }
        }
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    DELID2 = row.Cells["SID"].Value.ToString();
                    DELMOCTA001B = row.Cells["製令"].Value.ToString();
                    DELMOCTA002B = row.Cells["單號"].Value.ToString();



                }
                else
                {
                    DELID2 = null;

                }
            }
        }

        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];
                    DELID3 = row.Cells["SID"].Value.ToString();
                    DELMOCTA001C = row.Cells["製令"].Value.ToString();
                    DELMOCTA002C = row.Cells["單號"].Value.ToString();



                }
                else
                {
                    DELID3 = null;

                }
            }
        }

        private void dataGridView8_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView8.CurrentRow != null)
            {
                int rowindex = dataGridView8.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView8.Rows[rowindex];
                    DELID4 = row.Cells["SID"].Value.ToString();
                    DELMOCTA001D = row.Cells["製令"].Value.ToString();
                    DELMOCTA002D = row.Cells["單號"].Value.ToString();



                }
                else
                {
                    DELID4 = null;

                }
            }
        }

        public void DELMOCMANULINERESULT()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat("  WHERE SID='{0}' ", DELID1);
                    sbSql.AppendFormat("  AND [MOCTA001] ='{0}' AND [MOCTA002]='{1}'",DELMOCTA001A, DELMOCTA002A);
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

            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat("  WHERE SID='{0}'", DELID2);
                    sbSql.AppendFormat("  AND [MOCTA001] ='{0}' AND [MOCTA002]='{1}'", DELMOCTA001B, DELMOCTA002B);
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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat("  WHERE SID='{0}'", DELID3);
                    sbSql.AppendFormat("  AND [MOCTA001] ='{0}' AND [MOCTA002]='{1}'", DELMOCTA001C, DELMOCTA002C);
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat("  WHERE SID='{0}'", DELID4);
                    sbSql.AppendFormat("  AND [MOCTA001] ='{0}' AND [MOCTA002]='{1}'", DELMOCTA001D, DELMOCTA002D);
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

            else if (MANU.Equals("新廠統百包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat("  WHERE SID='{0}'", DELID6);
                    sbSql.AppendFormat("  AND [MOCTA001] ='{0}' AND [MOCTA002]='{1}'", DELMOCTA001F, DELMOCTA002F);
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

        }

        public void SEARCHMOCTB()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA003  AS '日期',[TA021] AS '線別號',[MD002] AS '線別',TB003 AS '品號',TB012 AS '品名',SUM(TB004)  AS '總數量',TB009  AS '入庫別'");
                sbSql.AppendFormat(@"  ,(SELECT  TOP 1 [MOCTA001]+' '+[MOCTA002] FROM [TKMOC].[dbo].[MOCMANULINETOATL] WHERE [TA003]=MOCTA.TA003 AND [TA021]=MOCTA.[TA021] AND [TB003]=MOCTB.TB003  AND [TB004]=SUM(MOCTB.TB004) ORDER BY [MOCTA001]+[MOCTA002] DESC) AS '製令' ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTB, [TK].dbo.MOCTA,[TK].dbo.CMSMD");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND [TA021]=MD001");
                sbSql.AppendFormat(@"  AND TB012 LIKE '%水麵%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'",dateTimePicker10.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  GROUP BY TB003,TB012,TB009,TA003,[TA021],[MD002] ");
                sbSql.AppendFormat(@"  ORDER BY TA003,[TA021],TB003");
                sbSql.AppendFormat(@"  ");

                adapter13 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder13 = new SqlCommandBuilder(adapter13);
                sqlConn.Open();
                ds13.Clear();
                adapter13.Fill(ds13, "TEMPds13");
                sqlConn.Close();


                if (ds13.Tables["TEMPds13"].Rows.Count == 0)
                {
                    dataGridView9.DataSource = null;
                }
                else
                {
                    if (ds13.Tables["TEMPds13"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView9.DataSource = ds13.Tables["TEMPds13"];
                        dataGridView9.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        private void dataGridView9_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView9.CurrentRow != null)
            {
                int rowindex = dataGridView9.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView9.Rows[rowindex];
                    textBox26.Text = row.Cells["日期"].Value.ToString();
                    textBox27.Text = row.Cells["線別號"].Value.ToString();
                    textBox28.Text = row.Cells["線別"].Value.ToString();
                    textBox29.Text = row.Cells["品號"].Value.ToString();
                    textBox30.Text = row.Cells["品名"].Value.ToString();
                    textBox31.Text = row.Cells["總數量"].Value.ToString();
                    textBox36.Text = row.Cells["入庫別"].Value.ToString();
                    dt5 = Convert.ToDateTime(row.Cells["日期"].Value.ToString().Substring(0,4)+"/"+row.Cells["日期"].Value.ToString().Substring(4, 2)+"/"+ row.Cells["日期"].Value.ToString().Substring(6, 2));

                    MB001E = row.Cells["品號"].Value.ToString();
                    MB002E = row.Cells["品名"].Value.ToString();                   

                    SEARCHMOCMANULINETOATL();
                }
                else
                {
                    textBox26.Text = null;
                    textBox27.Text = null;
                    textBox28.Text = null;
                    textBox29.Text = null;
                    textBox30.Text = null;
                    textBox31.Text = null;

                }
            }
        }

        public void SEARCHMOCMANULINETOATL()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '單別',[MOCTA002] AS '製令'");
                sbSql.AppendFormat(@"  ,[TA003] AS '日期',[TA021] AS '線別號',[TA021N] AS '線別',[TB003] AS '品號',[TB012] AS '品名',[TB004] AS '總數量',[TB009] AS '入庫別'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINETOATL]");
                sbSql.AppendFormat(@"  WHERE [TA003]='{0}' AND [TA021]='{1}' AND [TB003]='{2}'   AND [TB004]='{3}' ", textBox26.Text, textBox27.Text, textBox29.Text, textBox31.Text);
                sbSql.AppendFormat(@"  ");

                adapter14 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder14 = new SqlCommandBuilder(adapter14);
                sqlConn.Open();
                ds14.Clear();
                adapter14.Fill(ds14, "TEMPds14");
                sqlConn.Close();


                if (ds14.Tables["TEMPds14"].Rows.Count == 0)
                {
                    dataGridView10.DataSource = null;
                }
                else
                {
                    if (ds14.Tables["TEMPds14"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView10.DataSource = ds14.Tables["TEMPds14"];
                        dataGridView10.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            //CALPRODUCT();
        }

        public void CALPRODUCT()
        {
            try
            {
                if (MANU.Equals("新廠製二組"))
                {
                    textBox5.Text = (Convert.ToDecimal(textBox32.Text) * Convert.ToDecimal(textBox4.Text)).ToString();
                }

                else if (MANU.Equals("新廠包裝線"))
                {
                    
                    textBox12.Text = (Convert.ToDecimal(textBox33.Text) * Convert.ToDecimal(textBox8.Text) ).ToString();
                }
                else if (MANU.Equals("新廠製一組"))
                {
                    textBox19.Text = (Convert.ToDecimal(textBox34.Text) * Convert.ToDecimal(textBox15.Text)).ToString();
                }
                else if (MANU.Equals("新廠製三組(手工)"))
                {
                    textBox23.Text = (Convert.ToDecimal(textBox35.Text) * Convert.ToDecimal(textBox21.Text)).ToString();
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

        public void CALPRODUCTDETAIL()
        {
            try
            {
                if (MANU.Equals("新廠製二組"))
                {
                    textBox4.Text = Math.Round(Convert.ToDecimal(textBox5.Text)/ Convert.ToDecimal(textBox32.Text), 4).ToString();
                }

                else if (MANU.Equals("新廠包裝線"))
                {
                    SEARCHMB001BOX();
                    textBox8.Text = Math.Round(Convert.ToDecimal(textBox12.Text) / Convert.ToDecimal(textBox33.Text)/BOXNUMERB, 4).ToString();
                }
                else if (MANU.Equals("新廠製一組"))
                {
                    textBox15.Text = Math.Round(Convert.ToDecimal(textBox19.Text) / Convert.ToDecimal(textBox34.Text), 4).ToString();
                }
                else if (MANU.Equals("新廠製三組(手工)"))
                {
                    textBox21.Text = Math.Round(Convert.ToDecimal(textBox23.Text) / Convert.ToDecimal(textBox35.Text), 4).ToString();
                }
                else if (MANU.Equals("新廠統百包裝線"))
                {
                    SEARCHMB001BOX();
                    textBox59.Text = Math.Round(Convert.ToDecimal(textBox61.Text) / Convert.ToDecimal(textBox60.Text) / BOXNUMERB, 4).ToString();
                }
                else if (MANU.Equals("少量訂單"))
                {
                    SEARCHMB001BOX();
                    textBox753.Text = Math.Round(Convert.ToDecimal(textBox751.Text) / Convert.ToDecimal(textBox752.Text) / BOXNUMERB, 4).ToString();
                    textBox741.Text = Math.Round(Convert.ToDecimal(textBox742.Text) / Convert.ToDecimal(textBox752.Text), 4).ToString();
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
        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            //CALPRODUCT();
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            //CALPRODUCT();
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            //CALPRODUCT();
        }

        public void ADDMOCMANULINETOATL()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINETOATL]");
                sbSql.AppendFormat(" ([ID],[TA003],[TA021],[TA021N],[TB003],[TB012],[TB004],[TB009],[MOCTA001],[MOCTA002])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')", "NEWID()",textBox26.Text,textBox27.Text,textBox28.Text,textBox29.Text, textBox30.Text, textBox31.Text,textBox36.Text, TA001, TA002);
                sbSql.AppendFormat(" ");
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

        public void DELMOCMANULINETOATL()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINETOATL]");
                sbSql.AppendFormat("  WHERE ID='{0}'", DELID5);          
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

        private void dataGridView10_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView10.CurrentRow != null)
            {
                int rowindex = dataGridView10.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView10.Rows[rowindex];
                    DELID5 = row.Cells["ID"].Value.ToString();
                    DELMOCTA001E = row.Cells["單別"].Value.ToString();
                    DELMOCTA002E = row.Cells["製令"].Value.ToString();



                }
                else
                {
                    DELID5 = null;

                }
            }
        }

        public void SETIN()
        {
            label51.Text = "20001";
            label52.Text = "20001";
            label53.Text = "20001";
            label54.Text = "20001";
            IN1 = "20001";
            IN2 = "20001";
            IN3 = "20001";
            IN4 = "20001";
            label104.Text = "20001";

        }
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            label51.Text = comboBox5.SelectedValue.ToString();
            IN1= comboBox5.SelectedValue.ToString();
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            label52.Text = comboBox6.SelectedValue.ToString();
            IN2 = comboBox6.SelectedValue.ToString();
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            label53.Text = comboBox7.SelectedValue.ToString();
            IN3 = comboBox7.SelectedValue.ToString();

        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            label54.Text = comboBox8.SelectedValue.ToString();
            IN4 = comboBox8.SelectedValue.ToString();

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }
        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }
        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }

        private void textBox23_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }

        public void SEARCHBOMMD()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT MD001,MD003,MB002,CONVERT(decimal(18,2), MD006/MD007) AS MD006");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MD003=MB001");
                    sbSql.AppendFormat(@"  AND MB002 LIKE '%低筋%'");
                    sbSql.AppendFormat(@"  AND MD001='{0}'", textBox1.Text);
                    sbSql.AppendFormat(@"  ");


                    adapter15 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder15 = new SqlCommandBuilder(adapter15);
                    sqlConn.Open();
                    ds15.Clear();
                    adapter15.Fill(ds15, "TEMPds15");
                    sqlConn.Close();


                    if (ds15.Tables["TEMPds15"].Rows.Count == 0)
                    {
                        SETNULL5();
                    }
                    else
                    {
                        if (ds15.Tables["TEMPds15"].Rows.Count >= 1)
                        {
                            textBox37.Text = ds15.Tables["TEMPds15"].Rows[0]["MD006"].ToString();
                         ;
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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                   

                }
                catch
                {

                }
                finally
                {

                }


            }

            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MD001,MD003,MB002,CONVERT(decimal(18,2), MD006/MD007) AS MD006");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MD003=MB001");
                    sbSql.AppendFormat(@"  AND MB002 LIKE '%低筋%'");
                    sbSql.AppendFormat(@"  AND MD001='{0}'", textBox14.Text);
                    sbSql.AppendFormat(@"  ");

                    adapter16 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder16 = new SqlCommandBuilder(adapter16);
                    sqlConn.Open();
                    ds16.Clear();
                    adapter16.Fill(ds16, "TEMPds16");
                    sqlConn.Close();


                    if (ds16.Tables["TEMPds16"].Rows.Count == 0)
                    {
                        SETNULL5(); 
                    }
                    else
                    {
                        if (ds16.Tables["TEMPds16"].Rows.Count >= 1)
                        {
                            textBox38.Text = ds16.Tables["TEMPds16"].Rows[0]["MD006"].ToString();
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MD001,MD003,MB002,CONVERT(decimal(18,2), MD006/MD007) AS MD006");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MD003=MB001");
                    sbSql.AppendFormat(@"  AND MB002 LIKE '%低筋%'");
                    sbSql.AppendFormat(@"  AND MD001='{0}'", textBox20.Text);
                    sbSql.AppendFormat(@"  ");

                    adapter17 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder17 = new SqlCommandBuilder(adapter17);
                    sqlConn.Open();
                    ds17.Clear();
                    adapter17.Fill(ds17, "TEMPds17");
                    sqlConn.Close();


                    if (ds17.Tables["TEMPds17"].Rows.Count == 0)
                    {
                        SETNULL5();
                    }
                    else
                    {
                        if (ds17.Tables["TEMPds17"].Rows.Count >= 1)
                        {
                            textBox39.Text = ds17.Tables["TEMPds17"].Rows[0]["MD006"].ToString();
                            
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
        }

        public void SETNULL5()
        {
            //textBox1.Text = null;

            textBox37.Text = null;
            textBox38.Text = null;
            textBox39.Text = null;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker1.Value;
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker4.Value = dateTimePicker3.Value;
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker7.Value = dateTimePicker6.Value;
        }

        private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker9.Value = dateTimePicker8.Value;
        }

        public void SEARCHMB017()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004,MB017            ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL1();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            comboBox5.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            label51.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL1();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            comboBox6.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            label52.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();

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

            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL4();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            comboBox7.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            label53.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL4();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            comboBox8.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            label54.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
        }

        public void UPDATEMOCMANULINE()
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                frmMOCMANULINESub MOCMANULINESub = new frmMOCMANULINESub(ID1);
                MOCMANULINESub.ShowDialog();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                frmMOCMANULINESub MOCMANULINESub = new frmMOCMANULINESub(ID2);
                MOCMANULINESub.ShowDialog();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                frmMOCMANULINESub MOCMANULINESub = new frmMOCMANULINESub(ID3);
                MOCMANULINESub.ShowDialog();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                frmMOCMANULINESub MOCMANULINESub = new frmMOCMANULINESub(ID4);
                MOCMANULINESub.ShowDialog();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
            {
                
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage8"])
            {
                frmMOCMANULINESub MOCMANULINESub = new frmMOCMANULINESub(ID6);
                MOCMANULINESub.ShowDialog();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage10"])
            {
                frmMOCMANULINESubTEMP frmMOCMANULINESubTEMP = new frmMOCMANULINESubTEMP(ID10);
                frmMOCMANULINESubTEMP.ShowDialog();
            }



        }

        public void CHECKMOCTAB()
        {
            string CHECKID = null;

            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                CHECKID = ID1;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                CHECKID = ID2;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                CHECKID = ID3;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                CHECKID = ID4;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
            {

            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage8"])
            {
                CHECKID = ID6;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage10"])
            {
                CHECKID = ID10;
            }

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT	MOCTA001,MOCTA002");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                sbSql.AppendFormat(@"  WHERE [SID]='{0}'",CHECKID);
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT	TA001,TA002");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA]");
                sbSql.AppendFormat(@"  WHERE EXISTS (SELECT [MOCTA001],[MOCTA002] FROM [TKMOC].[dbo].[MOCMANULINERESULT] WHERE [SID]='{0}' AND TA001=MOCTA001 AND TA002=MOCTA002)", CHECKID);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter19 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder19 = new SqlCommandBuilder(adapter19);
                sqlConn.Open();
                ds19.Clear();
                adapter19.Fill(ds19, "TEMPds19");
                sqlConn.Close();


                if (ds19.Tables["TEMPds19"].Rows.Count == 0)
                {
                    UPDATEMOCMANULINE();
                }
                else
                {
                    MessageBox.Show("有製令未刪除，請檢查一下");
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SEARCHMB001BOX()
        {
            
            if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT TOP 1 MD001,MD003,MB001,MB002,ISNULL(MD007,1) AS MD007,ISNULL(MD010,1) AS MD010");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MD003=MB001");
                    sbSql.AppendFormat(@"  AND MB002 LIKE '%箱%'");
                    sbSql.AppendFormat(@"  AND MD003 LIKE '2%'");
                    sbSql.AppendFormat(@"  AND MD001='{0}'", textBox7.Text);
                    sbSql.AppendFormat(@"  ");

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
                            BOXNUMERB = (Convert.ToInt32(ds20.Tables["TEMPds20"].Rows[0]["MD007"].ToString())/ Convert.ToInt32(ds20.Tables["TEMPds20"].Rows[0]["MD010"].ToString()));
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

            else if (MANU.Equals("新廠統百包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT TOP 1 MD001,MD003,MB001,MB002,ISNULL(MD007,1) AS MD007,ISNULL(MD010,1) AS MD010");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MD003=MB001");
                    sbSql.AppendFormat(@"  AND MB002 LIKE '%箱%'");
                    sbSql.AppendFormat(@"  AND MD003 LIKE '2%'");
                    sbSql.AppendFormat(@"  AND MD001='{0}'", textBox56.Text);
                    sbSql.AppendFormat(@"  ");

                    adapter24 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder24 = new SqlCommandBuilder(adapter24);
                    sqlConn.Open();
                    ds24.Clear();
                    adapter24.Fill(ds24, "TEMPds24");
                    sqlConn.Close();


                    if (ds24.Tables["TEMPds24"].Rows.Count == 0)
                    {
                        BOXNUMERB = 1;
                    }
                    else
                    {
                        if (ds24.Tables["TEMPds24"].Rows.Count >= 1)
                        {
                            BOXNUMERB = (Convert.ToInt32(ds24.Tables["TEMPds24"].Rows[0]["MD007"].ToString()) / Convert.ToInt32(ds24.Tables["TEMPds24"].Rows[0]["MD010"].ToString()));
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
            else if (MANU.Equals("少量訂單"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT TOP 1 MD001,MD003,MB001,MB002,ISNULL(MD007,1) AS MD007,ISNULL(MD010,1) AS MD010");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MD003=MB001");
                    sbSql.AppendFormat(@"  AND MB002 LIKE '%箱%'");
                    sbSql.AppendFormat(@"  AND MD003 LIKE '2%'");
                    sbSql.AppendFormat(@"  AND MD001='{0}'", textBox7.Text);
                    sbSql.AppendFormat(@"  ");

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
                            BOXNUMERB = (Convert.ToInt32(ds20.Tables["TEMPds20"].Rows[0]["MD007"].ToString()) / Convert.ToInt32(ds20.Tables["TEMPds20"].Rows[0]["MD010"].ToString()));
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

        }
       


      

        private void dataGridView12_SelectionChanged(object sender, EventArgs e)
        {
           
        }

   

    

        public void DELMOCMANULINECOP()
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                DELMOCMANULINECOP2(ID1);
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                DELMOCMANULINECOP2(ID2);
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                DELMOCMANULINECOP2(ID3);
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                DELMOCMANULINECOP2(ID4);
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
            {
               
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage8"])
            {
                DELMOCMANULINECOP2(ID6);
            }
        }

        public void DELMOCMANULINECOP2(string SID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINECOP]");
                sbSql.AppendFormat("  WHERE SID='{0}'", SID);
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

        private void textBox7_Leave(object sender, EventArgs e)
        {
            SEARCHMB001();
        }


       
        public void SETCALENDAR()
        {
            string EVENT;
            DateTime dtEVENT;
            var ce2 = new CustomEvent();


            calendar1.RemoveAllEvents();
            calendar1.CalendarDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
            calendar1.CalendarView = CalendarViews.Month;
            calendar1.AllowEditingEvents = true;

            

         
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                

                sbSql.AppendFormat(@"  SELECT [EVENTDATE],[MOCLINE],[EVENT]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[CALENDAR]");
                sbSql.AppendFormat(@"  WHERE [EVENTDATE]>='{0}'", DateTime.Now.ToString("yyyy")+"0101");
                sbSql.AppendFormat(@"  ORDER BY [EVENTDATE]");
                sbSql.AppendFormat(@"  ");

                adapterCALENDAR = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderCALENDAR = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                dsCALENDAR.Clear();
                adapterCALENDAR.Fill(dsCALENDAR, "TEMPdsCALENDAR");
                sqlConn.Close();


                if (dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows.Count >= 1)
                    {
                        foreach (DataRow od in dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows)
                        {
                            EVENT = od["MOCLINE"].ToString()+"-"+od["EVENT"].ToString();
                            dtEVENT = Convert.ToDateTime(od["EVENTDATE"].ToString());

                            ce2 = new CustomEvent
                            {
                                IgnoreTimeComponent = false,
                                EventText = EVENT,
                                Date = new DateTime(dtEVENT.Year, dtEVENT.Month, dtEVENT.Day),
                                EventLengthInHours = 2f,
                                RecurringFrequency = RecurringFrequencies.None,
                                EventFont = new Font("Verdana", 12, FontStyle.Regular),
                                Enabled = true,
                                EventColor = Color.FromArgb(120, 255, 120),
                                EventTextColor = Color.Black,
                                ThisDayForwardOnly = true
                            };
                            
                            calendar1.AddEvent(ce2);
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

        public void ADDCALENDAR()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[CALENDAR]");
                sbSql.AppendFormat(" ([EVENTDATE],[MOCLINE],[EVENT])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')",  dateTimePicker11.Value.ToString("yyyy/MM/dd"), comboBox10.Text , comboBox9.Text + "-" + textBox48.Text);
                sbSql.AppendFormat(" ");
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

        public void DELCALENDAR()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[CALENDAR]");
                sbSql.AppendFormat(" WHERE convert(varchar, [EVENTDATE], 112)='{0}' AND [MOCLINE]='{1}'", dateTimePicker11.Value.ToString("yyyyMMdd"), comboBox10.Text);
                sbSql.AppendFormat(" ");
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

        public void SEARCHCOPTD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if(comboBox11.Text.Equals("未完成"))
                {
                    sbSqlQuery.AppendFormat(@" AND TD008-TD113>0 ");
                }
                else if (comboBox11.Text.Equals("已完成"))
                {
                    sbSqlQuery.AppendFormat(@" AND TD008-TD113=0 ");
                }
                else if (comboBox11.Text.Equals("全部"))
                {
                    sbSqlQuery.AppendFormat(@"  ");
                }


                //TD009不可用，改用TD113記錄已生產數量
                sbSql.AppendFormat(@"  SELECT TD013 AS '預交日',TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD008 AS '訂單數',TD113 AS '已生產數量',TD010 AS '單位',TC053 AS '客戶'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTD,[TK].dbo.COPTC");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD001='{0}'",comboBox12.SelectedValue.ToString());
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'",dateTimePicker12.Value.ToString("yyyyMMdd"), dateTimePicker13.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TD004 LIKE '401%'");
                sbSql.AppendFormat(@"  {0}", sbSqlQuery.ToString());
                sbSql.AppendFormat(@"  ORDER BY TD013,TD004,TD001,TD002");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter22 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder22 = new SqlCommandBuilder(adapter22);
                sqlConn.Open();
                ds22.Clear();
                adapter22.Fill(ds22, "TEMPds22");
                sqlConn.Close();


                if (ds22.Tables["TEMPds22"].Rows.Count == 0)
                {
                    dataGridView15.DataSource = null;
                }
                else
                {
                    if (ds22.Tables["TEMPds22"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView15.DataSource = ds22.Tables["TEMPds22"];
                        dataGridView15.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        public void UPDATECOPTD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" UPDATE [TK].dbo.COPTD");
                sbSql.AppendFormat(" SET TD113='{0}'",numericUpDown1.Value.ToString());
                sbSql.AppendFormat(" WHERE TD001='{0}' AND TD002='{1}' AND TD003='{2}'",textBox49.Text, textBox50.Text, textBox51.Text);
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

        private void dataGridView15_SelectionChanged(object sender, EventArgs e)
        {
            textBox49.Text = null;
            textBox50.Text = null;
            textBox51.Text = null;

            if (dataGridView15.CurrentRow != null)
            {
                int rowindex = dataGridView15.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView15.Rows[rowindex];
                    textBox49.Text = row.Cells["單別"].Value.ToString();
                    textBox50.Text = row.Cells["單號"].Value.ToString();
                    textBox51.Text = row.Cells["序號"].Value.ToString();
                }
                else
                {
                    textBox49.Text = null;
                    textBox50.Text = null;
                    textBox51.Text = null;
                }
            }
        }


        private void textBox56_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
        }

        private void textBox61_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }
        private void dataGridView16_SelectionChanged(object sender, EventArgs e)
        {
            textBox66.Text = null;
            textBox65.Text = null;

            if (dataGridView16.CurrentRow != null)
            {
                int rowindex = dataGridView16.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView16.Rows[rowindex];
                    ID6 = row.Cells["ID"].Value.ToString();
                    dt6 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001F = row.Cells["品號"].Value.ToString();
                    MB002F = row.Cells["品名"].Value.ToString();
                    MB003F = row.Cells["規格"].Value.ToString();
                    BOX = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    SUM5 = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    TA029 = row.Cells["備註"].Value.ToString();

                    TA026D = row.Cells["訂單單別"].Value.ToString();
                    TA027D = row.Cells["訂單號"].Value.ToString();
                    TA028D = row.Cells["訂單序號"].Value.ToString();

                    SUBID5 = row.Cells["ID"].Value.ToString();
                    SUBBAR5 = "";
                    SUBNUM5 = "";
                    SUBBOX5 = row.Cells["箱數"].Value.ToString();
                    SUBPACKAGE5 = row.Cells["包裝數"].Value.ToString();

                    SEARCHMOCMANULINERESULT();
                    //SEARCHMOCMANULINECOP();

                }
                else
                {
                    ID6 = null;
                    SUBID5 = null;
                    SUBBAR5 = null;
                    SUBNUM5 = null;
                    SUBBOX5 = null;
                    SUBPACKAGE5 = null;

                    TA026D = null;
                    TA027D = null;
                    TA028D = null;

                }
            }
        }

        private void dataGridView17_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView17.CurrentRow != null)
            {
                int rowindex = dataGridView17.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView17.Rows[rowindex];
                    DELID6= row.Cells["SID"].Value.ToString();
                    DELMOCTA001F = row.Cells["製令"].Value.ToString();
                    DELMOCTA002F = row.Cells["單號"].Value.ToString();



                }
                else
                {
                    DELID2 = null;

                }
            }
        }

        public void SEARCHCOPDEFAULT(string TD001,string TD002,string TD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015,TD013");
                sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(MD002,'')<>'' THEN (TD008+TD024)*MD004 ELSE (TD008+TD024)  END ) AS NUM");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND MB001=TD004");
                sbSql.AppendFormat(@"  AND TD001='{0}' AND TD002='{1}' AND TD003='{2}'",TD001,TD002,TD003);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter27 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder27 = new SqlCommandBuilder(adapter27);
                sqlConn.Open();
                ds27.Clear();
                adapter27.Fill(ds27, "ds27");
                sqlConn.Close();


                if (MANU.Equals("新廠製二組"))
                {
                    if (ds27.Tables["ds27"].Rows.Count == 0)
                    {
                        textBox1.Text = null;
                        textBox2.Text = null;
                        textBox3.Text = null;
                        textBox5.Text = null;
                        textBox6.Text = null;
                        textBox52.Text = null;
                        textBox40.Text = null;
                        textBox41.Text = null;
                        textBox73.Text = null;
                    }
                    else
                    {
                        if (ds27.Tables["ds27"].Rows.Count >= 1)
                        {
                            textBox1.Text = ds27.Tables["ds27"].Rows[0]["TD004"].ToString();
                            textBox2.Text = ds27.Tables["ds27"].Rows[0]["TD005"].ToString();
                            textBox3.Text = ds27.Tables["ds27"].Rows[0]["TD006"].ToString();
                            //textBox5.Text = ds27.Tables["ds27"].Rows[0]["NUM"].ToString();
                            textBox6.Text = ds27.Tables["ds27"].Rows[0]["TC053"].ToString();
                            textBox52.Text = ds27.Tables["ds27"].Rows[0]["TC015"].ToString();
                            dateTimePicker14.Value = Convert.ToDateTime(ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(0, 4) + "/" + ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(4, 2) + "/" + ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(6, 2));

                            if (SUM11>0)
                            {
                                textBox5.Text = (SUM11 + Convert.ToDecimal(ds27.Tables["ds27"].Rows[0]["NUM"].ToString())).ToString();

                                SUM11 = 0;
                            }
                            else
                            {
                                textBox5.Text = ds27.Tables["ds27"].Rows[0]["NUM"].ToString();
                            }
                        }
                    }
                }
                else if (MANU.Equals("新廠包裝線"))
                {
                    if (ds27.Tables["ds27"].Rows.Count == 0)
                    {
                        textBox7.Text = null;
                        textBox10.Text = null;
                        textBox11.Text = null;
                        textBox12.Text = null;
                        textBox53.Text = null;
                        textBox9.Text = null;
                        textBox42.Text = null;
                        textBox43.Text = null;
                        textBox72.Text = null;
                    }
                    else
                    {
                        if (ds27.Tables["ds27"].Rows.Count >= 1)
                        {
                            textBox7.Text = ds27.Tables["ds27"].Rows[0]["TD004"].ToString();
                            textBox10.Text = ds27.Tables["ds27"].Rows[0]["TD005"].ToString();
                            textBox11.Text = ds27.Tables["ds27"].Rows[0]["TD006"].ToString();                            
                            textBox9.Text = ds27.Tables["ds27"].Rows[0]["TC053"].ToString();
                            textBox53.Text = ds27.Tables["ds27"].Rows[0]["TC015"].ToString();
                            dateTimePicker5.Value = Convert.ToDateTime(ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(0, 4) + "/" + ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(4, 2) + "/" + ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(6, 2));

                            if (SUM21 > 0)
                            {
                                textBox12.Text = (SUM21 + Convert.ToDecimal(ds27.Tables["ds27"].Rows[0]["NUM"].ToString())).ToString();

                                SUM21 = 0;
                            }
                            else
                            {
                                textBox12.Text = ds27.Tables["ds27"].Rows[0]["NUM"].ToString();
                            }
                        }
                    }
                }
                else if (MANU.Equals("新廠製一組"))
                {
                    if (ds27.Tables["ds27"].Rows.Count == 0)
                    {
                        textBox14.Text = null;
                        textBox17.Text = null;
                        textBox18.Text = null;
                        textBox19.Text = null;
                        textBox16.Text = null;
                        textBox54.Text = null;
                        textBox44.Text = null;
                        textBox45.Text = null;
                        textBox74.Text = null;
                    }
                    else
                    {
                        if (ds27.Tables["ds27"].Rows.Count >= 1)
                        {
                            textBox14.Text = ds27.Tables["ds27"].Rows[0]["TD004"].ToString();
                            textBox17.Text = ds27.Tables["ds27"].Rows[0]["TD005"].ToString();
                            textBox18.Text = ds27.Tables["ds27"].Rows[0]["TD006"].ToString();
                            //textBox19.Text = ds27.Tables["ds27"].Rows[0]["NUM"].ToString();
                            textBox16.Text = ds27.Tables["ds27"].Rows[0]["TC053"].ToString();
                            textBox54.Text = ds27.Tables["ds27"].Rows[0]["TC015"].ToString();
                            dateTimePicker15.Value = Convert.ToDateTime(ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(0, 4) + "/" + ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(4, 2) + "/" + ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(6, 2));

                            if (SUM31 > 0)
                            {
                                textBox19.Text = (SUM31 + Convert.ToDecimal(ds27.Tables["ds27"].Rows[0]["NUM"].ToString())).ToString();

                                SUM31 = 0;
                            }
                            else
                            {
                                textBox19.Text = ds27.Tables["ds27"].Rows[0]["NUM"].ToString();
                            }
                        }
                    }
                }
                else if (MANU.Equals("新廠製三組(手工)"))
                {
                    if (ds27.Tables["ds27"].Rows.Count == 0)
                    {
                        textBox20.Text = null;
                        textBox24.Text = null;
                        textBox25.Text = null;
                        textBox23.Text = null;
                        textBox22.Text = null;
                        textBox55.Text = null;
                        textBox46.Text = null;
                        textBox47.Text = null;
                        textBox75.Text = null;
                    }
                    else
                    {
                        if (ds27.Tables["ds27"].Rows.Count >= 1)
                        {
                            textBox20.Text = ds27.Tables["ds27"].Rows[0]["TD004"].ToString();
                            textBox24.Text = ds27.Tables["ds27"].Rows[0]["TD005"].ToString();
                            textBox25.Text = ds27.Tables["ds27"].Rows[0]["TD006"].ToString();
                            //textBox23.Text = ds27.Tables["ds27"].Rows[0]["NUM"].ToString();
                            textBox22.Text = ds27.Tables["ds27"].Rows[0]["TC053"].ToString();
                            textBox55.Text = ds27.Tables["ds27"].Rows[0]["TC015"].ToString();
                            dateTimePicker16.Value = Convert.ToDateTime(ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(0, 4) + "/" + ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(4, 2) + "/" + ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(6, 2));

                            if (SUM41 > 0)
                            {
                                textBox23.Text = (SUM41 + Convert.ToDecimal(ds27.Tables["ds27"].Rows[0]["NUM"].ToString())).ToString();

                                SUM41 = 0;
                            }
                            else
                            {
                                textBox23.Text = ds27.Tables["ds27"].Rows[0]["NUM"].ToString();
                            }
                        }
                    }
                }
                else if (MANU.Equals("少量訂單"))
                {
                    if (ds27.Tables["ds27"].Rows.Count == 0)
                    {
                        textBox731.Text = null;
                        textBox721.Text = null;
                        textBox732.Text = null;
                        textBox751.Text = null;
                        textBox771.Text = null;
                        textBox761.Text = null;
                        textBox781.Text = null;
                        textBox782.Text = null;
                        textBox783.Text = null;
                    }
                    else
                    {
                        if (ds27.Tables["ds27"].Rows.Count >= 1)
                        {
                            textBox731.Text = ds27.Tables["ds27"].Rows[0]["TD004"].ToString();
                            textBox721.Text = ds27.Tables["ds27"].Rows[0]["TD005"].ToString();
                            textBox732.Text = ds27.Tables["ds27"].Rows[0]["TD006"].ToString();
                            textBox742.Text = ds27.Tables["ds27"].Rows[0]["NUM"].ToString();
                            textBox751.Text = ds27.Tables["ds27"].Rows[0]["NUM"].ToString();
                            textBox761.Text = ds27.Tables["ds27"].Rows[0]["TC053"].ToString();
                            textBox771.Text = ds27.Tables["ds27"].Rows[0]["TC015"].ToString();

                            dateTimePicker24.Value = Convert.ToDateTime(ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(0,4)+"/"+ ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(4, 2) + "/" + ds27.Tables["ds27"].Rows[0]["TD013"].ToString().Substring(6, 2));

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

        public void SEARCHCOPDEFAULT2(string TD001,string TD002,string TD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                //手工*INVMB.UDF08、其他*INVMB.UDF07
                if (MANU.Equals("新廠製三組(手工)"))
                {
                    sbSql.AppendFormat(@"  SELECT TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015");
                    sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM");
                    sbSql.AppendFormat(@"  ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,INVMB.UDF07");
                    sbSql.AppendFormat(@"  ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ))*INVMB.UDF08/1000 AS 'NUM2'");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD WITH(NOLOCK) ON INVMD.MD001=TD004 AND INVMD.MD002=TD010");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD WITH(NOLOCK) ON BOMMD.MD001=TD004 ");
                    sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                    sbSql.AppendFormat(@"  AND MB001=TD004");
                    sbSql.AppendFormat(@"  AND BOMMD.MD003 LIKE '3%'");
                    sbSql.AppendFormat(@"  AND TD001='{0}' AND TD002='{1}' AND TD003='{2}'", TD001, TD002, TD003);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
              
                else
                {
                    sbSql.AppendFormat(@"  SELECT TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015");
                    sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM");
                    sbSql.AppendFormat(@"  ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,INVMB.UDF07");
                    
                    sbSql.AppendFormat(@"   ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ))/BOMMC.MC004*BOMMD.MD006 AS 'NUM2'");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD WITH(NOLOCK) ON INVMD.MD001=TD004 AND INVMD.MD002=TD010");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC WITH(NOLOCK) ON BOMMC.MC001=TD004 ");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD WITH(NOLOCK) ON BOMMD.MD001=TD004 ");
                    sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                    sbSql.AppendFormat(@"  AND MB001=TD004");
                    sbSql.AppendFormat(@"  AND (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')");
                    sbSql.AppendFormat(@"  AND TD001='{0}' AND TD002='{1}' AND TD003='{2}'", TD001, TD002, TD003);
                    
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }

                //半成品的舊算法
                //sbSql.AppendFormat(@"  ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ))/BOMMC.MC004*INVMB.UDF07/1000 AS 'NUM2'");

                adapter28 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder28 = new SqlCommandBuilder(adapter28);
                sqlConn.Open();
                ds28.Clear();
                adapter28.Fill(ds28, "ds28");
                sqlConn.Close();


                if (MANU.Equals("新廠製二組"))
                {
                    if (ds28.Tables["ds28"].Rows.Count == 0)
                    {
                        textBox1.Text = null;
                        textBox2.Text = null;
                        textBox3.Text = null;
                        textBox5.Text = null;
                        textBox6.Text = null;
                        textBox52.Text = null;
                        textBox40.Text = null;
                        textBox41.Text = null;
                        textBox73.Text = null;
                    }
                    else
                    {
                        if (ds28.Tables["ds28"].Rows.Count >= 1)
                        {
                            textBox1.Text = ds28.Tables["ds28"].Rows[0]["MD003"].ToString();
                            textBox2.Text = ds28.Tables["ds28"].Rows[0]["MD035"].ToString();
                            textBox3.Text = ds28.Tables["ds28"].Rows[0]["MD036"].ToString();
                            textBox5.Text = ds28.Tables["ds28"].Rows[0]["NUM2"].ToString();
                            textBox6.Text = ds28.Tables["ds28"].Rows[0]["TC053"].ToString();
                            textBox52.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();

                        }
                    }
                }
               
                else if (MANU.Equals("新廠製一組"))
                {
                    if (ds28.Tables["ds28"].Rows.Count == 0)
                    {
                        textBox14.Text = null;
                        textBox17.Text = null;
                        textBox18.Text = null;
                        textBox19.Text = null;
                        textBox16.Text = null;
                        textBox54.Text = null;
                        textBox44.Text = null;
                        textBox45.Text = null;
                        textBox74.Text = null;
                    }
                    else
                    {
                        if (ds28.Tables["ds28"].Rows.Count >= 1)
                        {
                            textBox14.Text = ds28.Tables["ds28"].Rows[0]["MD003"].ToString();
                            textBox17.Text = ds28.Tables["ds28"].Rows[0]["MD035"].ToString();
                            textBox18.Text = ds28.Tables["ds28"].Rows[0]["MD036"].ToString();
                            textBox19.Text = ds28.Tables["ds28"].Rows[0]["NUM2"].ToString();
                            textBox16.Text = ds28.Tables["ds28"].Rows[0]["TC053"].ToString();
                            textBox54.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();

                        }
                    }
                }
                else if (MANU.Equals("新廠製三組(手工)"))
                {
                    if (ds28.Tables["ds28"].Rows.Count == 0)
                    {
                        textBox20.Text = null;
                        textBox24.Text = null;
                        textBox25.Text = null;
                        textBox23.Text = null;
                        textBox22.Text = null;
                        textBox55.Text = null;
                        textBox46.Text = null;
                        textBox47.Text = null;
                        textBox75.Text = null;
                    }
                    else
                    {
                        if (ds28.Tables["ds28"].Rows.Count >= 1)
                        {
                            textBox20.Text = ds28.Tables["ds28"].Rows[0]["MD003"].ToString();
                            textBox24.Text = ds28.Tables["ds28"].Rows[0]["MD035"].ToString();
                            textBox25.Text = ds28.Tables["ds28"].Rows[0]["MD036"].ToString();
                            textBox23.Text = ds28.Tables["ds28"].Rows[0]["NUM2"].ToString();
                            textBox22.Text = ds28.Tables["ds28"].Rows[0]["TC053"].ToString();
                            textBox55.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();

                        }
                    }
                }

                else if (MANU.Equals("新廠包裝線"))
                {
                    if (ds28.Tables["ds28"].Rows.Count == 0)
                    {
                        textBox7.Text = null;
                        textBox10.Text = null;                    
                        textBox11.Text = null;
                        textBox12.Text = null;
                        textBox9.Text = null;
                        textBox53.Text = null;
                        textBox42.Text = null;
                        textBox43.Text = null;
                        textBox72.Text = null;
                    }
                    else
                    {
                        if (ds28.Tables["ds28"].Rows.Count >= 1)
                        {
                            textBox7.Text = ds28.Tables["ds28"].Rows[0]["MD003"].ToString();
                            textBox10.Text = ds28.Tables["ds28"].Rows[0]["MD035"].ToString();
                            textBox11.Text = ds28.Tables["ds28"].Rows[0]["MD036"].ToString();
                            textBox12.Text = ds28.Tables["ds28"].Rows[0]["NUM2"].ToString();
                            textBox9.Text = ds28.Tables["ds28"].Rows[0]["TC053"].ToString();
                            textBox53.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();

                        }
                    }
                }
                else if (MANU.Equals("少量訂單"))
                {
                    if (ds28.Tables["ds28"].Rows.Count == 0)
                    {
                        textBox731.Text = null;
                        textBox721.Text = null;
                        textBox732.Text = null;
                        textBox751.Text = null;
                        textBox771.Text = null;
                        textBox761.Text = null;
                        textBox781.Text = null;
                        textBox782.Text = null;
                        textBox783.Text = null;
                    }
                    else
                    {
                        if (ds28.Tables["ds28"].Rows.Count >= 1)
                        {
                            textBox731.Text = ds28.Tables["ds28"].Rows[0]["MD003"].ToString();
                            textBox721.Text = ds28.Tables["ds28"].Rows[0]["MD035"].ToString();
                            textBox732.Text = ds28.Tables["ds28"].Rows[0]["MD036"].ToString();
                            textBox751.Text = ds28.Tables["ds28"].Rows[0]["NUM2"].ToString();
                            textBox761.Text = ds28.Tables["ds28"].Rows[0]["TC053"].ToString();
                            textBox771.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();

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


        public void SEARCHCOPDEFAULT3(string TD001, string TD002, string TD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                //手工*INVMB.UDF08、其他*INVMB.UDF07
                if (MANU.Equals("新廠製三組(手工)"))
                {
                    sbSql.AppendFormat(@"  SELECT TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015");
                    sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM");
                    sbSql.AppendFormat(@"  ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,INVMB.UDF07");
                    sbSql.AppendFormat(@"  ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ))*INVMB.UDF08/1000 AS 'NUM2'");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD WITH(NOLOCK) ON INVMD.MD001=TD004 AND INVMD.MD002=TD010");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD WITH(NOLOCK) ON BOMMD.MD001=TD004 ");
                    sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                    sbSql.AppendFormat(@"  AND MB001=TD004");
                    sbSql.AppendFormat(@"  AND BOMMD.MD003 LIKE '3%'");
                    sbSql.AppendFormat(@"  AND TD001='{0}' AND TD002='{1}' AND TD003='{2}'", TD001, TD002, TD003);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }

                else
                {
                    sbSql.AppendFormat(@"  SELECT TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015");
                    sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM");
                    sbSql.AppendFormat(@"  ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,INVMB.UDF07");

                    sbSql.AppendFormat(@"   ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ))/BOMMC.MC004*BOMMD.MD006 AS 'NUM2'");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD WITH(NOLOCK) ON INVMD.MD001=TD004 AND INVMD.MD002=TD010");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC WITH(NOLOCK) ON BOMMC.MC001=TD004 ");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD WITH(NOLOCK) ON BOMMD.MD001=TD004 ");
                    sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                    sbSql.AppendFormat(@"  AND MB001=TD004");
                    sbSql.AppendFormat(@"  AND (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%') ");
                    sbSql.AppendFormat(@"  AND TD001='{0}' AND TD002='{1}' AND TD003='{2}'", TD001, TD002, TD003);

                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }

                adapter28 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder28 = new SqlCommandBuilder(adapter28);
                sqlConn.Open();
                ds28.Clear();
                adapter28.Fill(ds28, "ds28");
                sqlConn.Close();


                if (MANU.Equals("新廠製二組"))
                {
                    if (ds28.Tables["ds28"].Rows.Count == 0)
                    {
                        textBox1.Text = null;
                        textBox2.Text = null;
                        textBox3.Text = null;
                        textBox5.Text = null;
                        textBox6.Text = null;
                        textBox52.Text = null;
                        textBox40.Text = null;
                        textBox41.Text = null;
                        textBox73.Text = null;
                    }
                    else
                    {
                        if (ds28.Tables["ds28"].Rows.Count >= 1)
                        {
                            textBox1.Text = ds28.Tables["ds28"].Rows[0]["MD003"].ToString();
                            textBox2.Text = ds28.Tables["ds28"].Rows[0]["MD035"].ToString();
                            textBox3.Text = ds28.Tables["ds28"].Rows[0]["MD036"].ToString();
                            textBox5.Text = ds28.Tables["ds28"].Rows[0]["NUM2"].ToString();
                            textBox6.Text = ds28.Tables["ds28"].Rows[0]["TC053"].ToString();
                            textBox52.Text = null;
                            //textBox52.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();

                        }
                    }
                }

                else if (MANU.Equals("新廠製一組"))
                {
                    if (ds28.Tables["ds28"].Rows.Count == 0)
                    {
                        textBox14.Text = null;
                        textBox17.Text = null;
                        textBox18.Text = null;
                        textBox19.Text = null;
                        textBox16.Text = null;
                        textBox54.Text = null;
                        textBox44.Text = null;
                        textBox45.Text = null;
                        textBox74.Text = null;
                    }
                    else
                    {
                        if (ds28.Tables["ds28"].Rows.Count >= 1)
                        {
                            textBox14.Text = ds28.Tables["ds28"].Rows[0]["MD003"].ToString();
                            textBox17.Text = ds28.Tables["ds28"].Rows[0]["MD035"].ToString();
                            textBox18.Text = ds28.Tables["ds28"].Rows[0]["MD036"].ToString();
                            textBox19.Text = ds28.Tables["ds28"].Rows[0]["NUM2"].ToString();
                            textBox16.Text = ds28.Tables["ds28"].Rows[0]["TC053"].ToString();
                            textBox54.Text = null;
                            //textBox54.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();

                        }
                    }
                }
                else if (MANU.Equals("新廠製三組(手工)"))
                {
                    if (ds28.Tables["ds28"].Rows.Count == 0)
                    {
                        textBox20.Text = null;
                        textBox24.Text = null;
                        textBox25.Text = null;
                        textBox23.Text = null;
                        textBox22.Text = null;
                        textBox55.Text = null;
                        textBox46.Text = null;
                        textBox47.Text = null;
                        textBox75.Text = null;
                    }
                    else
                    {
                        if (ds28.Tables["ds28"].Rows.Count >= 1)
                        {
                            textBox20.Text = ds28.Tables["ds28"].Rows[0]["MD003"].ToString();
                            textBox24.Text = ds28.Tables["ds28"].Rows[0]["MD035"].ToString();
                            textBox25.Text = ds28.Tables["ds28"].Rows[0]["MD036"].ToString();
                            textBox23.Text = ds28.Tables["ds28"].Rows[0]["NUM2"].ToString();
                            textBox22.Text = ds28.Tables["ds28"].Rows[0]["TC053"].ToString();
                            textBox55.Text = null;
                            //textBox55.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();

                        }
                    }
                }

                else if (MANU.Equals("新廠包裝線"))
                {
                    if (ds28.Tables["ds28"].Rows.Count == 0)
                    {
                        textBox7.Text = null;
                        textBox10.Text = null;
                        textBox11.Text = null;
                        textBox12.Text = null;
                        textBox9.Text = null;
                        textBox53.Text = null;
                        textBox42.Text = null;
                        textBox43.Text = null;
                        textBox72.Text = null;
                    }
                    else
                    {
                        if (ds28.Tables["ds28"].Rows.Count >= 1)
                        {
                            textBox7.Text = ds28.Tables["ds28"].Rows[0]["MD003"].ToString();
                            textBox10.Text = ds28.Tables["ds28"].Rows[0]["MD035"].ToString();
                            textBox11.Text = ds28.Tables["ds28"].Rows[0]["MD036"].ToString();
                            textBox12.Text = ds28.Tables["ds28"].Rows[0]["NUM2"].ToString();
                            textBox9.Text = ds28.Tables["ds28"].Rows[0]["TC053"].ToString();
                            textBox53.Text = null;
                            //textBox53.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();

                        }
                    }
                }
                else if (MANU.Equals("少量訂單"))
                {
                    if (ds28.Tables["ds28"].Rows.Count == 0)
                    {
                        textBox731.Text = null;
                        textBox721.Text = null;
                        textBox732.Text = null;
                        textBox751.Text = null;
                        textBox771.Text = null;
                        textBox761.Text = null;
                        textBox781.Text = null;
                        textBox782.Text = null;
                        textBox783.Text = null;
                    }
                    else
                    {
                        if (ds28.Tables["ds28"].Rows.Count >= 1)
                        {
                            textBox731.Text = ds28.Tables["ds28"].Rows[0]["MD003"].ToString();
                            textBox721.Text = ds28.Tables["ds28"].Rows[0]["MD035"].ToString();
                            textBox732.Text = ds28.Tables["ds28"].Rows[0]["MD036"].ToString();
                            textBox751.Text = ds28.Tables["ds28"].Rows[0]["NUM2"].ToString();
                            textBox761.Text = ds28.Tables["ds28"].Rows[0]["TC053"].ToString();
                            //textBox761.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();
                            textBox771.Text = null;

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

        public void SEARCHMOCMANULINECOP(string SID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [SERNO] AS '優先序',[TC001] AS '訂單單別',[TC002] AS '訂單單號',[TC003] AS '訂單序號',[NUM] AS '需求量',[ID],[SID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINECOP]");
                sbSql.AppendFormat(@"  WHERE [SID]='{0}'",SID);
                sbSql.AppendFormat(@"  ORDER BY [SERNO]");
                sbSql.AppendFormat(@"  ");

                adapter29 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder29 = new SqlCommandBuilder(adapter29);
                sqlConn.Open();
                ds29.Clear();
                adapter29.Fill(ds29, "ds29");
                sqlConn.Close();


                if (ds29.Tables["ds29"].Rows.Count == 0)
                {
                    dataGridView11.DataSource = null;
                }
                else
                {
                    if (ds29.Tables["ds29"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView11.DataSource = ds29.Tables["ds29"];
                        dataGridView11.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        public void INSERTMOCMANULINECOP(string SID,string TA001,string TA002,string TA003,string SERNO)
        {
            decimal TNUM = SEARCHINVMD(TA001.Trim(), TA002.Trim(), TA003.Trim());

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINECOP]");
                sbSql.AppendFormat(" ([ID],[SID],[SERNO],[TC001],[TC002],[TC003],[NUM])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" (NEWID(),'{0}','{1}','{2}','{3}','{4}',{5})",ID1,SERNO,TA001,TA002,TA003, TNUM);
                sbSql.AppendFormat(" ");
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

        private void dataGridView11_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView11.CurrentRow != null)
            {
                int rowindex = dataGridView11.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView11.Rows[rowindex];
                    DELMOCMANULINECOPID = row.Cells["ID"].Value.ToString();
                    
                }
                else
                {
                    DELMOCMANULINECOPID = null;

                }
            }
        }

        public void DELMOCMANULINECOP(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINECOP]");
                sbSql.AppendFormat("  WHERE ID='{0}'", ID);
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

        public decimal SEARCHINVMD(string TA001,string TA002,string TA003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                sbSql.AppendFormat(" SELECT '{0}',TD001,TD002,TD003,TD004,TD005,NUM,MB004,MD003,MD035,CASE WHEN [MD003] LIKE '2%' THEN ROUND((NUM*CAL),0) ELSE (NUM*CAL) END AS TNUM,MD004", TA001+TA002+TA003);
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT   TD001,TD002,TD003,TC053 ,TD013,TD004,TD005,TD006");
                sbSql.AppendFormat(" ,((CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*INVMD.MD004 END)-ISNULL(MOCTA.TA017,0)) AS 'NUM'");
                sbSql.AppendFormat(" ,MB004");
                sbSql.AppendFormat(" ,((TD008-TD009)+(TD024-TD025)) AS 'COPNUM'");
                sbSql.AppendFormat(" ,TD010");
                sbSql.AppendFormat(" ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN INVMD.MD002 ELSE TD010 END ) AS INVMDMD002");
                sbSql.AppendFormat(" ,(CASE WHEN INVMD.MD003>0 THEN INVMD.MD003 ELSE 1 END) AS INVMDMD003");
                sbSql.AppendFormat(" ,(CASE WHEN INVMD.MD004>0 THEN INVMD.MD004 ELSE (TD008-TD009) END ) AS INVMDMD004");
                sbSql.AppendFormat(" ,ISNULL(MOCTA.TA017,0) AS TA017");
                sbSql.AppendFormat(" ,[MC001],[MC004],BOMMD.[MD003],[MD035],BOMMD.[MD006],BOMMD.[MD007],BOMMD.[MD008],BOMMD.[MD004]");
                sbSql.AppendFormat(" ,CONVERT(decimal(16,4),(1/[MC004]*BOMMD.[MD006]/BOMMD.[MD007]*(1+BOMMD.[MD008]))) AS CAL");
                sbSql.AppendFormat(" FROM [TK].dbo.BOMMC,[TK].dbo.BOMMD,[TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.MOCTA ON TA026=TD001 AND TA027=TD002 AND TD028=TD003 AND TA006=TD004");
                sbSql.AppendFormat(" WHERE BOMMC.MC001=BOMMD.MD001");
                sbSql.AppendFormat(" AND  BOMMD.MD001=TD004");
                sbSql.AppendFormat(" AND TD004=MB001");
                sbSql.AppendFormat(" AND TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(" AND TD001+TD002+TD003 IN ('{0}')", TA001 + TA002 + TA003);
                sbSql.AppendFormat(" ) AS TEMP");
                sbSql.AppendFormat("  WHERE MD003 LIKE '3%'");
                sbSql.AppendFormat(" ");

                adapter30 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder30 = new SqlCommandBuilder(adapter30);
                sqlConn.Open();
                ds30.Clear();
                adapter30.Fill(ds30, "ds30");
                sqlConn.Close();


                if (ds30.Tables["ds30"].Rows.Count == 0)
                {
                    return 0;
                }
                else
                {
                    if (ds30.Tables["ds30"].Rows.Count >= 1)
                    {
                        return Convert.ToDecimal(ds30.Tables["ds30"].Rows[0]["TNUM"].ToString());
                    }
                    else
                    {
                        return 0;
                    }
                    
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

        public void SEARCHMOCMANULINE12(string MANU, string SDAY, string EDAY,string SATUS)
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if(SATUS.Equals("過濾已合併的"))
                {
                    Query.AppendFormat(@" AND [ID] NOT IN (SELECT [SID]  FROM [TKMOC].[dbo].[MOCMANULINEMERGE]) ");
                }
                else
                {
                    Query.AppendFormat(@" ");
                }

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'");
                sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BAR] AS '桶數',[NUM] AS '數量',[BOX] AS '箱數',[PACKAGE] AS '包裝數',[CLINET] AS '客戶',[MANUHOUR] AS '生產時間',[OUTDATE] AS '交期',[TA029] AS '備註',[HALFPRO] AS '半成品數量'");
                sbSql.AppendFormat(@"  ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                sbSql.AppendFormat(@"  WHERE [MANU]='{0}' ", MANU);
                sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112)>='{0}' AND CONVERT(varchar(100),[MANUDATE],112)<='{1}'", SDAY, EDAY);
                sbSql.AppendFormat(@"  {0}",Query.ToString());
                sbSql.AppendFormat(@"  ORDER BY [MB001],[MANUDATE],[SERNO]");
                sbSql.AppendFormat(@"  ");

                adapter31 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder31 = new SqlCommandBuilder(adapter31);
                sqlConn.Open();
                ds31.Clear();
                adapter31.Fill(ds31, "ds31");
                sqlConn.Close();


                if (ds31.Tables["ds31"].Rows.Count == 0)
                {
                    dataGridView12.DataSource = null;
                }
                else
                {
                    if (ds31.Tables["ds31"].Rows.Count >= 1)
                    {
                        if(dataGridView12.Columns.Count>0)
                        {
                            this.dataGridView12.Columns.RemoveAt(0);
                        }


                        //dataGridView1.Rows.Clear();
                        dataGridView12.DataSource = ds31.Tables["ds31"];
                        dataGridView12.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        //建立一個DataGridView的Column物件及其內容
                        DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
                        dgvc.Width = 40;
                        dgvc.Name = "選取";

                        this.dataGridView12.Columns.Insert(0, dgvc);

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

        public void SEARCHMOCMANULINEMERGE(DateTime dt)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MOCMANULINEMERGE].[NO] AS '編號', [MOCMANULINE].[MANU] AS '線別',[MOCMANULINE].[MB001] AS '品號',[MOCMANULINE].[MB002] AS '品名',[MOCMANULINE].[BAR] AS '桶數',[MOCMANULINE].[NUM] AS '數量',[MOCMANULINE].[BOX] AS '箱數',[MOCMANULINE].[PACKAGE] AS '包裝數'");
                sbSql.AppendFormat(@"  ,[MOCMANULINEMERGE].[ID],[MOCMANULINEMERGE].[SID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINEMERGE],[TKMOC].[dbo].[MOCMANULINE]");
                sbSql.AppendFormat(@"  WHERE [MOCMANULINEMERGE].[SID]=[MOCMANULINE].[ID]");
                sbSql.AppendFormat(@"  AND [MOCMANULINEMERGE].[NO] LIKE '{0}%'",dt.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY [MOCMANULINEMERGE].[NO]");
                sbSql.AppendFormat(@"  ");

                adapter33 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder33 = new SqlCommandBuilder(adapter33);
                sqlConn.Open();
                ds33.Clear();
                adapter33.Fill(ds33, "ds33");
                sqlConn.Close();


                if (ds33.Tables["ds33"].Rows.Count == 0)
                {
                    dataGridView13.DataSource = null;
                }
                else
                {
                    if (ds33.Tables["ds33"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView13.DataSource = ds33.Tables["ds33"];
                        dataGridView13.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];


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

        public void INSERTMOCMANULINEMERGE(DateTime dt)
        {
            string NO = GETMAXNOMOCMANULINEMERGE(dt);

            foreach (DataGridViewRow dr in this.dataGridView12.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    if(!string.IsNullOrEmpty(NO)&& !string.IsNullOrEmpty(dr.Cells["ID"].Value.ToString()))
                    {
                        ADDMOCMANULINEMERGE(NO.Trim(), dr.Cells["ID"].Value.ToString().Trim());
                    }
                   

                    //dr.Cells["ID"].Value.ToString();
                    //MessageBox.Show(NO+" "+dr.Cells["ID"].Value.ToString());
                }
            }
        }

        public string GETMAXNOMOCMANULINEMERGE(DateTime dt)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(NO),'00000000000') AS NO");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINEMERGE] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE [NO] LIKE '{0}%' ", dt.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter32 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder32 = new SqlCommandBuilder(adapter32);
                sqlConn.Open();
                ds32.Clear();
                adapter32.Fill(ds32, "ds32");
                sqlConn.Close();


                if (ds32.Tables["ds32"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds32.Tables["ds32"].Rows.Count >= 1)
                    {
                        TA002 = SETMAXNOMOCMANULINEMERG(dt,ds32.Tables["ds32"].Rows[0]["NO"].ToString());
                        return TA002;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public string SETMAXNOMOCMANULINEMERG(DateTime dt,string NO)
        {
            if (NO.Equals("00000000000"))
            {
                return dt.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(NO.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt.ToString("yyyyMMdd") + temp.ToString();
            }

          
        }

        public void ADDMOCMANULINEMERGE(string NO,string SID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINEMERGE]");
                sbSql.AppendFormat(" ([ID],[NO],[SID])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", Guid.NewGuid(),NO, SID);
                sbSql.AppendFormat(" ");
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

        private void dataGridView13_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView13.CurrentRow != null)
            {
                int rowindex = dataGridView13.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView13.Rows[rowindex];
                    textBox78.Text = row.Cells["編號"].Value.ToString();

                    SEARCHMOCTATA020(row.Cells["編號"].Value.ToString());
                    SEARCHMOCMANULINENO(row.Cells["編號"].Value.ToString());
                }
                else
                {
                    textBox78.Text = null;

                }
            }
        }


        public void CALSUMMOCMANULINEMERGE(string NO)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT [MOCMANULINEMERGE].[NO] AS '編號',[MOCMANULINE].[MB001] AS '品號',[INVMB].MB002 AS '品名',[INVMB].MB003 AS '規格',SUM([MOCMANULINE].[BAR]) AS '加總桶數',SUM([MOCMANULINE].[NUM]) AS '加總數量',SUM([MOCMANULINE].[BOX]) AS '加總箱數',SUM([MOCMANULINE].[PACKAGE]) AS '加總包裝數' 
                                    ,SUBSTRING( 
                                    ( 
                                    SELECT ',' +[MOCMANULINE].TA029 AS 'data()'
                                    FROM   [TKMOC].[dbo].[MOCMANULINEMERGE],[TKMOC].[dbo].[MOCMANULINE]
                                    WHERE [MOCMANULINEMERGE].[SID]=[MOCMANULINE].[ID]
                                    AND [MOCMANULINEMERGE].[NO]='{0}' FOR XML PATH('') 
                                    ), 2 , 250) As 備註
                                    FROM [TKMOC].[dbo].[MOCMANULINEMERGE],[TKMOC].[dbo].[MOCMANULINE],[TK].dbo.[INVMB]
                                    WHERE [INVMB].MB001=[MOCMANULINE].[MB001]
                                    AND [MOCMANULINEMERGE].[SID]=[MOCMANULINE].[ID]
                                    AND [MOCMANULINEMERGE].[NO]='{0}'
                                    GROUP BY [MOCMANULINEMERGE].[NO],[MOCMANULINE].[MB001],[INVMB].MB002,[INVMB].MB003
                                    ", NO);

                adapter34 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder34 = new SqlCommandBuilder(adapter34);
                sqlConn.Open();
                ds34.Clear();
                adapter34.Fill(ds34, "ds34");
                sqlConn.Close();


                if (ds34.Tables["ds34"].Rows.Count == 0)
                {
                    dataGridView14.DataSource = null;
                }
                else
                {
                    if (ds34.Tables["ds34"].Rows.Count >= 1)
                    {
                        ADDMOCMANULINEMERGERESLUT(ds34);

                        //dataGridView1.Rows.Clear();
                        dataGridView14.DataSource = ds34.Tables["ds34"];
                        dataGridView14.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                      

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

        public void SEARCHMOCTATA020(string TA033)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA001 AS '製令單別',TA002 AS '製令單號',TA020 AS '編號'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  WHERE TA033='{0}'", TA033);
                sbSql.AppendFormat(@"  ");

                adapter36 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder36 = new SqlCommandBuilder(adapter36);
                sqlConn.Open();
                ds36.Clear();
                adapter36.Fill(ds36, "ds36");
                sqlConn.Close();


                if (ds36.Tables["ds36"].Rows.Count == 0)
                {
                    dataGridView18.DataSource = null;
                }
                else
                {
                    if (ds36.Tables["ds36"].Rows.Count >= 1)
                    {
                     
                        //dataGridView1.Rows.Clear();
                        dataGridView18.DataSource = ds36.Tables["ds36"];
                        dataGridView18.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];



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

        public void SEARCHMOCMANULINENO(string NO)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MOCMANULINE].[MANU] AS '線別',CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS '預排日',[MOCMANULINE].[MB001] AS '品號',[MOCMANULINE].[MB002] AS '品名',[MOCMANULINE].[NUM] AS '數量'");
                sbSql.AppendFormat(@"  ,[MOCMANULINEMERGE].[ID],[MOCMANULINEMERGE].[NO],[MOCMANULINEMERGE].[SID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINEMERGE],[TKMOC].[dbo].[MOCMANULINE]");
                sbSql.AppendFormat(@"  WHERE [MOCMANULINEMERGE].[SID]=[MOCMANULINE].ID");
                sbSql.AppendFormat(@"  AND [NO]='{0}'",NO);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter37 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder37 = new SqlCommandBuilder(adapter37);
                sqlConn.Open();
                ds37.Clear();
                adapter37.Fill(ds37, "ds37");
                sqlConn.Close();


                if (ds37.Tables["ds37"].Rows.Count == 0)
                {
                    dataGridView19.DataSource = null;
                }
                else
                {
                    if (ds37.Tables["ds37"].Rows.Count >= 1)
                    {

                        //dataGridView1.Rows.Clear();
                        dataGridView19.DataSource = ds37.Tables["ds37"];
                        dataGridView19.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];



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

        public void ADDMOCMANULINEMERGERESLUT(DataSet ds)
        {
            if(ds.Tables[0].Rows.Count>0)
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    if (comboBox15.Text.Equals("新廠包裝線"))
                    {
                        sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[MOCMANULINEMERGERESLUT]");
                        sbSql.AppendFormat(" ");
                        sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINEMERGERESLUT]");
                        sbSql.AppendFormat(" ([NO],[MB001],[MB002],[MB003],[NUM],[BAR],[COMMENT])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", ds34.Tables["ds34"].Rows[0]["編號"].ToString(), ds34.Tables["ds34"].Rows[0]["品號"].ToString(), ds34.Tables["ds34"].Rows[0]["品名"].ToString(), ds34.Tables["ds34"].Rows[0]["規格"].ToString(), ds34.Tables["ds34"].Rows[0]["加總包裝數"].ToString(), ds34.Tables["ds34"].Rows[0]["加總箱數"].ToString(), ds34.Tables["ds34"].Rows[0]["備註"].ToString());
                        sbSql.AppendFormat(" ");
                        sbSql.AppendFormat(" ");
                    }
                    else
                    {
                        sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[MOCMANULINEMERGERESLUT]");
                        sbSql.AppendFormat(" ");
                        sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINEMERGERESLUT]");
                        sbSql.AppendFormat(" ([NO],[MB001],[MB002],[MB003],[NUM],[BAR],[COMMENT])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", ds34.Tables["ds34"].Rows[0]["編號"].ToString(), ds34.Tables["ds34"].Rows[0]["品號"].ToString(), ds34.Tables["ds34"].Rows[0]["品名"].ToString(), ds34.Tables["ds34"].Rows[0]["規格"].ToString(), ds34.Tables["ds34"].Rows[0]["加總數量"].ToString(), ds34.Tables["ds34"].Rows[0]["加總桶數"].ToString(), ds34.Tables["ds34"].Rows[0]["備註"].ToString());
                        sbSql.AppendFormat(" ");
                        sbSql.AppendFormat(" ");
                    }

                   

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

        private void dataGridView14_SelectionChanged(object sender, EventArgs e)
        {
            textBox80.Text = null;
            textBox81.Text = null;
            textBox82.Text = null;
            textBox83.Text = null;
            textBox84.Text = null;
            textBox79.Text = null;

            if (dataGridView14.CurrentRow != null)
            {
                int rowindex = dataGridView14.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView14.Rows[rowindex];
                    textBox80.Text = row.Cells["品號"].Value.ToString();
                    textBox81.Text = row.Cells["品名"].Value.ToString();
                    textBox82.Text = row.Cells["規格"].Value.ToString();
                    textBox83.Text = row.Cells["加總數量"].Value.ToString();
                    textBox84.Text = row.Cells["加總包裝數"].Value.ToString();
                    textBox79.Text = row.Cells["備註"].Value.ToString();
                }
                else
                {
                    textBox80.Text = null;
                    textBox81.Text = null;
                    textBox82.Text = null;
                    textBox83.Text = null;
                    textBox84.Text = null;
                    textBox79.Text = null;

                }
            }
        }

        public string GETMAXTA002MERGE(DateTime dt,string TA001)
        {
            string TA002;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter35 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder35 = new SqlCommandBuilder(adapter35);
                sqlConn.Open();
                ds35.Clear();
                adapter35.Fill(ds35, "ds35");
                sqlConn.Close();


                if (ds35.Tables["ds35"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds35.Tables["ds35"].Rows.Count >= 1)
                    {
                        TA002 = SETTA002MERGE(dt, ds35.Tables["ds35"].Rows[0]["TA002"].ToString());
                        return TA002;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }

        }
        public string SETTA002MERGE(DateTime dt,string TA002)
        {

            if (TA002.Equals("00000000000"))
            {
                return dt.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TA002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt.ToString("yyyyMMdd") + temp.ToString();
            }
          
        }


        public void ADDMOCTATBMERGE(string TA001, string TA002, string TA006, string TA034, string TA035, string TA020,string TA029,string TA021,string TA015,string TA033)
        {
            MOCTADATA MOCTA = new MOCTADATA();
            MOCTA = SETMOCTAMERGE(TA006,dateTimePicker22.Value);

            string MOCMB001 = null;
            decimal MOCTA004 = 0; ;
            string MOCTB009 = null;


            const int MaxLength = 100;

            MOCMB001 = TA006;
            MOCTA004 = BAR;

            MOCTA.TA001 = TA001;
            MOCTA.TA002 = TA002;
            MOCTA.TA006 = TA006;
            MOCTA.TA015 = TA015;
            MOCTA.TA020 = TA020;
            MOCTA.TA021 = TA021;
            MOCTA.TA024 = TA001;
            MOCTA.TA025 = TA002;
            MOCTA.TA029 = TA029;
            MOCTA.TA033 = TA033;
            MOCTA.TA034 = TA034;
            MOCTA.TA035 = TA035;
           
            //MOCTA.TA026 = TA026A;
            //MOCTA.TA027 = TA027A;
            //MOCTA.TA028 = TA028A;

            try
            {
                //check TA002=2,TA040=2
                if (MOCTA.TA002.Substring(0, 1).Equals("2") && MOCTA.TA040.Substring(0, 1).Equals("2"))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTA]");
                    sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_count],[DataGroup],[TA001],[TA002],[TA003],[TA004],[TA005],[TA006],[TA007]");
                    sbSql.AppendFormat(" ,[TA009],[TA010],[TA011],[TA012],[TA013],[TA014],[TA015],[TA016],[TA017],[TA018]");
                    sbSql.AppendFormat(" ,[TA019],[TA020],[TA021],[TA022],[TA024],[TA025],[TA029],[TA030],[TA031],[TA033],[TA034],[TA035]");
                    sbSql.AppendFormat(" ,[TA040],[TA041],[TA043],[TA044],[TA045],[TA046],[TA047],[TA049],[TA050],[TA200]");
                    sbSql.AppendFormat(" ,[TA026],[TA027],[TA028]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" VALUES");
                    sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA.TA003, MOCTA.TA004, MOCTA.TA005, MOCTA.TA006, MOCTA.TA007);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TA009, MOCTA.TA010, MOCTA.TA011, MOCTA.TA012, MOCTA.TA013, MOCTA.TA014, MOCTA.TA015, MOCTA.TA016, MOCTA.TA017, MOCTA.TA018);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}',N'{6}','{7}','{8}','{9}','{10}','{11}',", MOCTA.TA019, MOCTA.TA020, MOCTA.TA021, MOCTA.TA022, MOCTA.TA024, MOCTA.TA025, MOCTA.TA029, MOCTA.TA030, MOCTA.TA031, MOCTA.TA033, MOCTA.TA034, MOCTA.TA035);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TA040, MOCTA.TA041, MOCTA.TA043, MOCTA.TA044, MOCTA.TA045, MOCTA.TA046, MOCTA.TA047, MOCTA.TA049, MOCTA.TA050, MOCTA.TA200);
                    sbSql.AppendFormat(" '{0}','{1}','{2}'", MOCTA.TA026, MOCTA.TA027, MOCTA.TA028);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" INSERT INTO [TK].dbo.[MOCTB]");
                    sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_count],[DataGroup],[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007]");
                    sbSql.AppendFormat(" ,[TB009],[TB011],[TB012],[TB013],[TB014],[TB018],[TB019],[TB020],[TB022],[TB024]");
                    sbSql.AppendFormat(" ,[TB025],[TB026],[TB027],[TB029],[TB030],[TB031],[TB501],[TB554],[TB556],[TB560])");
                    sbSql.AppendFormat(" (SELECT ");
                    sbSql.AppendFormat(" '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE]", MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE);
                    sbSql.AppendFormat(" ,'{0}' [TRANS_NAME],{1} [sync_count],'{2}' [DataGroup],'{3}' [TB001],'{4}' [TB002],[BOMMD].MD003 [TB003],ROUND([MOCMANULINEMERGERESLUT].NUM/MC004*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),3) [TB004],0 [TB005],'****' [TB006],[INVMB].MB004  [TB007]", MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002);
                    sbSql.AppendFormat(" ,[INVMB].MB017 [TB009],'1' [TB011],[INVMB].MB002 [TB012],[INVMB].MB003 [TB013],[BOMMD].MD001 [TB014],'N' [TB018],'0' [TB019],'0' [TB020],'2' [TB022],'0' [TB024]");
                    sbSql.AppendFormat(" ,'****' [TB025],'0' [TB026],'1' [TB027],'0' [TB029],'0' [TB030],'0' [TB031],'0' [TB501],'N' [TB554],'0' [TB556],'0' [TB560]");
                    sbSql.AppendFormat(" FROM [TK].dbo.[BOMMC],[TK].dbo.[BOMMD],[TK].dbo.[INVMB],[TKMOC].[dbo].[MOCMANULINEMERGERESLUT]");
                    sbSql.AppendFormat(" WHERE [BOMMC].MC001=[BOMMD].MD001");
                    sbSql.AppendFormat(" AND [BOMMD].MD003=[INVMB].MB001");
                    sbSql.AppendFormat(" AND MD001=[MOCMANULINEMERGERESLUT].MB001");
                    sbSql.AppendFormat(" AND MD001='{0}' AND ISNULL(MD012,'')=''", TA006);
                    sbSql.AppendFormat(" AND [MOCMANULINEMERGERESLUT].NO='{0}'",TA033);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
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


            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public MOCTADATA SETMOCTAMERGE(string TA006,DateTime dt)
        {
            SEARCHBOMMCMERGE(TA006);

            MOCTADATA MOCTA = new MOCTADATA();
            MOCTA.COMPANY = "TK";
            MOCTA.CREATOR = "140020";
            MOCTA.USR_GROUP = "103000";
            //MOCTA.CREATE_DATE = dt1.ToString("yyyyMMdd");
            MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            MOCTA.MODIFIER = "140020";
            MOCTA.MODI_DATE = dt.ToString("yyyyMMdd");
            MOCTA.FLAG = "0";
            MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            MOCTA.TRANS_TYPE = "P001";
            MOCTA.TRANS_NAME = "MOCMI02";
            MOCTA.sync_count = "0";
            MOCTA.DataGroup = "103000";
            MOCTA.TA001 = "";
            MOCTA.TA002 = "";
            MOCTA.TA003 = dt.ToString("yyyyMMdd");
            MOCTA.TA004 = dt.ToString("yyyyMMdd");
            MOCTA.TA005 = BOMVARSION;
            MOCTA.TA006 = MB001;
            MOCTA.TA007 = UNIT;
            MOCTA.TA009 = dt.ToString("yyyyMMdd");
            MOCTA.TA010 = dt.ToString("yyyyMMdd");
            MOCTA.TA011 = "1";
            MOCTA.TA012 = dt.ToString("yyyyMMdd");
            MOCTA.TA013 = "N";
            //MOCTA.TA014 = dt1.ToString("yyyyMMdd");
            MOCTA.TA014 = "";
            //MOCTA.TA015 = (BAR * BOMBAR).ToString();
            MOCTA.TA015 = "0";
            MOCTA.TA016 = "0";
            MOCTA.TA017 = "0";
            MOCTA.TA018 = "0";
            MOCTA.TA019 = "20";
            MOCTA.TA020 = "";
            MOCTA.TA021 = "02";
            MOCTA.TA022 = "0";
            MOCTA.TA024 = "A510";
            MOCTA.TA025 = "";
            MOCTA.TA029 = "";
            MOCTA.TA030 = "1";
            MOCTA.TA031 = "0";
            MOCTA.TA033= "";
            MOCTA.TA034 = "";
            MOCTA.TA035 = "";
            MOCTA.TA040 = dt.ToString("yyyyMMdd");
            MOCTA.TA041 = "";
            MOCTA.TA043 = "1";
            MOCTA.TA044 = "N";
            MOCTA.TA045 = "0";
            MOCTA.TA046 = "0";
            MOCTA.TA047 = "0";
            MOCTA.TA049 = "0";
            MOCTA.TA050 = "0";
            MOCTA.TA200 = "1";


            return MOCTA;

        }


        public void SEARCHBOMMCMERGE(string MB001)
        {
            BOMVARSION = null;
            UNIT = null;
            BOMBAR = 0;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                sbSql.AppendFormat(@"  ,INVMB.MB004");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001);
                sbSql.AppendFormat(@"  ");

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                dsBOMMC.Clear();
                adapter5.Fill(dsBOMMC, "dsBOMMC");
                sqlConn.Close();


                if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                {
                    BOMVARSION = null;
                    UNIT = null;
                    BOMBAR = 0;
                }
                else
                {
                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                    {
                        BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                        //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                        UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                        BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            label104.Text = comboBox16.SelectedValue.ToString();
           
        }

        public void  DATAGRIDCLEAR()
        {
            

            dataGridView12.DataSource = null;
            dataGridView13.DataSource = null;
            dataGridView14.DataSource = null;
            dataGridView18.DataSource = null;
            dataGridView19.DataSource = null;
        }

        private void textBox731_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
            SEARCHBOMMD();
        }

        private void textBox741_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }

        private void dataGridView20_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView20.CurrentRow != null)
            {
                int rowindex = dataGridView20.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView20.Rows[rowindex];
                    ID10 = row.Cells["ID"].Value.ToString();
                    LIMITSERCHTD002 = row.Cells["訂單號"].Value.ToString();

                    //dt2 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    //MB001B = row.Cells["品號"].Value.ToString();
                    //MB002B = row.Cells["品名"].Value.ToString();
                    //MB003B = row.Cells["規格"].Value.ToString();
                    //BOX = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    //SUM2 = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    //TA029 = row.Cells["備註"].Value.ToString();
                    //TA026 = row.Cells["訂單單別"].Value.ToString();
                    //TA027 = row.Cells["訂單號"].Value.ToString();
                    //TA028 = row.Cells["訂單序號"].Value.ToString();

                    //SUBID2 = row.Cells["ID"].Value.ToString();
                    //SUBBAR2 = "";
                    //SUBNUM2 = "";
                    //SUBBOX2 = row.Cells["箱數"].Value.ToString();
                    //SUBPACKAGE2 = row.Cells["包裝數"].Value.ToString();

                    //SEARCHMOCMANULINERESULT();
                    //SEARCHMOCMANULINECOP();

                }
                else
                {
                    ID10 = null;
                    //SUBID2 = null;
                    //SUBBAR2 = null;
                    //SUBNUM2 = null;
                    //SUBBOX2 = null;
                    //SUBPACKAGE2 = null;
                    //TA026 = null;
                    //TA027 = null;
                    //TA028 = null;

                }
            }
        }

        public void SEARCHMOCMANULINETEMPDATAS(string MB001)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            SqlDataAdapter adapter2= new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder2= new SqlCommandBuilder();
            DataSet ds2 = new DataSet();

            SqlDataAdapter adapter3 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
            DataSet ds3 = new DataSet();

            SqlDataAdapter adapter4 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
            DataSet ds4 = new DataSet();

            SUM11 = 0;
            SUM21 = 0;
            SUM31 = 0;
            SUM41 = 0;

            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@" SELECT [ID]  FROM [TKMOC].[dbo].[MOCMANULINETEMP] WHERE [MB001]='{0}' AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE] )", MB001);
                    sbSql.AppendFormat(@"  ");

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
                            //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                            //dataGridView1.AutoResizeColumns();
                            //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                            TEMPds.Clear();
                            frmMOCMANULINESubTEMPADD MOCMANULINESubTEMPADD = new frmMOCMANULINESubTEMPADD(MB001, TEMPds);
                            MOCMANULINESubTEMPADD.ShowDialog();

                            TEMPds = MOCMANULINESubTEMPADD.SETDATASET;

                            if (TEMPds.Tables[0].Rows.Count >= 1)
                            {
                                foreach (DataRow dr in TEMPds.Tables[0].Rows)
                                {
                                    SUM11 = SUM11 + Convert.ToDecimal(dr["數量"].ToString());
                                    //SUM2 = SUM2 + Convert.ToDecimal(dr["桶數"].ToString());
                                }
                            }
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

            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@" SELECT [ID]  FROM [TKMOC].[dbo].[MOCMANULINETEMP] WHERE [MB001]='{0}' AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE] )", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "ds2");
                    sqlConn.Close();


                    if (ds2.Tables["ds2"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds2.Tables["ds2"].Rows.Count >= 1)
                        {
                            //dataGridView1.Rows.Clear();
                            //dataGridView3.DataSource = ds5.Tables["TEMPds5"];
                            //dataGridView3.AutoResizeColumns();
                            //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                            TEMPds.Clear();
                            frmMOCMANULINESubTEMPADD MOCMANULINESubTEMPADD = new frmMOCMANULINESubTEMPADD(MB001, TEMPds);
                            MOCMANULINESubTEMPADD.ShowDialog();

                            TEMPds = MOCMANULINESubTEMPADD.SETDATASET;

                            if(TEMPds.Tables[0].Rows.Count>= 1)
                            {                               
                                foreach (DataRow dr in TEMPds.Tables[0].Rows)
                                {
                                    SUM21 = SUM21 + Convert.ToDecimal(dr["包裝數"].ToString());
                                    //SUM2 = SUM2 + Convert.ToDecimal(dr["箱數"].ToString());
                                }
                            }
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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@" SELECT [ID]  FROM [TKMOC].[dbo].[MOCMANULINETEMP] WHERE [MB001]='{0}' AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE] )", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                    sqlConn.Open();
                    ds3.Clear();
                    adapter3.Fill(ds3, "ds3");
                    sqlConn.Close();


                    if (ds3.Tables["ds3"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds3.Tables["ds3"].Rows.Count >= 1)
                        {
                            ////dataGridView1.Rows.Clear();
                            //dataGridView5.DataSource = ds7.Tables["TEMPds7"];
                            //dataGridView5.AutoResizeColumns();
                            //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                            TEMPds.Clear();
                            frmMOCMANULINESubTEMPADD MOCMANULINESubTEMPADD = new frmMOCMANULINESubTEMPADD(MB001, TEMPds);
                            MOCMANULINESubTEMPADD.ShowDialog();

                            TEMPds = MOCMANULINESubTEMPADD.SETDATASET;

                            if (TEMPds.Tables[0].Rows.Count >= 1)
                            {                             
                                foreach (DataRow dr in TEMPds.Tables[0].Rows)
                                {
                                    SUM31 = SUM31 + Convert.ToDecimal(dr["數量"].ToString());
                                    //SUM2 = SUM2 + Convert.ToDecimal(dr["桶數"].ToString());
                                }
                                
                            }
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@" SELECT [ID]  FROM [TKMOC].[dbo].[MOCMANULINETEMP] WHERE [MB001]='{0}' AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE] )", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "ds4");
                    sqlConn.Close();


                    if (ds4.Tables["ds4"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds4.Tables["ds4"].Rows.Count >= 1)
                        {
                            //dataGridView1.Rows.Clear();
                            //dataGridView7.DataSource = ds8.Tables["TEMPds8"];
                            //dataGridView7.AutoResizeColumns();
                            //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                            TEMPds.Clear();
                            frmMOCMANULINESubTEMPADD MOCMANULINESubTEMPADD = new frmMOCMANULINESubTEMPADD(MB001, TEMPds);
                            MOCMANULINESubTEMPADD.ShowDialog();

                            TEMPds = MOCMANULINESubTEMPADD.SETDATASET;

                            if (TEMPds.Tables[0].Rows.Count >= 1)
                            {                               
                                foreach (DataRow dr in TEMPds.Tables[0].Rows)
                                {
                                    SUM41 = SUM41 + Convert.ToDecimal(dr["數量"].ToString());
                                    //SUM2 = SUM2 + Convert.ToDecimal(dr["桶數"].ToString());
                                }

                            }
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

            

        }

        public DataSet SETTEMPDATASET
        {
            get
            {
                return TEMPds;
            }
            set
            {              
                TEMPds = value;
            }
        }

        public void UPDATEMOCMANULINETEMP(Guid NEWGUID,DataSet ds)
        {
            StringBuilder IDMOCMANULINETEMP = new StringBuilder();

            if (ds.Tables[0].Rows.Count >= 1)
            {               
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    IDMOCMANULINETEMP.AppendFormat(@"'{0}', ", dr["ID"].ToString());
                   
                }

            }

            IDMOCMANULINETEMP.AppendFormat(@"'d22acdff-fee6-40f4-92cd-acce2a353749' ");

            if(ds.Tables[0].Rows.Count >= 1)
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[MOCMANULINETEMP]");
                    sbSql.AppendFormat(" SET [TID]='{0}'", NEWGUID.ToString());
                    sbSql.AppendFormat(" WHERE [TID] IS NULL AND [ID] IN ({0})", IDMOCMANULINETEMP.ToString());
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
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

        }

        public void UPDATEMOCMANULINETEMPTONULL(string NEWGUID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[MOCMANULINETEMP]");
                sbSql.AppendFormat(" SET [TID]=NULL");
                sbSql.AppendFormat(" WHERE [TID] IS NOT NULL AND [TID]='{0}'", NEWGUID.ToString());
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
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

        private void textBox742_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }


        public void INSERTTOMOCMANULINE(string ID,DateTime MANUDATE)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINE]");
                sbSql.AppendFormat(" (");
                sbSql.AppendFormat(" [ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003]");
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" SELECT [ID],[MANU],'{0}',[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003]", MANUDATE.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINETEMP]");
                sbSql.AppendFormat(" WHERE  [ID]='{0}'",ID);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
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

                    UPDATEMOCMANULINETEMPTID(ID);

                   
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

        public void UPDATEMOCMANULINETEMPTID(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


               
                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[MOCMANULINETEMP]");
                sbSql.AppendFormat(" SET [TID]='{0}' ", ID);
                sbSql.AppendFormat(" WHERE  [ID]='{0}'", ID);
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

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox17.SelectedIndex = comboBox15.SelectedIndex;
        }

        public void SEARCHMOCMANULINEMERGERESLUTMOCTA(string ID)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

          

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
     
                sbSql.AppendFormat(@"  
                                SELECT '合併' AS '合',TA001 AS '製',TA002 AS '單號',TA033 AS '批號'
                                FROM [TK].dbo.MOCTA
                                WHERE TA033 IN (
                                SELECT [NO]
                                FROM [TKMOC].[dbo].[MOCMANULINEMERGE]
                                WHERE [SID] IN (
                                SELECT ID
                                FROM [TKMOC].[dbo].[MOCMANULINE]
                                WHERE ID='{0}'
                                )
                                )
                                ",ID);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (MANU.Equals("新廠製二組"))
                {
                    if (ds1.Tables["ds1"].Rows.Count == 0)
                    {
                        dataGridView11.DataSource = null;
                    }
                    else
                    {
                        if (ds1.Tables["ds1"].Rows.Count >= 1)
                        {

                            dataGridView11.DataSource = ds1.Tables["ds1"];
                            dataGridView11.Columns[0].Width = 40;
                            dataGridView11.Columns[1].Width = 60;
                            dataGridView11.Columns[2].Width = 120;
                            dataGridView11.Columns[3].Width = 120;
                        }
                    }
                }
                else if (MANU.Equals("新廠包裝線"))
                {
                    if (ds1.Tables["ds1"].Rows.Count == 0)
                    {
                        dataGridView21.DataSource = null;
                    }
                    else
                    {
                        if (ds1.Tables["ds1"].Rows.Count >= 1)
                        {

                            dataGridView21.DataSource = ds1.Tables["ds1"];
                            dataGridView21.Columns[0].Width = 40;
                            dataGridView21.Columns[1].Width = 60;
                            dataGridView21.Columns[2].Width = 120;
                            dataGridView21.Columns[3].Width = 120;
                        }
                    }
                }
                else if (MANU.Equals("新廠製一組"))
                {
                    if (ds1.Tables["ds1"].Rows.Count == 0)
                    {
                        dataGridView22.DataSource = null;
                    }
                    else
                    {
                        if (ds1.Tables["ds1"].Rows.Count >= 1)
                        {

                            dataGridView22.DataSource = ds1.Tables["ds1"];
                            dataGridView22.Columns[0].Width = 40;
                            dataGridView22.Columns[1].Width = 60;
                            dataGridView22.Columns[2].Width = 120;
                            dataGridView22.Columns[3].Width = 120;
                        }
                    }
                }
                else if (MANU.Equals("新廠製三組(手工)"))
                {
                    if (ds1.Tables["ds1"].Rows.Count == 0)
                    {
                        dataGridView22.DataSource = null;
                    }
                    else
                    {
                        if (ds1.Tables["ds1"].Rows.Count >= 1)
                        {

                            dataGridView23.DataSource = ds1.Tables["ds1"];
                            dataGridView23.Columns[0].Width = 40;
                            dataGridView23.Columns[1].Width = 60;
                            dataGridView23.Columns[2].Width = 120;
                            dataGridView23.Columns[3].Width = 120;
                        }
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

        public void ADDMULTIMOCMANULINETEMP(string TD001,string TD002,string TD003)
        {
            if(comboBox21.Text.Equals("成品"))
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
                                        INSERT INTO [TKMOC].[dbo].[MOCMANULINETEMP]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                        SELECT ID,'{0}','{1}',TD004 [MB001],TD005 [MB002],TD006 [MB003],0 [BAR],NUM [NUM],TC053 [CLINET],0 [MANUHOUR],(NUM/MD007) [BOX],NUM [PACKAGE],TD013 [OUTDATE],TC015 [TA029],0 [HALFPRO],TD001 [COPTD001] ,TD002 [TCOPTD002], TD003 [TCOPTD003]
                                        FROM 
                                        (
                                        SELECT NEWID() AS ID,TD001,TD002,TD003,TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015,TD013,(CASE WHEN ISNULL(MD002,'')<>'' THEN (TD008+TD024)*MD004 ELSE (TD008+TD024)  END ) AS NUM
                                        ,ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE (MD.MD003 LIKE '201%') AND MD.MD001=TD004),1) AS MD007
                                        ,ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=TD004),1) AS MC004
                                        FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                        LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002
                                        WHERE TC001=TD001 AND TC002=TD002
                                        AND MB001=TD004
                                        AND TD001='{2}' AND TD002='{3}' AND TD003='{4}'
                                        ) AS TEMP 
                                        ", comboBox19.Text.Trim(),dateTimePicker23.Value.ToString("yyyyMMdd"),TD001,TD002,TD003);

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

                        MessageBox.Show("成功");
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
            else if (comboBox21.Text.Equals("第一層"))
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
                                        INSERT INTO [TKMOC].[dbo].[MOCMANULINETEMP]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                        SELECT ID,'{0}','{1}',MD003 [MB001],MD035 [MB002],MD036 [MB003],0 [BAR],BOMNUMS [NUM],TC053 [CLINET],0 [MANUHOUR],CONVERT(DECIMAL(16,4),(BOMNUMS/MD007B)) [BOX],BOMNUMS [PACKAGE],TD013 [OUTDATE],TC015 [TA029],0 [HALFPRO],TD001 [COPTD001] ,TD002 [TCOPTD002], TD003 [TCOPTD003]
                                        FROM 
                                        (
                                        SELECT  NEWID() AS ID,TD001,TD002,TD003,TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015,TD013,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM
                                        ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,BOMMD.MD006,BOMMD.MD007
                                        ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END )*BOMMD.MD006/BOMMD.MD007) AS BOMNUMS
                                        ,ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE (MD.MD003 LIKE '201%') AND MD.MD001=BOMMD.MD003),1) AS MD007B
                                        ,ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=BOMMD.MD003),1) AS MC004
                                        FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                        LEFT JOIN [TK].dbo.INVMD ON INVMD.MD001=TD004 AND TD010=MD002
                                        LEFT JOIN [TK].dbo.BOMMD ON BOMMD.MD001=TD004
                                        WHERE TC001=TD001 AND TC002=TD002
                                        AND MB001=TD004
                                        AND TD001='{2}' AND TD002='{3}' AND TD003='{4}'
                                        AND (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')
                                        ) AS TEMP
                                        ", comboBox19.Text.Trim(), dateTimePicker23.Value.ToString("yyyyMMdd"), TD001, TD002, TD003);

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

                        MessageBox.Show("成功");
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
            else if (comboBox21.Text.Equals("第二層"))
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
                                        INSERT INTO [TKMOC].[dbo].[MOCMANULINETEMP]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                        SELECT ID,'{0}','{1}',MD003B [MB001],MD035B [MB002],MD036B [MB003],CONVERT(DECIMAL(16,4),(BOMNUMS2/MC004C)) [BAR],BOMNUMS2 [NUM],TC053 [CLINET],0 [MANUHOUR],0 [BOX],0 [PACKAGE],TD013 [OUTDATE],TC015 [TA029],0 [HALFPRO],TD001 [COPTD001] ,TD002 [TCOPTD002], TD003 [TCOPTD003]
                                        FROM (
                                        SELECT NEWID() AS ID,TD001,TD002,TD003,TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015,TD013,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM
                                        ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,BOMMD.MD006,BOMMD.MD007
                                        ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END )*BOMMD.MD006/BOMMD.MD007) AS BOMNUMS
                                        ,BOMMD2.MD003 MD003B,BOMMD2.MD035 MD035B,BOMMD2.MD036 MD036B,BOMMD2.MD006 MD006B,BOMMD2.MD007 MD007B
                                        ,(((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END )*BOMMD.MD006/BOMMD.MD007)*BOMMD2.MD006/BOMMD2.MD007)AS BOMNUMS2
                                        ,ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE (MD.MD003 LIKE '201%') AND MD.MD001=BOMMD2.MD003),1) AS MD007C
                                        ,ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=BOMMD2.MD003),1) AS MC004C
                                        FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                        LEFT JOIN [TK].dbo.INVMD ON INVMD.MD001=TD004 AND TD010=MD002
                                        LEFT JOIN [TK].dbo.BOMMD ON BOMMD.MD001=TD004
                                        LEFT JOIN [TK].dbo.BOMMD BOMMD2 ON BOMMD2.MD001=BOMMD.MD003
                                        WHERE TC001=TD001 AND TC002=TD002
                                        AND MB001=TD004
                                        AND TD001='{2}' AND TD002='{3}' AND TD003='{4}'
                                        AND (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')
                                        AND (BOMMD2.MD003 LIKE '3%' OR BOMMD2.MD003 LIKE '4%')
                                        ) AS TEMP
                                        ", comboBox19.Text.Trim(), dateTimePicker23.Value.ToString("yyyyMMdd"), TD001, TD002, TD003);

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

                        MessageBox.Show("成功");
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                ADDMOCMANULINE();
                SETNULL2();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE();
             
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETNULL2();
            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox1.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
  
        }


        private void button5_Click(object sender, EventArgs e)
        {

            TA002 = GETMAXTA002(TA001);
            ADDMOCMANULINERESULT();
            ADDMOCTATB();
            SEARCHMOCMANULINERESULT();

            MessageBox.Show("完成");
        }
        private void button6_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            SETNULL3();

            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox7.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox7.Text))
            {
                ADDMOCMANULINE();
                SETNULL3();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(TA028))
            {
                TA002 = GETMAXTA002(TA001);
                ADDMOCMANULINERESULT();
                ADDMOCTATB();
                SEARCHMOCMANULINERESULT();

                MessageBox.Show("完成");
            }
            else
            {
                MessageBox.Show("訂單沒有指定");
            }
            
        }
        private void button11_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }
        private void button12_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox14.Text))
            {
                ADDMOCMANULINE();
                SETNULL4();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            SETNULL4();
            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox14.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(TA028B))
            {
                TA002 = GETMAXTA002(TA001);
                ADDMOCMANULINERESULT();
                ADDMOCTATB();
                SEARCHMOCMANULINERESULT();

                MessageBox.Show("完成");
            }
            else
            {
                MessageBox.Show("訂單沒有指定");
            }
           
        }
        private void button16_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }
        private void button17_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox20.Text))
            {
                ADDMOCMANULINE();
                SETNULL6();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
        }
        private void button19_Click(object sender, EventArgs e)
        {
            SETNULL6();
            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox20.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {

            TA002 = GETMAXTA002(TA001);
            ADDMOCMANULINERESULT();
            ADDMOCTATB();
            SEARCHMOCMANULINERESULT();

            MessageBox.Show("完成");
        }


        private void button21_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINERESULT();
                SEARCHMOCMANULINE();

                DELMOCMANULINECOP();
                
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINERESULT();
                SEARCHMOCMANULINE();

                DELMOCMANULINECOP();
               
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINERESULT();
                SEARCHMOCMANULINE();

                DELMOCMANULINECOP();
                
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINERESULT();
                SEARCHMOCMANULINE();

                DELMOCMANULINECOP();
                
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }


        private void button25_Click(object sender, EventArgs e)
        {
            SEARCHMOCTB();
        }



        private void button26_Click(object sender, EventArgs e)
        {
            TA002 = GETMAXTA002(TA001);
            ADDMOCMANULINETOATL();
            ADDMOCTATB();
            SEARCHMOCMANULINETOATL();

            MessageBox.Show("完成");
        }

        private void button27_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINETOATL();
                SEARCHMOCMANULINETOATL();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox6.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }
        

        private void button29_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox9.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }

        private void button30_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox16.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }

        private void button31_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox22.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            SEARCHMOCMANULINE();
        }
        private void button35_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            SEARCHMOCMANULINE();
        }

        private void button34_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            SEARCHMOCMANULINE();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            SEARCHMOCMANULINE();
        }

       

      

       

        private void button40_Click(object sender, EventArgs e)
        {
            ADDCALENDAR();
            SETCALENDAR();
        }

        private void button41_Click(object sender, EventArgs e)
        {
            DELCALENDAR();
            SETCALENDAR();
        }
        private void button42_Click(object sender, EventArgs e)
        {
            SEARCHCOPTD();
        }
        private void button43_Click(object sender, EventArgs e)
        {
            UPDATECOPTD();
            SEARCHCOPTD();
        }

        private void button44_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }

        private void button46_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox56.Text))
            {
                ADDMOCMANULINE();
                SETNULL7();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
        }

        private void button49_Click(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button48_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            SEARCHMOCMANULINE();
        }

        private void button47_Click(object sender, EventArgs e)
        {
            SETNULL7();

            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox56.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
        }

        private void button45_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox57.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }

        private void button51_Click(object sender, EventArgs e)
        {
            TA002 = GETMAXTA002(TA001);
            ADDMOCMANULINERESULT();
            ADDMOCTATB();

            SEARCHMOCMANULINERESULT();

            MessageBox.Show("完成");
        }

        private void button50_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINERESULT();
                SEARCHMOCMANULINE();

                DELMOCMANULINECOP();
               
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

     

        private void button36_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox40.Text) & !string.IsNullOrEmpty(textBox41.Text) & !string.IsNullOrEmpty(textBox73.Text))
            {
                SEARCHCOPDEFAULT(textBox40.Text, textBox41.Text, textBox73.Text);
            }


        }
        private void button37_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox42.Text) & !string.IsNullOrEmpty(textBox43.Text) & !string.IsNullOrEmpty(textBox72.Text))
            {
                SEARCHCOPDEFAULT(textBox42.Text, textBox43.Text, textBox72.Text);
            }
        }

        private void button38_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox44.Text) & !string.IsNullOrEmpty(textBox45.Text) & !string.IsNullOrEmpty(textBox74.Text))
            {
                SEARCHCOPDEFAULT(textBox44.Text, textBox45.Text, textBox74.Text);
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox46.Text) & !string.IsNullOrEmpty(textBox47.Text) & !string.IsNullOrEmpty(textBox75.Text))
            {
                SEARCHCOPDEFAULT(textBox46.Text, textBox47.Text, textBox75.Text);
            }
        }

        private void button52_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox40.Text) & !string.IsNullOrEmpty(textBox41.Text) & !string.IsNullOrEmpty(textBox73.Text))
            {
                SEARCHCOPDEFAULT2(textBox40.Text, textBox41.Text, textBox73.Text);
            }

        }

        private void button53_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox44.Text) & !string.IsNullOrEmpty(textBox45.Text) & !string.IsNullOrEmpty(textBox74.Text))
            {
                SEARCHCOPDEFAULT2(textBox44.Text, textBox45.Text, textBox74.Text);
            }
        }

        private void button54_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox46.Text) & !string.IsNullOrEmpty(textBox47.Text) & !string.IsNullOrEmpty(textBox75.Text))
            {
                SEARCHCOPDEFAULT2(textBox46.Text, textBox47.Text, textBox75.Text);
            }
        }


        private void button55_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox42.Text) & !string.IsNullOrEmpty(textBox43.Text) & !string.IsNullOrEmpty(textBox72.Text))
            {
                SEARCHCOPDEFAULT2(textBox42.Text, textBox43.Text, textBox72.Text);
            }
        }
        private void button56_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox42.Text) & !string.IsNullOrEmpty(textBox43.Text) & !string.IsNullOrEmpty(textBox72.Text))
            {
                SEARCHCOPDEFAULT3(textBox42.Text, textBox43.Text, textBox72.Text);
            }
        }

        private void button57_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox40.Text) & !string.IsNullOrEmpty(textBox41.Text) & !string.IsNullOrEmpty(textBox73.Text))
            {
                SEARCHCOPDEFAULT3(textBox40.Text, textBox41.Text, textBox73.Text);
            }
        }

        private void button58_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox44.Text) & !string.IsNullOrEmpty(textBox45.Text) & !string.IsNullOrEmpty(textBox74.Text))
            {
                SEARCHCOPDEFAULT3(textBox44.Text, textBox45.Text, textBox74.Text);
            }
        }

        private void button59_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox46.Text) & !string.IsNullOrEmpty(textBox47.Text) & !string.IsNullOrEmpty(textBox75.Text))
            {
                SEARCHCOPDEFAULT3(textBox46.Text, textBox47.Text, textBox75.Text);
            }
        }
        private void button61_Click(object sender, EventArgs e)
        {
            
            if (!string.IsNullOrEmpty(ID1) & !string.IsNullOrEmpty(textBox40.Text) & !string.IsNullOrEmpty(textBox41.Text) & !string.IsNullOrEmpty(textBox73.Text) & !string.IsNullOrEmpty(textBox77.Text))
            {
                INSERTMOCMANULINECOP(ID1,textBox40.Text, textBox41.Text, textBox73.Text, textBox77.Text);

                SEARCHMOCMANULINECOP(ID1);
            }
                
        }

        private void button62_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(DELMOCMANULINECOPID))
                {
                    DELMOCMANULINECOP(DELMOCMANULINECOPID);

                    SEARCHMOCMANULINECOP(ID1);
                }
            }

        }

        private void button63_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE12(comboBox15.Text.Trim(),dateTimePicker20.Value.ToString("yyyyMMdd"), dateTimePicker21.Value.ToString("yyyyMMdd"),comboBox18.Text.Trim());
        }
        private void button64_Click(object sender, EventArgs e)
        {
            INSERTMOCMANULINEMERGE(dateTimePicker22.Value);
            SEARCHMOCMANULINEMERGE(dateTimePicker22.Value);
        }


        private void button65_Click(object sender, EventArgs e)
        {
            CALSUMMOCMANULINEMERGE(textBox78.Text.Trim());

               
        }

        private void button66_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox78.Text))
            {
                TA001 = "A510";

                TA002 = GETMAXTA002MERGE(dateTimePicker22.Value,TA001);

                if(comboBox17.Text.Equals("新廠包裝線"))
                {
                    //TA015=textBox84
                    ADDMOCTATBMERGE(TA001, TA002, textBox80.Text, textBox81.Text, textBox82.Text, label104.Text, textBox79.Text, comboBox17.SelectedValue.ToString().Trim(), textBox84.Text.Trim(), textBox78.Text.Trim());
                }
                else
                {
                    //TA015=textBox83
                    ADDMOCTATBMERGE(TA001, TA002, textBox80.Text, textBox81.Text, textBox82.Text, label104.Text, textBox79.Text, comboBox17.SelectedValue.ToString().Trim(), textBox83.Text.Trim(), textBox78.Text.Trim());
                }
                
                //SEARCHMOCMANULINERESULT();

                MessageBox.Show(TA001+" "+ TA002+"完成");
            }
            else
            {
                MessageBox.Show("訂單沒有指定");
            }
            


        }

        private void button68_Click(object sender, EventArgs e)
        {
            DATAGRIDCLEAR();
        }
        private void button76_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINETEMP(comboBox20.Text.Trim(),textBox722.Text.Trim());
        }

        private void button75_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox781.Text) & !string.IsNullOrEmpty(textBox782.Text) & !string.IsNullOrEmpty(textBox783.Text))
            {
                SEARCHCOPDEFAULT(textBox781.Text, textBox782.Text, textBox783.Text);
            }
        }

        private void button73_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox781.Text) & !string.IsNullOrEmpty(textBox782.Text) & !string.IsNullOrEmpty(textBox783.Text))
            {
                SEARCHCOPDEFAULT2(textBox781.Text, textBox782.Text, textBox783.Text);
            }
        }

        private void button74_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox781.Text) & !string.IsNullOrEmpty(textBox782.Text) & !string.IsNullOrEmpty(textBox783.Text))
            {
                SEARCHCOPDEFAULT3(textBox781.Text, textBox782.Text, textBox783.Text);
            }
        }
        private void button69_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox731.Text))
            {
                ADDMOCMANULINE();
                SEARCHMOCMANULINETEMP(comboBox20.Text.Trim(), textBox722.Text.Trim());
                SETNULL8();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
        }

        private void button71_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE();
                SEARCHMOCMANULINETEMP(comboBox20.Text.Trim(), textBox722.Text.Trim());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button72_Click(object sender, EventArgs e)
        {
            textBox722.Text = LIMITSERCHTD002;
            CHECKMOCTAB();
            SEARCHMOCMANULINETEMP(comboBox20.Text.Trim(), textBox722.Text.Trim());
        }


        private void button77_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in this.dataGridView20.Rows)
            {
                ID10 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[18].ToString();
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    INSERTTOMOCMANULINE(ID10, dateTimePicker25.Value);
                }
            }

            MessageBox.Show("移轉成功");

            SEARCHMOCMANULINETEMP(comboBox20.Text.Trim(), textBox722.Text.Trim());

        }
        private void button78_Click(object sender, EventArgs e)
        {
            frmMOCMANULINESubTEMPADDBACTH SUBfrmMOCMANULINESubTEMPADDBACTH = new frmMOCMANULINESubTEMPADDBACTH();
            SUBfrmMOCMANULINESubTEMPADDBACTH.ShowDialog();

            SEARCHMOCMANULINETEMP(comboBox20.Text.Trim(), textBox722.Text.Trim());
        }



        private void button79_Click(object sender, EventArgs e)
        {
            ADDMULTIMOCMANULINETEMP(textBox781.Text.Trim(), textBox782.Text.Trim(), textBox783.Text.Trim());
        }


        #endregion

       
    }
}
