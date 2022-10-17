using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MySql.Data.MySqlClient;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace DBAdminTool
{
    public partial class Form1 : Form
    {
        public DateTime UploadedReportDateTime;
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkBook;
        private Excel.Worksheet xlWorkSheet, xlWorkSheet2, xlWorkSheet3,xlWorkSheet4,xlWorkSheet5;
        private Excel.Range range, range2, range3,range4,range5;
        private string filepath;
        private int check_exist;
        private DateTime datetimelocal;
        public Excel.Application XlApp
        {
            get
            {
                return xlApp;
            }

            set
            {
                xlApp = value;
            }
        }
        public Excel.Workbook XlWorkBook
        {
            get
            {
                return xlWorkBook;
            }

            set
            {
                xlWorkBook = value;
            }
        }
        public Excel.Worksheet XlWorkSheet2
        {
            get
            {
                return xlWorkSheet2;
            }

            set
            {
                xlWorkSheet2 = value;
            }
        }
        public Excel.Worksheet XlWorkSheet
        {
            get
            {
                return xlWorkSheet;
            }

            set
            {
                xlWorkSheet = value;
            }
        }
        public Excel.Range Range
        {
            get
            {
                return range;
            }

            set
            {
                range = value;
            }
        }
        public Excel.Range Range2
        {
            get
            {
                return range2;
            }

            set
            {
                range2 = value;
            }
        }
        public string Filepath
        {
            get
            {
                return filepath;
            }

            set
            {
                filepath = value;
            }
        }
        public int Check_exist
        {
            get
            {
                return check_exist;
            }

            set
            {
                check_exist = value;
            }
        }
        public Excel.Worksheet XlWorkSheet3
        {
            get
            {
                return xlWorkSheet3;
            }

            set
            {
                xlWorkSheet3 = value;
            }
        }
        public Excel.Worksheet XlWorkSheet4
        {
            get
            {
                return xlWorkSheet4;
            }

            set
            {
                xlWorkSheet4 = value;
            }
        }
        public Excel.Worksheet XlWorkSheet5
        {
            get
            {
                return xlWorkSheet5;
            }

            set
            {
                xlWorkSheet5 = value;
            }
        }
        public Excel.Range Range3
        {
            get
            {
                return range3;
            }

            set
            {
                range3 = value;
            }
        }
        public Excel.Range Range4
        {
            get
            {
                return range4;
            }

            set
            {
                range4 = value;
            }
        }
        public Excel.Range Range5
        {
            get
            {
                return range5;
            }

            set
            {
                range5 = value;
            }
        }
        public Form1()
        {
            InitializeComponent();
            dateTimePicker1.MinDate = new DateTime(2015, 10, 01);
            dateTimePicker1.MaxDate = DateTime.Today;
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            checkExistance();
            if(Check_exist==1)
            {
                MessageBox.Show("This Date exist already in database");
            }
            else
            {
                insert();
            }
            
        }      
        private void button2_Click(object sender, EventArgs e)
        {
            string localdatebackup = DateTime.Today.Day.ToString() + DateTime.Today.Month.ToString() + DateTime.Today.Year.ToString();
            string conString = "SERVER=10.66.8.137;DATABASE=wind;UID=root;Password=ZAQ!2wsx;";
            string file = "C:\\Users\\DiamantisK\\Desktop\\" + localdatebackup + "backup.sql";
            MySqlConnection conn = new MySqlConnection(conString);
            MySqlCommand cmd = new MySqlCommand();
            MySqlBackup mb = new MySqlBackup(cmd);
            cmd.Connection = conn;
            conn.Open();
            mb.ExportToFile(file);
            conn.Close();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {  
                    Filepath= openFileDialog1.InitialDirectory + openFileDialog1.FileName;
                    textBox1.Text =openFileDialog1.InitialDirectory + openFileDialog1.FileName;  
            }
        }
        public void checkExistance()
        {
          
            try
            {
                MySqlConnection connection = getConnection();
                MySqlConnection connection1 = getConnection();

                //SQL query assignment
                MySqlCommand mycm = new MySqlCommand("", connection);
                connection.Open();


                connection1.Open();
                mycm.Prepare();
                mycm.CommandText = String.Format("select * FROM date WHERE DateOfReport=?dateofday");
                mycm.Parameters.AddWithValue("?dateofday", getDaytime());
                try
                {
                    MySqlDataReader msdr = mycm.ExecuteReader();
                    if(msdr.HasRows)
                    {
                        Check_exist = 1;
                    }
                    else
                    {
                        Check_exist = 0;
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
               
                mycm.Parameters.Clear();
                mycm.Cancel();
                mycm.Dispose();
                connection1.Close();
           
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            
        }
        private void button7_Click(object sender, EventArgs e)
        {
            deleteTableWithSelectedDate("date");

            deleteTableWithSelectedDate("prefecture_report");
            deleteTableWithSelectedDate("availability");
            deleteTableWithSelectedDate("static");
            deleteTableWithSelectedDate("operational_2g");
            deleteTableWithSelectedDate("operational_3g");
            deleteTableWithSelectedDate("operational_4g");
            deleteTableWithSelectedDate("retention_2g");
            deleteTableWithSelectedDate("retention_3g");
            deleteTableWithSelectedDate("retention_4g");
            deleteTableWithSelectedDate("total_operational_2g");
            deleteTableWithSelectedDate("total_operational_3g");
            deleteTableWithSelectedDate("total_operational_4g");
            deleteTableWithSelectedDate("total_retention_2g");
            deleteTableWithSelectedDate("total_retention_3g");
            deleteTableWithSelectedDate("total_retention_4g");
            deleteTableWithSelectedDate("operational_affected");
            deleteTableWithSelectedDate("retention_affected");
            deleteTableWithSelectedDate("licensing_affected");
            MessageBox.Show("Deletion completed ");
        }
        private void deleteTableWithSelectedDate(string tablename)
        {
            try
            {
                MySqlConnection conn = getConnection();
                conn.Open();
                MySqlCommand mycm = new MySqlCommand("", conn);
                
                mycm.Prepare();
                mycm.CommandText = String.Format("DELETE FROM " + tablename + " WHERE DateOfReport=?datep");
                mycm.Parameters.AddWithValue("?datep", getDaytime());
                MySqlDataReader msdr = mycm.ExecuteReader();
                while (msdr.Read())
                {

                }

                conn.Close();
                mycm.Parameters.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(tablename);
                Console.WriteLine(ex.ToString());
            }

        }
        private DateTime getDaytime()
        {
            return (DateTime)dateTimePicker1.Value.Date;
        }
        public MySqlConnection getConnection()
        {
            //open a connection to the localhost mysql database that I created with the Create Table statement
            string myConnectionString = "SERVER=10.66.8.137;DATABASE=wind;UID=root;Password=ZAQ!2wsx;";
            MySqlConnection conn = new MySqlConnection(myConnectionString);
            return conn;
        }
        public void insert()
        {
            //string theDate = DateTime.Today.Day.ToString() + DateTime.Today.Month.ToString() + DateTime.Today.Year.ToString();
            DateTime localdate = getDaytime();
            //MessageBox.Show(theDate);

            MySqlConnection connection = getConnection();
            MySqlConnection connection1 = getConnection();
            MySqlConnection connection2 = getConnection();
            MySqlConnection connection3 = getConnection();

            //open connection to excel file
            XlApp = new Excel.Application();
            XlWorkBook = XlApp.Workbooks.Open(Filepath, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            XlWorkSheet = (Excel.Worksheet)XlWorkBook.Worksheets["Data Table -1"];
            XlWorkSheet2 = (Excel.Worksheet)XlWorkBook.Worksheets["Data Table -2"];
            XlWorkSheet3 = (Excel.Worksheet)XlWorkBook.Worksheets["Operational"];
            XlWorkSheet4 = (Excel.Worksheet)XlWorkBook.Worksheets["Retention-Deployment"];
            XlWorkSheet5 = (Excel.Worksheet)XlWorkBook.Worksheets["Licensing"];

            Range = XlWorkSheet.UsedRange;
            Range2 = XlWorkSheet2.UsedRange;
            Range3 = XlWorkSheet3.UsedRange;
            Range4 = XlWorkSheet4.UsedRange;
            Range5 = XlWorkSheet5.UsedRange;


            try
            {
                //Open Connection
                connection.Open();
                connection1.Open();
                connection2.Open();
                connection3.Open();

                //SQL query assignment
                MySqlCommand mycm = new MySqlCommand("", connection);
                MySqlCommand mycm1 = new MySqlCommand("", connection);
                MySqlCommand mycm2 = new MySqlCommand("", connection);
                MySqlCommand mycm3 = new MySqlCommand("", connection);
                MySqlCommand mycm4 = new MySqlCommand("", connection);
                MySqlCommand mycm5 = new MySqlCommand("", connection);
                MySqlCommand mycm6 = new MySqlCommand("", connection);
                MySqlCommand mycm7 = new MySqlCommand("", connection);
                MySqlCommand mycm8 = new MySqlCommand("", connection);
                MySqlCommand mycm9 = new MySqlCommand("", connection);
                MySqlCommand mycm10 = new MySqlCommand("", connection);
                MySqlCommand myquery = new MySqlCommand("", connection2);
                MySqlCommand mycmt2go = new MySqlCommand("", connection1);
                MySqlCommand mycmt3go = new MySqlCommand("", connection1);
                MySqlCommand mycmt4go = new MySqlCommand("", connection1);
                MySqlCommand mycmt2gr = new MySqlCommand("", connection1);
                MySqlCommand mycmt3gr = new MySqlCommand("", connection1);
                MySqlCommand mycmt4gr = new MySqlCommand("", connection1);
                MySqlCommand mycmreasonsope = new MySqlCommand("", connection3);
                MySqlCommand mycmreasonsret = new MySqlCommand("", connection3);
                MySqlCommand mycmreasonslic = new MySqlCommand("", connection3);

                //add the specified values to the book in order to store it to the database
                mycm.Prepare();
                mycm.CommandText = String.Format("insert into date(DateOfReport) values (?date_para)");
                mycm.Parameters.AddWithValue("?date_para", localdate);
                mycm.ExecuteNonQuery();
                mycm.Parameters.Clear();


                //availability
                mycm1.Prepare();
                mycm1.CommandText = String.Format("insert into availability(DateOfReport,Unavailable2G,Unavailable3G,Unavailable4G,Unavailable2GOperational,Unavailable2GRetention,Unavailable2GLicensing,Unavailable3GOperational,Unavailable3GRetention,Unavailable3GLicensing,Unavailable4GOperational,Unavailable4GRetention,Unavailable4GLicensing) values (?date1_para,?g2,?g3,?g4,?g2o,?g2r,?g2l,?g3o,?g3r,?g3l,?g4o,?g4r,?g4l)");
                mycm1.Parameters.AddWithValue("?date1_para", localdate);
                mycm1.Parameters.AddWithValue("?g2", int.Parse((Range2.Cells[6, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g3", int.Parse((Range2.Cells[7, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g4", int.Parse((Range2.Cells[8, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g2o", int.Parse((Range2.Cells[9, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g3o", int.Parse((Range2.Cells[10, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g4o", int.Parse((Range2.Cells[11, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g2r", int.Parse((Range2.Cells[12, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g3r", int.Parse((Range2.Cells[13, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g4r", int.Parse((Range2.Cells[14, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g2l", int.Parse((Range2.Cells[15, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g3l", int.Parse((Range2.Cells[16, 3] as Excel.Range).Text));
                mycm1.Parameters.AddWithValue("?g4l", int.Parse((Range2.Cells[17, 3] as Excel.Range).Text));
                mycm1.ExecuteNonQuery();
                mycm1.Parameters.Clear();

                mycm2.Prepare();
                mycm2.CommandText = String.Format("insert into static(DateOfReport,Available2G,Available3G,Available4G) values (?date2_para,?ag2,?ag3,?ag4)");
                mycm2.Parameters.AddWithValue("?date2_para", localdate);
                mycm2.Parameters.AddWithValue("?ag2", int.Parse((Range2.Cells[3, 3] as Excel.Range).Text));
                mycm2.Parameters.AddWithValue("?ag3", int.Parse((Range2.Cells[4, 3] as Excel.Range).Text));
                mycm2.Parameters.AddWithValue("?ag4", int.Parse((Range2.Cells[5, 3] as Excel.Range).Text));
                mycm2.ExecuteNonQuery();
                mycm2.Parameters.Clear();

                //Prepare total Operational Cards
                mycmt2go.Prepare();
                mycmt2go.CommandText = String.Format("insert into total_operational_2G(DateOfReport,Antenna,CosmotePowerProblem,Disinfection,FiberCut,GeneratorFailure,Link,LinkDueToPowerProblem,OTEProblem,PowerProblem,PPCPowerFailure,Quality,RBSProblem,Temperature,VodafoneLinkProblem,VodafonePowerProblem,Modem) values (?date4_para,?p1,?p2,?p3,?p4,?p5,?p6,?p7,?p8,?p9,?p10,?p11,?p12,?p13,?p14,?p15,?p16)");
                mycmt2go.Parameters.AddWithValue("?date4_para", localdate);
                mycmt2go.Parameters.AddWithValue("?p1", int.Parse((Range2.Cells[72, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p2", int.Parse((Range2.Cells[73, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p3", int.Parse((Range2.Cells[74, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p4", int.Parse((Range2.Cells[75, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p5", int.Parse((Range2.Cells[76, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p6", int.Parse((Range2.Cells[77, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p7", int.Parse((Range2.Cells[78, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p8", int.Parse((Range2.Cells[79, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p9", int.Parse((Range2.Cells[80, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p10", int.Parse((Range2.Cells[81, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p11", int.Parse((Range2.Cells[82, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p12", int.Parse((Range2.Cells[83, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p13", int.Parse((Range2.Cells[84, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p14", int.Parse((Range2.Cells[85, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p15", int.Parse((Range2.Cells[86, 3] as Excel.Range).Text));
                mycmt2go.Parameters.AddWithValue("?p16", int.Parse((Range2.Cells[87, 3] as Excel.Range).Text));
                mycmt2go.ExecuteNonQuery();
                mycmt2go.Parameters.Clear();

                mycmt3go.Prepare();
                mycmt3go.CommandText = String.Format("insert into total_operational_3G(DateOfReport,Antenna,CosmotePowerProblem,Disinfection,FiberCut,GeneratorFailure,Link,LinkDueToPowerProblem,OTEProblem,PowerProblem,PPCPowerFailure,Quality,RBSProblem,Temperature,VodafoneLinkProblem,VodafonePowerProblem,Modem) values (?date4_para,?p1,?p2,?p3,?p4,?p5,?p6,?p7,?p8,?p9,?p10,?p11,?p12,?p13,?p14,?p15,?p16)");
                mycmt3go.Parameters.AddWithValue("?date4_para", localdate);
                mycmt3go.Parameters.AddWithValue("?p1", int.Parse((Range2.Cells[91, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p2", int.Parse((Range2.Cells[92, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p3", int.Parse((Range2.Cells[93, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p4", int.Parse((Range2.Cells[94, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p5", int.Parse((Range2.Cells[95, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p6", int.Parse((Range2.Cells[96, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p7", int.Parse((Range2.Cells[97, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p8", int.Parse((Range2.Cells[98, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p9", int.Parse((Range2.Cells[99, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p10", int.Parse((Range2.Cells[100, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p11", int.Parse((Range2.Cells[101, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p12", int.Parse((Range2.Cells[102, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p13", int.Parse((Range2.Cells[103, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p14", int.Parse((Range2.Cells[104, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p15", int.Parse((Range2.Cells[105, 3] as Excel.Range).Text));
                mycmt3go.Parameters.AddWithValue("?p16", int.Parse((Range2.Cells[106, 3] as Excel.Range).Text));
                mycmt3go.ExecuteNonQuery();
                mycmt3go.Parameters.Clear();

                mycmt4go.Prepare();
                mycmt4go.CommandText = String.Format("insert into total_operational_4G(DateOfReport,Antenna,CosmotePowerProblem,Disinfection,FiberCut,GeneratorFailure,Link,LinkDueToPowerProblem,OTEProblem,PowerProblem,PPCPowerFailure,Quality,RBSProblem,Temperature,VodafoneLinkProblem,VodafonePowerProblem,Modem) values (?date4_para,?p1,?p2,?p3,?p4,?p5,?p6,?p7,?p8,?p9,?p10,?p11,?p12,?p13,?p14,?p15,?p16)");
                mycmt4go.Parameters.AddWithValue("?date4_para", localdate);
                mycmt4go.Parameters.AddWithValue("?p1", int.Parse((Range2.Cells[110, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p2", int.Parse((Range2.Cells[111, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p3", int.Parse((Range2.Cells[112, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p4", int.Parse((Range2.Cells[113, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p5", int.Parse((Range2.Cells[114, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p6", int.Parse((Range2.Cells[115, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p7", int.Parse((Range2.Cells[116, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p8", int.Parse((Range2.Cells[117, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p9", int.Parse((Range2.Cells[118, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p10", int.Parse((Range2.Cells[119, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p11", int.Parse((Range2.Cells[120, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p12", int.Parse((Range2.Cells[121, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p13", int.Parse((Range2.Cells[122, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p14", int.Parse((Range2.Cells[123, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p15", int.Parse((Range2.Cells[124, 3] as Excel.Range).Text));
                mycmt4go.Parameters.AddWithValue("?p16", int.Parse((Range2.Cells[125, 3] as Excel.Range).Text));
                mycmt4go.ExecuteNonQuery();
                mycmt4go.Parameters.Clear();

                //prepare retention total

                mycmt2gr.Prepare();
                mycmt2gr.CommandText = String.Format("insert into total_retention_2G(DateOfReport,Access,Antenna,Cabinet,DisasterDueToFire,DisasterDueToFlood,OwnerReaction,PeopleReaction,PPCIntention,Renovation,Shelter,Thievery,UnpaidBill,Vandalism,Reengineering) values (?date7_para,?f1,?f2,?f3,?f4,?f5,?f6,?f7,?f8,?f9,?f10,?f11,?f12,?f13,?f14)");
                mycmt2gr.Parameters.AddWithValue("?date7_para", localdate);
                mycmt2gr.Parameters.AddWithValue("?f1", int.Parse((Range2.Cells[21, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f2", int.Parse((Range2.Cells[22, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f3", int.Parse((Range2.Cells[23, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f4", int.Parse((Range2.Cells[24, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f5", int.Parse((Range2.Cells[25, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f6", int.Parse((Range2.Cells[26, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f7", int.Parse((Range2.Cells[27, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f8", int.Parse((Range2.Cells[28, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f9", int.Parse((Range2.Cells[29, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f10", int.Parse((Range2.Cells[30, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f11", int.Parse((Range2.Cells[31, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f12", int.Parse((Range2.Cells[32, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f13", int.Parse((Range2.Cells[33, 3] as Excel.Range).Text));
                mycmt2gr.Parameters.AddWithValue("?f14", int.Parse((Range2.Cells[34, 3] as Excel.Range).Text));
                mycmt2gr.ExecuteNonQuery();
                mycmt2gr.Parameters.Clear();

                mycmt3gr.Prepare();
                mycmt3gr.CommandText = String.Format("insert into total_retention_3G(DateOfReport,Access,Antenna,Cabinet,DisasterDueToFire,DisasterDueToFlood,OwnerReaction,PeopleReaction,PPCIntention,Renovation,Shelter,Thievery,UnpaidBill,Vandalism,Reengineering) values (?date7_para,?f1,?f2,?f3,?f4,?f5,?f6,?f7,?f8,?f9,?f10,?f11,?f12,?f13,?f14)");
                mycmt3gr.Parameters.AddWithValue("?date7_para", localdate);
                mycmt3gr.Parameters.AddWithValue("?f1", int.Parse((Range2.Cells[38, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f2", int.Parse((Range2.Cells[39, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f3", int.Parse((Range2.Cells[40, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f4", int.Parse((Range2.Cells[41, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f5", int.Parse((Range2.Cells[42, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f6", int.Parse((Range2.Cells[43, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f7", int.Parse((Range2.Cells[44, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f8", int.Parse((Range2.Cells[45, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f9", int.Parse((Range2.Cells[46, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f10", int.Parse((Range2.Cells[47, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f11", int.Parse((Range2.Cells[48, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f12", int.Parse((Range2.Cells[49, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f13", int.Parse((Range2.Cells[50, 3] as Excel.Range).Text));
                mycmt3gr.Parameters.AddWithValue("?f14", int.Parse((Range2.Cells[51, 3] as Excel.Range).Text));
                mycmt3gr.ExecuteNonQuery();
                mycmt3gr.Parameters.Clear();

                mycmt4gr.Prepare();
                mycmt4gr.CommandText = String.Format("insert into total_retention_4G(DateOfReport,Access,Antenna,Cabinet,DisasterDueToFire,DisasterDueToFlood,OwnerReaction,PeopleReaction,PPCIntention,Renovation,Shelter,Thievery,UnpaidBill,Vandalism,Reengineering) values (?date7_para,?f1,?f2,?f3,?f4,?f5,?f6,?f7,?f8,?f9,?f10,?f11,?f12,?f13,?f14)");
                mycmt4gr.Parameters.AddWithValue("?date7_para", localdate);
                mycmt4gr.Parameters.AddWithValue("?f1", int.Parse((Range2.Cells[55, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f2", int.Parse((Range2.Cells[56, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f3", int.Parse((Range2.Cells[57, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f4", int.Parse((Range2.Cells[58, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f5", int.Parse((Range2.Cells[59, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f6", int.Parse((Range2.Cells[60, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f7", int.Parse((Range2.Cells[61, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f8", int.Parse((Range2.Cells[62, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f9", int.Parse((Range2.Cells[63, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f10", int.Parse((Range2.Cells[64, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f11", int.Parse((Range2.Cells[65, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f12", int.Parse((Range2.Cells[66, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f13", int.Parse((Range2.Cells[67, 3] as Excel.Range).Text));
                mycmt4gr.Parameters.AddWithValue("?f14", int.Parse((Range2.Cells[68, 3] as Excel.Range).Text));
                mycmt4gr.ExecuteNonQuery();
                mycmt4gr.Parameters.Clear();


                //operational tab
                int i_for_reasons = 4;
                while (!String.IsNullOrEmpty((Range3.Cells[i_for_reasons, 4] as Excel.Range).Text))
                {
                   
                    mycmreasonsope.Prepare();
                    mycmreasonsope.CommandText = String.Format("insert into operational_affected(DateOfReport,SiteName,Region,IndicatorPrefArea,NameofPrefArea,Latitude,Longitude,Technology,Status,EventDateTime,OperationalReason,ActionsTaken,TTid,Comments) values (?date3_para,?i2,?i3,?i4,?name_para,?k1,?k2,?k3,?k4,?k5,?k6,?k7,?k8,?k9)");
                    mycmreasonsope.Parameters.AddWithValue("?date3_para", localdate);
                    mycmreasonsope.Parameters.AddWithValue("?i2", (Range3.Cells[i_for_reasons, 3] as Excel.Range).Text);
                    mycmreasonsope.Parameters.AddWithValue("?i3",(Range3.Cells[i_for_reasons, 4] as Excel.Range).Text);
                    if (!String.IsNullOrEmpty((Range3.Cells[i_for_reasons, 6] as Excel.Range).Text))
                    {
                       // MessageBox.Show(i_for_reasons.ToString() + "  " + (Range3.Cells[i_for_reasons, 5] as Excel.Range).Text+"Area");
                        mycmreasonsope.Parameters.AddWithValue("?i4", "Area");
                        mycmreasonsope.Parameters.AddWithValue("?name_para", (Range3.Cells[i_for_reasons, 6] as Excel.Range).Text);
                    }
                    else if(String.IsNullOrEmpty((Range3.Cells[i_for_reasons, 6] as Excel.Range).Text))
                    {
                     //   MessageBox.Show(i_for_reasons.ToString()+"  " + (Range3.Cells[i_for_reasons, 5] as Excel.Range).Text+"Pre");
                        mycmreasonsope.Parameters.AddWithValue("?i4", "Prefecture");
                        mycmreasonsope.Parameters.AddWithValue("?name_para", (Range3.Cells[i_for_reasons, 5] as Excel.Range).Text);
                    }
                    mycmreasonsope.Parameters.AddWithValue("?k1", double.Parse((Range3.Cells[i_for_reasons, 7] as Excel.Range).Text));
                    mycmreasonsope.Parameters.AddWithValue("?k2", double.Parse((Range3.Cells[i_for_reasons, 8] as Excel.Range).Text));
                    mycmreasonsope.Parameters.AddWithValue("?k3",((Range3.Cells[i_for_reasons, 9] as Excel.Range).Text));
                    mycmreasonsope.Parameters.AddWithValue("?k4", ((Range3.Cells[i_for_reasons, 10] as Excel.Range).Text));

                    try
                    {
                        // Convert EventDateTime Column DateTime Data Type
                        //MessageBox.Show((Range3.Cells[i_for_reasons, 11] as Excel.Range).Value2.ToString());
                        double d = double.Parse((Range3.Cells[i_for_reasons, 11] as Excel.Range).Value2.ToString());
                        DateTime EventDateTimeColumn = DateTime.FromOADate(d);
                        mycmreasonsope.Parameters.AddWithValue("?k5", EventDateTimeColumn);
                    }
                    catch( Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    mycmreasonsope.Parameters.AddWithValue("?k6",((Range3.Cells[i_for_reasons, 12] as Excel.Range).Text));
                    mycmreasonsope.Parameters.AddWithValue("?k7",((Range3.Cells[i_for_reasons, 13] as Excel.Range).Text));
                    mycmreasonsope.Parameters.AddWithValue("?k8", ((Range3.Cells[i_for_reasons, 14] as Excel.Range).Text));
                    mycmreasonsope.Parameters.AddWithValue("?k9", (Range3.Cells[i_for_reasons, 15] as Excel.Range).Text);
                   
                    mycmreasonsope.ExecuteNonQuery();
                    mycmreasonsope.Parameters.Clear();


                    i_for_reasons++;
                }

                //retention tab
                i_for_reasons = 4;
                while (!String.IsNullOrEmpty((Range4.Cells[i_for_reasons, 4] as Excel.Range).Text))
                {
                    
                    mycmreasonsret.Prepare();
                    mycmreasonsret.CommandText = String.Format("insert into retention_affected(DateOfReport,SiteName,Region,IndicatorPrefArea,NameofPrefArea,Latitude,Longitude,Technology,Status,EventDateTime,RetentionReason,ActionsTaken,TTid,Comments) values (?date3_para,?i2,?i3,?i4,?name_para,?k1,?k2,?k3,?k4,?k5,?k6,?k7,?k8,?k9)");
                    mycmreasonsret.Parameters.AddWithValue("?date3_para", localdate);
                    mycmreasonsret.Parameters.AddWithValue("?i2", ((Range4.Cells[i_for_reasons, 3] as Excel.Range).Text));
                    mycmreasonsret.Parameters.AddWithValue("?i3", ((Range4.Cells[i_for_reasons, 4] as Excel.Range).Text));
                    if (!String.IsNullOrEmpty((Range4.Cells[i_for_reasons, 6] as Excel.Range).Text))
                    {
                        mycmreasonsret.Parameters.AddWithValue("?i4", "Area");
                        mycmreasonsret.Parameters.AddWithValue("?name_para", (Range4.Cells[i_for_reasons, 6] as Excel.Range).Text);
                    }
                    else if(String.IsNullOrEmpty((Range4.Cells[i_for_reasons, 6] as Excel.Range).Text))
                    {
                        mycmreasonsret.Parameters.AddWithValue("?i4", "Prefecture");
                        mycmreasonsret.Parameters.AddWithValue("?name_para", (Range4.Cells[i_for_reasons, 5] as Excel.Range).Text);
                    }
                    mycmreasonsret.Parameters.AddWithValue("?k1", double.Parse((Range4.Cells[i_for_reasons, 7] as Excel.Range).Text));
                    mycmreasonsret.Parameters.AddWithValue("?k2", double.Parse((Range4.Cells[i_for_reasons, 8] as Excel.Range).Text));
                    mycmreasonsret.Parameters.AddWithValue("?k3", ((Range4.Cells[i_for_reasons, 9] as Excel.Range).Text));
                    mycmreasonsret.Parameters.AddWithValue("?k4",((Range4.Cells[i_for_reasons, 10] as Excel.Range).Text));
                    try
                    {
                        // Convert EventDateTime Column DateTime Data Type
                        double d = double.Parse((Range4.Cells[i_for_reasons, 11] as Excel.Range).Value2.ToString());
                        DateTime EventDateTimeColumn = DateTime.FromOADate(d);
                        mycmreasonsret.Parameters.AddWithValue("?k5", EventDateTimeColumn);
                    }
                    catch( Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    mycmreasonsret.Parameters.AddWithValue("?k6", ((Range4.Cells[i_for_reasons, 12] as Excel.Range).Text));
                    mycmreasonsret.Parameters.AddWithValue("?k7", ((Range4.Cells[i_for_reasons, 13] as Excel.Range).Text));
                    mycmreasonsret.Parameters.AddWithValue("?k8", ((Range4.Cells[i_for_reasons, 14] as Excel.Range).Text));
                    mycmreasonsret.Parameters.AddWithValue("?k9", ((Range4.Cells[i_for_reasons, 15] as Excel.Range).Text));

                    mycmreasonsret.ExecuteNonQuery();
                    mycmreasonsret.Parameters.Clear();


                    i_for_reasons++;
                }

                //licensing tab
                i_for_reasons = 4;
                while (!String.IsNullOrEmpty((Range5.Cells[i_for_reasons, 5] as Excel.Range).Text))
                {
                    
                    mycmreasonslic.Prepare();
                    mycmreasonslic.CommandText = String.Format("insert into licensing_affected(DateOfReport,SiteName,Region,IndicatorPrefArea,NameofPrefArea,Latitude,Longitude,Technology,Status,DeactivationDateTime,TTid,AffectedCoverage,ReactivationDate) values (?date3_para,?i2,?i3,?i4,?name_para,?k1,?k2,?k3,?k4,?k5,?k6,?k7,?k8)");
                    mycmreasonslic.Parameters.AddWithValue("?date3_para", localdate);
                    mycmreasonslic.Parameters.AddWithValue("?i2", ((Range5.Cells[i_for_reasons, 4] as Excel.Range).Text));
                    mycmreasonslic.Parameters.AddWithValue("?i3",((Range5.Cells[i_for_reasons, 5] as Excel.Range).Text));
                    if (!String.IsNullOrEmpty((Range5.Cells[i_for_reasons, 7] as Excel.Range).Text))
                    {
                        mycmreasonslic.Parameters.AddWithValue("?i4", "Area");
                        mycmreasonslic.Parameters.AddWithValue("?name_para", (Range5.Cells[i_for_reasons, 7] as Excel.Range).Text);
                    }
                    else if(String.IsNullOrEmpty((Range5.Cells[i_for_reasons, 7] as Excel.Range).Text))
                    {
                        mycmreasonslic.Parameters.AddWithValue("?i4", "Prefecture");
                        mycmreasonslic.Parameters.AddWithValue("?name_para", (Range5.Cells[i_for_reasons, 6] as Excel.Range).Text);
                    }
                    mycmreasonslic.Parameters.AddWithValue("?k1", double.Parse((Range5.Cells[i_for_reasons, 8] as Excel.Range).Text));
                    mycmreasonslic.Parameters.AddWithValue("?k2", double.Parse((Range5.Cells[i_for_reasons, 9] as Excel.Range).Text));
                    mycmreasonslic.Parameters.AddWithValue("?k3", ((Range5.Cells[i_for_reasons, 10] as Excel.Range).Text));
                    mycmreasonslic.Parameters.AddWithValue("?k4",((Range5.Cells[i_for_reasons, 11] as Excel.Range).Text));

                    // Convert EventDateTime Column DateTime Data Type
                    
                    try
                    {
                        // Convert EventDateTime Column DateTime Data Type
                        double d = double.Parse((Range5.Cells[i_for_reasons, 12] as Excel.Range).Value2.ToString());
                        DateTime DeActicationDateTimeColumn = DateTime.FromOADate(d);
                        mycmreasonslic.Parameters.AddWithValue("?k5", DeActicationDateTimeColumn);
                    }
                    catch( Exception ex)
                    {
                        MessageBox.Show((Range5.Cells[i_for_reasons, 12] as Excel.Range).Value2.ToString());
                        MessageBox.Show(ex.ToString());
                    }
                    mycmreasonslic.Parameters.AddWithValue("?k6",((Range5.Cells[i_for_reasons, 13] as Excel.Range).Text));
                    mycmreasonslic.Parameters.AddWithValue("?k7", ((Range5.Cells[i_for_reasons, 14] as Excel.Range).Text));
                    mycmreasonslic.Parameters.AddWithValue("?k8",((Range5.Cells[i_for_reasons, 15] as Excel.Range).Text));
                    

                    mycmreasonslic.ExecuteNonQuery();
                    mycmreasonslic.Parameters.Clear();


                    i_for_reasons++;
                }

                int i = 3, c;
                while (!String.IsNullOrEmpty((Range.Cells[i, 7] as Excel.Range).Text))
                {
                    myquery.Prepare();
                    myquery.CommandText = String.Format("select * FROM prefecture WHERE Name=?nameq");
                    //add the title parameter to search
                    myquery.Parameters.AddWithValue("?nameq", (Range.Cells[i, 7] as Excel.Range).Text);
                    try
                    {

                        //execute query
                        MySqlDataReader msdr = myquery.ExecuteReader();
                        if (!msdr.Read())
                        {
                            //prepare prefectures
                            mycm3.Prepare();
                            mycm3.CommandText = String.Format("insert into prefecture(Type,Latitude,Longtitude,Name) values (?i1,?i5,?i6,?i7)");
                            //MessageBox.Show(((Range.Cells[i, 1] as Excel.Range).Text));
                            mycm3.Parameters.AddWithValue("?i1", ((Range.Cells[i, 1] as Excel.Range).Text));
                            
                            mycm3.Parameters.AddWithValue("?i5", double.Parse((Range.Cells[i, 5] as Excel.Range).Text));
                            mycm3.Parameters.AddWithValue("?i6", double.Parse((Range.Cells[i, 6] as Excel.Range).Text));
                            mycm3.Parameters.AddWithValue("?i7", ((Range.Cells[i, 7] as Excel.Range).Text));
                            mycm3.ExecuteNonQuery();
                            mycm3.Parameters.Clear();
                        }
                        //closing the data reader and the connection
                       
                        msdr.Close();
                       
                     
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        MessageBox.Show(ex.ToString());
                    }

                    myquery.Parameters.Clear();



                    //prepare prefectures report
                    mycm4.Prepare();
                    mycm4.CommandText = String.Format("insert into prefecture_report(DateOfReport,Available2G,Available3G,Available4G,Name,Operational2G,Operational3G,Operational4G,Retention2G,Retention3G,Retention4G,Licensing2G,Licensing3G,Licensing4G,Unavailable2G,Unavailable3G,Unavailable4G) values (?date3_para,?i2,?i3,?i4,?name_para,?k1,?k2,?k3,?k4,?k5,?k6,?k7,?k8,?k9,?k10,?k11,?k12)");
                    mycm4.Parameters.AddWithValue("?date3_para", localdate);
                    mycm4.Parameters.AddWithValue("?i2", int.Parse((Range.Cells[i, 2] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?i3", int.Parse((Range.Cells[i, 3] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?i4", int.Parse((Range.Cells[i, 4] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?name_para", (Range.Cells[i, 7] as Excel.Range).Text);
                    
                    mycm4.Parameters.AddWithValue("?k1", int.Parse((Range.Cells[i, 8] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k2", int.Parse((Range.Cells[i, 9] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k3", int.Parse((Range.Cells[i, 10] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k4", int.Parse((Range.Cells[i, 60] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k5", int.Parse((Range.Cells[i, 61] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k6", int.Parse((Range.Cells[i, 62] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k7", int.Parse((Range.Cells[i, 106] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k8", int.Parse((Range.Cells[i, 107] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k9", int.Parse((Range.Cells[i, 108] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k10", int.Parse((Range.Cells[i, 111] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k11", int.Parse((Range.Cells[i, 112] as Excel.Range).Text));
                    mycm4.Parameters.AddWithValue("?k12", int.Parse((Range.Cells[i, 113] as Excel.Range).Text));
                    mycm4.ExecuteNonQuery();
                    mycm4.Parameters.Clear();



                    //Prepare Operational Cards
                    mycm5.Prepare();
                    mycm5.CommandText = String.Format("insert into operational_2G(DateOfReport,Antenna,CosmotePowerProblem,Disinfection,FiberCut,GeneratorFailure,Link,LinkDueToPowerProblem,OTEProblem,PowerProblem,PPCPowerFailure,Quality,RBSProblem,Temperature,VodafoneLinkProblem,VodafonePowerProblem,Modem) values (?date4_para,?p1,?p2,?p3,?p4,?p5,?p6,?p7,?p8,?p9,?p10,?p11,?p12,?p13,?p14,?p15,?p16)");
                    mycm5.Parameters.AddWithValue("?date4_para", localdate);
                    mycm5.Parameters.AddWithValue("?p1", int.Parse((Range.Cells[i, 11] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p2", int.Parse((Range.Cells[i, 12] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p3", int.Parse((Range.Cells[i, 13] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p4", int.Parse((Range.Cells[i, 14] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p5", int.Parse((Range.Cells[i, 15] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p6", int.Parse((Range.Cells[i, 16] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p7", int.Parse((Range.Cells[i, 17] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p8", int.Parse((Range.Cells[i, 18] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p9", int.Parse((Range.Cells[i, 19] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p10", int.Parse((Range.Cells[i, 20] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p11", int.Parse((Range.Cells[i, 21] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p12", int.Parse((Range.Cells[i, 22] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p13", int.Parse((Range.Cells[i, 23] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p14", int.Parse((Range.Cells[i, 24] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p15", int.Parse((Range.Cells[i, 25] as Excel.Range).Text));
                    mycm5.Parameters.AddWithValue("?p16", int.Parse((Range.Cells[i, 26] as Excel.Range).Text));
                    mycm5.ExecuteNonQuery();
                    mycm5.Parameters.Clear();



                    mycm6.Prepare();
                    mycm6.CommandText = String.Format("insert into operational_3G(DateOfReport,Antenna,CosmotePowerProblem,Disinfection,FiberCut,GeneratorFailure,Link,LinkDueToPowerProblem,OTEProblem,PowerProblem,PPCPowerFailure,Quality,RBSProblem,Temperature,VodafoneLinkProblem,VodafonePowerProblem,Modem) values (?date5_para,?s1,?s2,?s3,?s4,?s5,?s6,?s7,?s8,?s9,?s10,?s11,?s12,?s13,?s14,?s15,?s16)");
                    mycm6.Parameters.AddWithValue("?date5_para", localdate);
                    mycm6.Parameters.AddWithValue("?s1", int.Parse((Range.Cells[i, 27] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s2", int.Parse((Range.Cells[i, 28] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s3", int.Parse((Range.Cells[i, 29] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s4", int.Parse((Range.Cells[i, 30] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s5", int.Parse((Range.Cells[i, 31] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s6", int.Parse((Range.Cells[i, 32] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s7", int.Parse((Range.Cells[i, 33] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s8", int.Parse((Range.Cells[i, 34] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s9", int.Parse((Range.Cells[i, 35] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s10", int.Parse((Range.Cells[i, 36] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s11", int.Parse((Range.Cells[i, 37] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s12", int.Parse((Range.Cells[i, 38] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s13", int.Parse((Range.Cells[i, 39] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s14", int.Parse((Range.Cells[i, 40] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s15", int.Parse((Range.Cells[i, 41] as Excel.Range).Text));
                    mycm6.Parameters.AddWithValue("?s16", int.Parse((Range.Cells[i, 42] as Excel.Range).Text));
                    mycm6.ExecuteNonQuery();
                    mycm6.Parameters.Clear();




                    mycm7.Prepare();
                    mycm7.CommandText = String.Format("insert into operational_4G(DateOfReport,Antenna,CosmotePowerProblem,Disinfection,FiberCut,GeneratorFailure,Link,LinkDueToPowerProblem,OTEProblem,PowerProblem,PPCPowerFailure,Quality,RBSProblem,Temperature,VodafoneLinkProblem,VodafonePowerProblem,Modem) values (?date6_para,?d1,?d2,?d3,?d4,?d5,?d6,?d7,?d8,?d9,?d10,?d11,?d12,?d13,?d14,?d15,?d16)");
                    mycm7.Parameters.AddWithValue("?date6_para", localdate);
                    mycm7.Parameters.AddWithValue("?d1", int.Parse((Range.Cells[i, 43] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d2", int.Parse((Range.Cells[i, 44] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d3", int.Parse((Range.Cells[i, 45] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d4", int.Parse((Range.Cells[i, 46] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d5", int.Parse((Range.Cells[i, 47] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d6", int.Parse((Range.Cells[i, 48] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d7", int.Parse((Range.Cells[i, 49] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d8", int.Parse((Range.Cells[i, 50] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d9", int.Parse((Range.Cells[i, 51] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d10", int.Parse((Range.Cells[i, 52] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d11", int.Parse((Range.Cells[i, 53] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d12", int.Parse((Range.Cells[i, 54] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d13", int.Parse((Range.Cells[i, 55] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d14", int.Parse((Range.Cells[i, 56] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d15", int.Parse((Range.Cells[i, 57] as Excel.Range).Text));
                    mycm7.Parameters.AddWithValue("?d16", int.Parse((Range.Cells[i, 58] as Excel.Range).Text));
                    mycm7.ExecuteNonQuery();
                    mycm7.Parameters.Clear();

                    //Prepare retention Cards
                    mycm8.Prepare();
                    mycm8.CommandText = String.Format("insert into retention_2G(DateOfReport,Access,Antenna,Cabinet,DisasterDueToFire,DisasterDueToFlood,OwnerReaction,PeopleReaction,PPCIntention,Renovation,Shelter,Thievery,UnpaidBill,Vandalism,Reengineering) values (?date7_para,?f1,?f2,?f3,?f4,?f5,?f6,?f7,?f8,?f9,?f10,?f11,?f12,?f13,?f14)");
                    mycm8.Parameters.AddWithValue("?date7_para", localdate);
                    mycm8.Parameters.AddWithValue("?f1", int.Parse((Range.Cells[i, 63] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f2", int.Parse((Range.Cells[i, 64] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f3", int.Parse((Range.Cells[i, 65] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f4", int.Parse((Range.Cells[i, 66] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f5", int.Parse((Range.Cells[i, 67] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f6", int.Parse((Range.Cells[i, 68] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f7", int.Parse((Range.Cells[i, 69] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f8", int.Parse((Range.Cells[i, 70] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f9", int.Parse((Range.Cells[i, 71] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f10", int.Parse((Range.Cells[i, 72] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f11", int.Parse((Range.Cells[i, 73] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f12", int.Parse((Range.Cells[i, 74] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f13", int.Parse((Range.Cells[i, 75] as Excel.Range).Text));
                    mycm8.Parameters.AddWithValue("?f14", int.Parse((Range.Cells[i, 76] as Excel.Range).Text));
                    mycm8.ExecuteNonQuery();
                    mycm8.Parameters.Clear();


                    mycm9.Prepare();
                    mycm9.CommandText = String.Format("insert into retention_3G(DateOfReport,Access,Antenna,Cabinet,DisasterDueToFire,DisasterDueToFlood,OwnerReaction,PeopleReaction,PPCIntention,Renovation,Shelter,Thievery,UnpaidBill,Vandalism,Reengineering) values (?date8_para,?z1,?z2,?z3,?z4,?z5,?z6,?z7,?z8,?z9,?z10,?z11,?z12,?z13,?z14)");
                    mycm9.Parameters.AddWithValue("?date8_para", localdate);
                    mycm9.Parameters.AddWithValue("?z1", int.Parse((Range.Cells[i, 77] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z2", int.Parse((Range.Cells[i, 78] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z3", int.Parse((Range.Cells[i, 79] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z4", int.Parse((Range.Cells[i, 80] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z5", int.Parse((Range.Cells[i, 81] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z6", int.Parse((Range.Cells[i, 82] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z7", int.Parse((Range.Cells[i, 83] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z8", int.Parse((Range.Cells[i, 84] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z9", int.Parse((Range.Cells[i, 85] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z10", int.Parse((Range.Cells[i, 86] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z11", int.Parse((Range.Cells[i, 87] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z12", int.Parse((Range.Cells[i, 88] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z13", int.Parse((Range.Cells[i, 89] as Excel.Range).Text));
                    mycm9.Parameters.AddWithValue("?z14", int.Parse((Range.Cells[i, 90] as Excel.Range).Text));
                    mycm9.ExecuteNonQuery();
                    mycm9.Parameters.Clear();


                    mycm10.Prepare();
                    mycm10.CommandText = String.Format("insert into retention_4G(DateOfReport,Access,Antenna,Cabinet,DisasterDueToFire,DisasterDueToFlood,OwnerReaction,PeopleReaction,PPCIntention,Renovation,Shelter,Thievery,UnpaidBill,Vandalism,Reengineering) values (?date9_para,?x1,?x2,?x3,?x4,?x5,?x6,?x7,?x8,?x9,?x10,?x11,?x12,?x13,?x14)");
                    mycm10.Parameters.AddWithValue("?date9_para", localdate);
                    mycm10.Parameters.AddWithValue("?x1", int.Parse((Range.Cells[i, 91] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x2", int.Parse((Range.Cells[i, 92] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x3", int.Parse((Range.Cells[i, 93] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x4", int.Parse((Range.Cells[i, 94] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x5", int.Parse((Range.Cells[i, 95] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x6", int.Parse((Range.Cells[i, 96] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x7", int.Parse((Range.Cells[i, 97] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x8", int.Parse((Range.Cells[i, 98] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x9", int.Parse((Range.Cells[i, 99] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x10", int.Parse((Range.Cells[i, 100] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x11", int.Parse((Range.Cells[i, 101] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x12", int.Parse((Range.Cells[i, 102] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x13", int.Parse((Range.Cells[i, 103] as Excel.Range).Text));
                    mycm10.Parameters.AddWithValue("?x14", int.Parse((Range.Cells[i, 104] as Excel.Range).Text));
                    mycm10.ExecuteNonQuery();
                    mycm10.Parameters.Clear();

                    i++;                   
                }
                connection.Close();
                connection1.Close();
                MessageBox.Show("AddedFinally");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
              
            }







            XlWorkBook.Close(false, null, null);
            XlApp.Quit();
            releaseObject(XlWorkSheet);
            releaseObject(XlWorkSheet2);
            releaseObject(XlWorkBook);
            releaseObject(XlApp);
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object" + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

    }
}
