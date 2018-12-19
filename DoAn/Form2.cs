using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.AnalysisServices.AdomdClient;
namespace DoAn
{
    public partial class Form2 : Form
    {
        String CountryCode = null;
        String RegionCode = null;
        String Nam = null;
        String Query = null;
        String member = null;
        AdomdConnection con1 = new AdomdConnection();
        SqlConnection con2 = new SqlConnection();
        public Form2()
        {
            InitializeComponent();

        }
        public Form2(AdomdConnection conn, SqlConnection con)
        {
            this.con1 = conn;
            this.con2 = con;
            InitializeComponent();
            FillNuoc();
            comboBox4.Hide();
            label2.Hide();
            FillNam();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            String time = this.comboBox1.GetItemText(this.comboBox1.SelectedItem);
            if (time == "Year")
            {
                Query = "SELECT {[Measures].[Profit],[Measures].[Income]} "
        + " ON COLUMNS,[Order Date].[Hierarchy].[Calendar Year] ON ROWS "
         + "FROM[Adventure Works DW2012] ";
                member = "[Order Date].[Hierarchy].[Calendar Year].[MEMBER_CAPTION]";

                comboBox3.SelectedItem = null;
            }
            if (time == "Quarter")
            {
                Query = "SELECT {[Measures].[Profit],[Measures].[Income]}"
        + "ON COLUMNS,[Order Date].[Hierarchy].[Calendar Year].&["+Nam+"].Children ON ROWS "
         + " FROM[Adventure Works DW2012] ";
                member = "[Order Date].[Hierarchy].[Calendar Quarter].[MEMBER_CAPTION]";
            }
            if (time == "Month")
            {
                Query = "SELECT {[Measures].[Profit],[Measures].[Income]} "
       + "ON COLUMNS,([Order Date].[Calendar Year].&[" + Nam + "],[Order Date].[Month Number Of Year].[Month Number Of Year]) ON ROWS "
        + " FROM[Adventure Works DW2012] ";
                member = "[Order Date].[Month Number Of Year].[Month Number Of Year].[MEMBER_CAPTION]";
            }
            if (time == "Week")
            {
                Query = "SELECT {[Measures].[Profit],[Measures].[Income]} "
       + "ON COLUMNS,([Order Date].[Calendar Year].&[" + Nam + "],[Order Date].[Week Number Of Year].[Week Number Of Year]) ON ROWS "
        + " FROM[Adventure Works DW2012] ";
                member = "[Order Date].[Week Number Of Year].[Week Number Of Year].[MEMBER_CAPTION]";
            }
        }

        private void chart1_Click(object sender, EventArgs e)
        {

            
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //SqlConnection sql = con.CreateCommand();
            //query theo quarter
            String profit = @" WITH MEMBER [Measures].[Profit] AS  
   (  
        [Measures].[Sales Amount] -   
         [Measures].[Total Product Cost]
  ), FORMAT_STRING = '#,#.###'  ";
            String income = "MEMBER [Measures].[Income] AS ([Measures].[Sales Amount] -[Measures].[Tax Amt]) ,FORMAT_STRING = '#,#.###'  ";

   String condition= " WHERE[Dim Customer].[State Province Code].&["+RegionCode+"] &["+CountryCode+"]";
            AdomdCommand cmd = con1.CreateCommand();
            cmd.CommandText = profit + Query + condition;
            AdomdDataAdapter ad = new AdomdDataAdapter(profit+income + Query+condition, con1);
            DataTable dt = new DataTable();
            ad.Fill(dt);
            chart1.DataSource = dt;
            chart1.ChartAreas["ChartArea1"].AxisY.Title = "Money";
            chart1.ChartAreas["ChartArea1"].AxisX.Title = "Date";

            chart1.Series["Income"].XValueMember =member;
            chart1.Series["Income"].YValueMembers = "[Measures].[Income]";
            chart1.Series["Profit"].YValueMembers = "[Measures].[Profit]";
        }

        private void chart1_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            String nuoc=null;
            nuoc= this.comboBox2.GetItemText(this.comboBox2.SelectedItem);
            string query = "select distinct CountryRegionCode from [dbo].[DimGeography] where EnglishCountryRegionName='" +nuoc +"'";
            SqlDataReader reader = null;
            SqlCommand cmd = new SqlCommand(query, con2);
            try
            {
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                     CountryCode = reader["CountryRegionCode"].ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                }
            }
            comboBox4.Items.Clear();
            comboBox4.Show();
            label2.Show();
            FillKhuVuc(CountryCode);
        }

        private void chart1_Click_2(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        void FillNam()
        {
            string query = " Select distinct CalendarYear from dbo.DimDate order by CalendarYear";
            SqlCommand cmd = new SqlCommand(query, con2);
            SqlDataReader reader = null;
            try
            {
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string Name = reader["CalendarYear"].ToString();
                    comboBox3.Items.Add(Name);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                }
            }
        }
        void  FillNuoc()
        {
            
            string query = " Select distinct EnglishCountryRegionName from dbo.DimGeography";
            SqlCommand cmd = new SqlCommand(query, con2);
            SqlDataReader reader = null;
            try
            {
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string Name = reader["EnglishCountryRegionName"].ToString();
                    comboBox2.Items.Add(Name);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if(reader!=null)
                {
                    reader.Close();
                }
            }
        }
        void FillKhuVuc(string Nuoc)
        {
            
            string query = "Select distinct StateProvinceName from dbo.DimGeography where CountryRegionCode='"+Nuoc+"'";
            SqlCommand cmd = new SqlCommand(query, con2);
            SqlDataReader reader=null;
            try
            {
                reader = cmd.ExecuteReader();
                while(reader.Read())
                {
                    string Name = reader["StateProvinceName"].ToString();
                    comboBox4.Items.Add(Name);
                }
            }catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                }
            }

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            String region = null;
            region = this.comboBox4.GetItemText(this.comboBox4.SelectedItem);
            string query = "select distinct StateProvinceCode from [dbo].[DimGeography] where StateProvinceName='" +region+ "'";
            SqlDataReader reader = null;
            SqlCommand cmd = new SqlCommand(query, con2);
            try
            {
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    RegionCode = reader["StateProvinceCode"].ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                }
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Nam = this.comboBox3.GetItemText(this.comboBox3.SelectedItem);
        }
    }
}
