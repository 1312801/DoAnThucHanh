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
        String CountryCode=null;
        String RegionCode=null;
        String Nam = null;
        AdomdConnection con1= new AdomdConnection();
        SqlConnection con2 = new SqlConnection();
        public Form2()
        {
            InitializeComponent();
            
        }
        public Form2(AdomdConnection conn,SqlConnection con)
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
            int selectedIndex = comboBox1.SelectedIndex;
            if(selectedIndex==0)
            {
                comboBox2.Show();
            }
        }

        private void chart1_Click(object sender, EventArgs e)
        {

            
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //SqlConnection sql = con.CreateCommand();
            //query quarter
            string query = @" WITH MEMBER [Measures].[Profit] AS  
   (  
        [Measures].[Sales Amount] -   
         [Measures].[Total Product Cost]
  ), FORMAT_STRING = '#,#.###'  " + "SELECT {[Measures].[Profit],[Measures].[Sales Amount]}  ON COLUMNS,[Order Date].[Hierarchy].[Calendar Year].&["+Nam+"].Children  ON ROWS "
   + " FROM[Adventure Works DW2012] "
   + " WHERE[Dim Customer].[State Province Code].&["+RegionCode+"] &["+CountryCode+"]";
            AdomdCommand cmd = con1.CreateCommand();
            cmd.CommandText = query;
            AdomdDataAdapter ad = new AdomdDataAdapter(query, con1);
            DataTable dt = new DataTable();
            ad.Fill(dt);
            chart1.DataSource = dt;
            chart1.ChartAreas["ChartArea1"].AxisX.Title = "Sales Amount";
            chart1.ChartAreas["ChartArea1"].AxisX.Title = "Date";

            chart1.Series["Income"].XValueMember = "[Order Date].[Hierarchy].[Calendar Quarter].[MEMBER_CAPTION]";
            chart1.Series["Income"].YValueMembers = "[Measures].[Sales Amount]";
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
