using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.AnalysisServices.AdomdClient;
namespace DoAn
{
    public partial class Form3 : Form
    {
        AdomdConnection con = new AdomdConnection();
        public Form3()
        {
            InitializeComponent();
        }
        public Form3(AdomdConnection conn)
        {
            this.con = conn;
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Hide();
            string query = @" WITH MEMBER [Measures].[Profit] AS  
 (  
[Measures].[Sales Amount] -   
[Measures].[Total Product Cost]
  ), FORMAT_STRING = '#,#.###'  " + "SELECT {[Measures].[Profit],[Measures].[Sales Amount]}  ON COLUMNS,[Order Date].[Hierarchy].[Calendar Year].&[2013].Children  ON ROWS "
+ " FROM[Adventure Works DW2012] " 
+ " WHERE[Dim Customer].[State Province Code].&[BY] &[DE]";
            AdomdCommand cmd = con.CreateCommand();
            cmd.CommandText = query;
            AdomdDataAdapter ad = new AdomdDataAdapter(query, con);
            DataTable dt = new DataTable();
            ad.Fill(dt);
            chart1.DataSource = dt;
            chart1.ChartAreas["ChartArea1"].AxisX.Title = "Sales Amount";
            chart1.ChartAreas["ChartArea1"].AxisX.Title = "Date";

            chart1.Series["Series1"].XValueMember = "[Order Date].[Hierarchy].[Calendar Quarter].[MEMBER_CAPTION]";
            chart1.Series["Series2"].XValueMember = "[Order Date].[Hierarchy].[Calendar Quarter].[MEMBER_CAPTION]";
            chart1.Series["Series2"].YValueMembers = "[Measures].[Profit]";
            chart1.Series["Series1"].YValueMembers = "[Measures].[Sales Amount]";

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string query = @" WITH MEMBER [Measures].[Profit] AS  
 (  
[Measures].[Sales Amount] -   
[Measures].[Total Product Cost]
  ), FORMAT_STRING = '#,#.###'  " + "SELECT {[Measures].[Profit],[Measures].[Sales Amount]}  ON COLUMNS,[Order Date].[Hierarchy].[Calendar Year].&[2013].Children  ON ROWS "
+ " FROM[Adventure Works DW2012] "
+ " WHERE[Dim Customer].[State Province Code].&[BY] &[DE]";
            AdomdCommand cmd = con.CreateCommand();
            cmd.CommandText = query;
            AdomdDataAdapter ad = new AdomdDataAdapter(query, con);
            DataTable dt = new DataTable();
            ad.Fill(dt);
            dataGridView1.DataSource = dt;
           
        }
    }
}
