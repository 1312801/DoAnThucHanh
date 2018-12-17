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
using System.Data.SqlClient;
namespace DoAn
{   
    public partial class Form1 : Form
    {
        string ConnectionString = "Data Source=DESKTOP-H4IGQS0;Catalog=AdventureWorksDW2012";
        string constring = "Server=DESKTOP-H4IGQS0;Database=AdventureWorksDW2012;Trusted_Connection=True;";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            AdomdConnection conn = new AdomdConnection(ConnectionString);
            SqlConnection con = new SqlConnection(constring);
            con.Open();
            conn.Open();
            Form2 form = new Form2(conn,con);
            form.Show();
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AdomdConnection conn = new AdomdConnection(ConnectionString);
            conn.Open();
            Form3 form = new Form3(conn);
            form.Show();
        }
    }
}
