using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;

namespace WindowsFormsAppBD
{
    public partial class Form1 : Form
    {
        private SqlConnection sqlConnection = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["ivent"].ConnectionString);
            sqlConnection.Open();
          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand(
                "INSERT INTO [Inventorizacia] (Invent_nomer,Tip,Nazvanie,Nomer_Kabineta) values(@Invent_nomer,@Tip,@Nazvanie,@Nomer_Kabineta)",
                sqlConnection);
            command.Parameters.AddWithValue("Invent_nomer", textBox1.Text);
            command.Parameters.AddWithValue("Tip", textBox2.Text);
            command.Parameters.AddWithValue("Nazvanie", textBox3.Text);
            command.Parameters.AddWithValue("Nomer_Kabineta", textBox4.Text);
            
            command.ExecuteNonQuery();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();


        }

        private void button3_Click(object sender, EventArgs e)
         {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
              textBox5.Text,
              sqlConnection);
            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0];
        }
    }
}
