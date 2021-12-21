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
using Excel = Microsoft.Office.Interop.Excel;


namespace WindowsFormsAppBD
{
    public partial class Form1 : Form

    {
        private SqlCommandBuilder sqlBuilder = null;
        private SqlDataAdapter sqlDataAdapter = null;
        private SqlConnection sqlConnection = null;
        private DataSet dataSet = null;
        private SqlDataAdapter sqlDataAdapter2= null;
        private DataSet dataSet2 = null;
        private bool neRowAdding = false;
      
        public Form1()
        {
            InitializeComponent();
        }
        private void LoadData()
        {
            try
            {
                sqlDataAdapter = new SqlDataAdapter("SELECT *,  'Delete' AS[Delete] FROM Inventorizacia ", sqlConnection);
                sqlBuilder = new SqlCommandBuilder(sqlDataAdapter);
                sqlBuilder.GetInsertCommand();
                sqlBuilder.GetUpdateCommand();
                sqlBuilder.GetDeleteCommand();
                dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet, "Inventorizacia");
                dataGridView1.DataSource = dataSet.Tables["Inventorizacia"];
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[5, i] = linkCell;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

          
        }
        private void LoadData2()
        {
            try
            {
                sqlDataAdapter2 = new SqlDataAdapter("SELECT *,  'Delete' AS[Delete] FROM History ", sqlConnection);
                sqlBuilder = new SqlCommandBuilder(sqlDataAdapter2);
                sqlBuilder.GetInsertCommand();
                sqlBuilder.GetUpdateCommand();
                sqlBuilder.GetDeleteCommand();
                dataSet2 = new DataSet();
                sqlDataAdapter2.Fill(dataSet2, "History");
                dataGridView2.DataSource = dataSet2.Tables["History"];
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView2[4, i] = linkCell;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }




        private void ReloadData()
        {
            try
            {
                dataSet.Tables["Inventorizacia"].Clear();
                sqlDataAdapter.Fill(dataSet, "Inventorizacia");
                dataGridView1.DataSource = dataSet.Tables["Inventorizacia"];
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[5, i] = linkCell;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ReloadData2()
        {
            try
            {
                dataSet2.Tables["History"].Clear();
                sqlDataAdapter2.Fill(dataSet2, "History");
                dataGridView2.DataSource = dataSet2.Tables["History"];
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView2[4, i] = linkCell;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "database1DataSet2.History". При необходимости она может быть перемещена или удалена.
            this.historyTableAdapter.Fill(this.database1DataSet2.History);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "database1DataSet.Inventorizacia". При необходимости она может быть перемещена или удалена.
            this.inventorizaciaTableAdapter.Fill(this.database1DataSet.Inventorizacia);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["ivent"].ConnectionString);
            sqlConnection.Open();
            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM History",sqlConnection);
            DataSet db = new DataSet();
            dataAdapter.Fill(db);         
            dataGridView2.DataSource = db.Tables [0];
            LoadData();
            LoadData2();
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

        

        private void button4_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Parse(textBox8.Text);
            SqlCommand command = new SqlCommand(
                "INSERT INTO [History] (Invent_nomer,Nomera_Kabenetov,Data) values(@Invent_nomer,@Nomera_Kabenetov,@Data)",
                sqlConnection);
                
            command.Parameters.AddWithValue("Invent_nomer", textBox6.Text);
            command.Parameters.AddWithValue("Nomera_Kabenetov", textBox7.Text);
            command.Parameters.AddWithValue("Data", $"{date.Month}/{date.Day}/{date.Year}");
           

            command.ExecuteNonQuery();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = $"Nomera_Kabenetov LIKE '%{textBox9.Text}%'";
            
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = $"Invent_nomer LIKE '%{textBox10.Text}%'";
        }

       

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            ReloadData();
        }
        private void dataGridView2_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try 
            { 
                if(neRowAdding == false)
                {
                    neRowAdding = true;
                    int lastRow = dataGridView2.Rows.Count - 2;
                    DataGridViewRow row = dataGridView2.Rows[lastRow];
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    
                    dataGridView2[4, lastRow] = linkCell;
                    row.Cells["Delete"].Value = "Insert";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (neRowAdding == false)
                {
                    int rowIndex = dataGridView1.SelectedCells[0].RowIndex;

                    DataGridViewRow editingRow = dataGridView1.Rows[rowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[5, rowIndex] = linkCell;
                    editingRow.Cells["Delete"].Value = "Update";

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try 
            {
                if (e.ColumnIndex==5)
                    
                {
                    string task = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
                    if (task == "Delete")
                    {
                        if(MessageBox.Show("Удалить эту строку?","Удаление",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question)
                            ==DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;
                            dataGridView1.Rows.RemoveAt(rowIndex);
                            dataSet.Tables["Inventorizacia"].Rows[rowIndex].Delete();
                            sqlDataAdapter.Update(dataSet, "Inventorizacia");
                        }
                    }
                    else if(task == "Insert")
                    {



                        int rowIndex = dataGridView1.Rows.Count - 2;
                        DataRow row = dataSet.Tables["Inventorizacia"].NewRow();
                        row["Invent_nomer"] = dataGridView1.Rows[rowIndex].Cells["Invent_nomer"].Value;
                        row["Tip"] = dataGridView1.Rows[rowIndex].Cells["Tip"].Value;
                        row["Nazvanie"] = dataGridView1.Rows[rowIndex].Cells["Nazvanie"].Value;
                        row["Nomer_Kabineta"] = dataGridView1.Rows[rowIndex].Cells["Nomer_Kabineta"].Value;
                        dataSet.Tables["Inventorizacia"].Rows.Add(row);
                        dataSet.Tables["Inventorizacia"].Rows.RemoveAt(dataSet.Tables["Inventorizacia"].Rows.Count - 1);
                        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 2);
                        dataGridView1.Rows[e.RowIndex].Cells[5].Value = "Delete";
                        sqlDataAdapter.Update(dataSet, "Inventorizacia");
                        neRowAdding = false;

                    }
                    else if(task == "Update")
                    {
                        int r = e.RowIndex;
                        dataSet.Tables["Inventorizacia"].Rows[r]["Invent_nomer"] = dataGridView1.Rows[r].Cells["Invent_nomer"].Value;
                        dataSet.Tables["Inventorizacia"].Rows[r]["Tip"] = dataGridView1.Rows[r].Cells["Tip"].Value;
                        dataSet.Tables["Inventorizacia"].Rows[r]["Nazvanie"] = dataGridView1.Rows[r].Cells["Nazvanie"].Value;
                        dataSet.Tables["Inventorizacia"].Rows[r]["Nomer_Kabineta"] = dataGridView1.Rows[r].Cells["Nomer_Kabineta"].Value;
                        dataGridView1.Rows[e.RowIndex].Cells[5].Value = "Delete";
                        sqlDataAdapter.Update(dataSet, "Inventorizacia");
                    }
                    ReloadData();                       
                }
            }
            catch
            {

            }
        }

     

     

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = $"Invent_nomer LIKE '%{textBox11.Text}%'";
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = $"Tip LIKE '%{textBox12.Text}%'";
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ReloadData2();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex ==4 )

                {
                    string task = dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
                    if (task == "Delete")
                    {
                        if (MessageBox.Show("Удалить эту строку?", "Удаление", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                            == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;
                            dataGridView2.Rows.RemoveAt(rowIndex);
                            dataSet2.Tables["History"].Rows[rowIndex].Delete();
                            sqlDataAdapter2.Update(dataSet, "History");
                        }
                    }
                    else if (task == "Insert")
                    {
                        int rowIndex = dataGridView2.Rows.Count  -2;
                        DataRow row = dataSet2.Tables["History"].NewRow();
                        row["Invent_nomer"] = dataGridView2.Rows[rowIndex].Cells["History"].Value;
                        row["Nomera_Kabenetov"] = dataGridView2.Rows[rowIndex].Cells["Nomera_Kabenetov"].Value;
                        row["Data"] = dataGridView2.Rows[rowIndex].Cells["Data"].Value;
                        dataSet2.Tables["History"].Rows.Add(row);
                        dataSet2.Tables["History"].Rows.RemoveAt(dataSet2.Tables["History"].Rows.Count  -1);
                        dataGridView2.Rows.RemoveAt(dataGridView2.Rows.Count -2);
                        dataGridView2.Rows[e.RowIndex].Cells[4].Value = "Delete";
                        sqlDataAdapter2.Update(dataSet2,"History");
                        neRowAdding = false;
                    }

                    else if (task == "Update")
                    {
                        int r = e.RowIndex;
                        dataSet2.Tables["History"].Rows[r]["Invent_nomer"] = dataGridView2.Rows[r].Cells["Invent_nomer"].Value;
                        dataSet2.Tables["History"].Rows[r]["Nomera_Kabenetov"] = dataGridView2.Rows[r].Cells["Nomera_Kabenetov"].Value;
                        dataSet2.Tables["History"].Rows[r]["Data"] = dataGridView2.Rows[r].Cells["Data"].Value;                                            
                        dataGridView2.Rows[e.RowIndex].Cells[4].Value = "Delete";
                        sqlDataAdapter2.Update(dataSet2, "History");
                    }
                    ReloadData2();
                }
            }
            catch
            {

            }
        }

        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if (neRowAdding == false)
                {
                    neRowAdding = true;
                    int lastRow = dataGridView1.Rows.Count - 2;
                    DataGridViewRow row = dataGridView1.Rows[lastRow];
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[5, lastRow] = linkCell;
                    row.Cells["Delete"].Value = "Insert";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (neRowAdding == false)
                {
                    int rowIndex = dataGridView2.SelectedCells[0].RowIndex;

                    DataGridViewRow editingRow2 = dataGridView2.Rows[rowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView2[4, rowIndex] = linkCell;
                    editingRow2.Cells["Delete"].Value = "Update";

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            int j,i;
            
            for (i=0;i<=dataGridView2.RowCount-2;i++)
            {
               for ( j= 0;j
                          < dataGridView2.ColumnCount - 1;j++)
                {
                    wsh.Cells[i+1, j+1] = dataGridView2[j, i].Value.ToString();
                }

            }                          
            exApp.Visible = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            int j, i;

            for (i = 0; i <= dataGridView1.RowCount - 2; i++)
            {
                for (j = 0; j
                           < dataGridView1.ColumnCount - 1; j++)
                {
                    wsh.Cells[i + 1, j + 1] = dataGridView1[j, i].Value.ToString();
                }

            }
            exApp.Visible = true;
        }
    }

        
    

    
}
