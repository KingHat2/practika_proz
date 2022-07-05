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

    {   //переменые для работы с бд
        private SqlCommandBuilder sqlBuilder = null;
        private SqlDataAdapter sqlDataAdapter = null;
        private SqlConnection sqlConnection = null;
        private DataSet dataSet = null;
        private SqlDataAdapter sqlDataAdapter2= null;
        private DataSet dataSet2 = null;
        private bool neRowAdding = false;
        private SqlCommandBuilder sqlBuilder2 = null;
        public Form1()
        {
            InitializeComponent();
        }
        //загрузка таблицы из бд в dataGridView
        private void LoadData()
        {
            //блок  обработки исключения
            try
            {   //инцилизация класса SqlDataAdapter  //передача в качестве первого параметра Sql запрос а в качетсве второго параметра экземпляр класса sqlConnection
                sqlDataAdapter = new SqlDataAdapter("SELECT *,  'Delete' AS[Delete] FROM Inventorizacia ", sqlConnection);
               //инцилизация поля sqlCommanBuilder
                sqlBuilder = new SqlCommandBuilder(sqlDataAdapter);
                //генерация команд Insert/Update/Delete
                sqlBuilder.GetInsertCommand();
                sqlBuilder.GetUpdateCommand();
                sqlBuilder.GetDeleteCommand();
                //инцилизация  поля dataSet
                dataSet = new DataSet();
                //заполнение dataSet с помощью sqlDataAdapter
                sqlDataAdapter.Fill(dataSet, "Inventorizacia");
                //установка таблицы Inventorizacia из dataSet для dataGridView
                dataGridView1.DataSource = dataSet.Tables["Inventorizacia"];
                //переопределие 6 колонки в linlLable
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    //обращение к ячейки
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[5, i] = linkCell;

                }
            }
            //вывод сообщения о ошибки
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

          
        }

        //загрузка таблицы из бд в dataGridView
        private void LoadData2()
        {
            //блок  обработки исключения
            try
            {
                //инцилизация класса SqlDataAdapter  //передача в качестве первого параметра Sql запрос а в качетсве второго параметра экземпляр класса sqlConnection
                sqlDataAdapter2 = new SqlDataAdapter("SELECT *,  'Delete' AS[Delete] FROM History ", sqlConnection);
                //инцилизация поля sqlCommanBuilder
                sqlBuilder2 = new SqlCommandBuilder(sqlDataAdapter2);
                //генерация команд Insert/Update/Delete
                sqlBuilder2.GetInsertCommand();
                sqlBuilder2.GetUpdateCommand();
                sqlBuilder2.GetDeleteCommand();
                //инцилизация  поля dataSet
                dataSet2 = new DataSet();
                //заполнение dataSet с помощью sqlDataAdapter
                sqlDataAdapter2.Fill(dataSet2, "History");
                //установка таблицы History из dataSet для dataGridView
                dataGridView2.DataSource = dataSet2.Tables["History"];
                //переопределие 5 колонки в linlLable
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    //обращение к ячейки
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView2[4, i] = linkCell;

                }
            }
            //вывод сообщения о ошибки
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }



        //обновление Бд
        private void ReloadData()
        {
            //блок  обработки исключения
            try
            {   //очищиение таблицы
                dataSet.Tables["Inventorizacia"].Clear();
                //заполнение dataSet с помощью sqlDataAdapter
                sqlDataAdapter.Fill(dataSet, "Inventorizacia");
                //установка таблицы Inventorizacia из dataSet для dataGridView
                dataGridView1.DataSource = dataSet.Tables["Inventorizacia"];
                //переопределие 6 колонки в linlLable
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    //обращение к ячейки
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[5, i] = linkCell;
                }

            }
            //вывод сообщения о ошибки
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        //обновление Бд
        private void ReloadData2()
        {
            //блок  обработки исключения
            try
            {   //очищиение таблицы
                dataSet2.Tables["History"].Clear();
                //заполнение dataSet с помощью sqlDataAdapter
                sqlDataAdapter2.Fill(dataSet2, "History");
                //установка таблицы History из dataSet для dataGridView
                dataGridView2.DataSource = dataSet2.Tables["History"];
                //переопределие 5 колонки в linlLable
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    //обращение к ячейки
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView2[4, i] = linkCell;
                }
            }
            //вывод сообщения о ошибки
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //загрузка формы
        private void Form1_Load(object sender, EventArgs e)
        {

            // TODO: данная строка кода позволяет загрузить данные в таблицу "database1DataSet2.History". При необходимости она может быть перемещена или удалена.
            this.historyTableAdapter.Fill(this.database1DataSet2.History);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "database1DataSet.Inventorizacia". При необходимости она может быть перемещена или удалена.
            this.inventorizaciaTableAdapter.Fill(this.database1DataSet.Inventorizacia);
            //строка подключения к бд
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["ivent"].ConnectionString);
            //открытие подключения к бд
            sqlConnection.Open();
            //передача в конструктор запрос выбора всех сталбцов из таблицы History
            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM History",sqlConnection);
            //создание экземпляра класса
            DataSet db = new DataSet();
            //заполнение DataSet c помощью DataAdapter
            dataAdapter.Fill(db);
            //устновка таблице в качестве источника данных для dataGridView
            dataGridView2.DataSource = db.Tables [0];
            //вызов метода LoadData
            LoadData();
            LoadData2();
        }

        //добавление новых данных с помощью button
        private void button1_Click(object sender, EventArgs e)
        { //создание экземпляра класса
            SqlCommand command = new SqlCommand(             
                "INSERT INTO [Inventorizacia] (Invent_nomer,Tip,Nazvanie,Nomer_Kabineta) values(@Invent_nomer,@Tip,@Nazvanie,@Nomer_Kabineta)",sqlConnection);
            //связывание ключей и texbox
            command.Parameters.AddWithValue("Invent_nomer", textBox1.Text);
            command.Parameters.AddWithValue("Tip", textBox2.Text);
            command.Parameters.AddWithValue("Nazvanie", textBox3.Text);
            command.Parameters.AddWithValue("Nomer_Kabineta", textBox4.Text);
            // выполнение команды не возращят не каких данных только int значение
            command.ExecuteNonQuery();
        }

        //выход из программы
        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }


        //добавление новых данных с помощью button
        private void button4_Click(object sender, EventArgs e)
        {  //приведение строки к DateTime
            DateTime date = DateTime.Parse(textBox8.Text);
            //создание экземпляра класса
            SqlCommand command = new SqlCommand(
                "INSERT INTO [History] (Invent_nomer,Nomera_Kabenetov,Data) values(@Invent_nomer,@Nomera_Kabenetov,@Data)",
                sqlConnection);
            //связывание ключей и texbox
            command.Parameters.AddWithValue("Invent_nomer", textBox6.Text);
            command.Parameters.AddWithValue("Nomera_Kabenetov", textBox7.Text);
            command.Parameters.AddWithValue("Data", $"{date.Month}/{date.Day}/{date.Year}");
            // выполнение команды не возращят не каких данных только int значение
            command.ExecuteNonQuery();
        }

        //поиск данных с помощью TexBox
        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            // получение DataSoucre из dataGridView  //фильтр с присвоинной строкой
             (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = $"Nomera_Kabenetov LIKE '%{textBox9.Text}%'";
            
        }

        //поиск данных с помощью TexBox
        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            //получение DataSoucre из dataGridView  //фильтр с присвоинной строкой
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = $"Invent_nomer LIKE '%{textBox10.Text}%'";
        }

        //кнопка обновления
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            //вызов метода обновления
            ReloadData();
        }
     
        //обработка выбора команды Insert
        private void dataGridView2_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            { //провенрка neRowAdding == false что бы не путались переменные при Insert и Update
                if (neRowAdding == false)
                {
                    //добавление новой строки
                    neRowAdding = true;
                    //добавление строки в последнию строку
                    int lastRow = dataGridView2.Rows.Count - 2;
                    //используя индекс последний строки в которой добавляем новую ячейку для создания класса DataGridViewRow
                    DataGridViewRow row = dataGridView2.Rows[lastRow];
                    //создание linkCell
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    //устновка LinkCell в 5 ячейку
                    dataGridView2[4, lastRow] = linkCell;
                    //переменование строки Delete в Insert
                    row.Cells["Delete"].Value = "Insert";
                }
            }
            //вывод сообщения о ошибки
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //обработчик события для Update
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //neRowAdding == false проверка на редактирование данных
                if (neRowAdding == false)
                {
                    //получение индекса выделенной строки
                    int rowIndex = dataGridView1.SelectedCells[0].RowIndex;
                    //созданние экземпляра класса dataGridViewRow с присвоеннием строки по индексу по индексу которой мы присвоили переменную
                    DataGridViewRow editingRow = dataGridView1.Rows[rowIndex];
                    //создание linkCell
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    //устновка LinkCell в 6 ячейку
                    dataGridView1[5, rowIndex] = linkCell;
                    //переименование 6 ячейки Delete в Update
                    editingRow.Cells["Delete"].Value = "Update";
                }
            }
            //вывод сообщения о ошибки
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Insert/Update/Delete
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //блок  обработки исключения
            try
            {   //проверка нажатия на 5 ячейку
                if (e.ColumnIndex==5)
                    
                {   
                    //получение текста из linkLabel 
                    string task = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
                    //проверка какую команды хотел выполнить пользователь
                    if (task == "Delete")
                    {
                        //вывод MessageBox с вопросом удаление строки
                        if (MessageBox.Show("Удалить эту строку?","Удаление",MessageBoxButtons.YesNo,MessageBoxIcon.Question)
                            ==DialogResult.Yes)
                        {
                            //создание переменной
                            int rowIndex = e.RowIndex;
                            //вызов метода RemoveAt
                            dataGridView1.Rows.RemoveAt(rowIndex);
                            //удаление этой строки из dataSet
                            dataSet.Tables["Inventorizacia"].Rows[rowIndex].Delete();
                            //обновление данных в бд
                            sqlDataAdapter.Update(dataSet, "Inventorizacia");
                        }
                    }

                    //проверка какую команды хотел выполнить пользователь
                    else if (task == "Insert")
                    {    
                        //созданние int переменой с индексом стрроки
                        int rowIndex = dataGridView1.Rows.Count - 2;
                        //создание переменой куда запишим ссылку на новую строку которую мы создадим в DataSet в таблице Inventorizacia
                        DataRow row = dataSet.Tables["Inventorizacia"].NewRow(); 
                        //занесение в новые строки данные из dataGridView 
                        row["Invent_nomer"] = dataGridView1.Rows[rowIndex].Cells["Invent_nomer"].Value;
                        row["Tip"] = dataGridView1.Rows[rowIndex].Cells["Tip"].Value;
                        row["Nazvanie"] = dataGridView1.Rows[rowIndex].Cells["Nazvanie"].Value;
                        row["Nomer_Kabineta"] = dataGridView1.Rows[rowIndex].Cells["Nomer_Kabineta"].Value;
                        //добавление новой строки в dataSet
                        dataSet.Tables["Inventorizacia"].Rows.Add(row);
                        dataSet.Tables["Inventorizacia"].Rows.RemoveAt(dataSet.Tables["Inventorizacia"].Rows.Count - 1);
                        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 2);
                        //установка для 6 ячейки текс Delete
                        dataGridView1.Rows[e.RowIndex].Cells[5].Value = "Delete";
                        //обновление данных в бд и занесение строки из dataGridView и dataSet
                        sqlDataAdapter.Update(dataSet, "Inventorizacia");
                        //установка neRowAdding в значение false
                        neRowAdding = false;

                    }

                    //проверка какую команды хотел выполнить пользователь
                    else if (task == "Update")
                    {                       
                        //полученние индекса выделенной строки
                        int r = e.RowIndex;
                        //обновление всех данных в dataSet
                        dataSet.Tables["Inventorizacia"].Rows[r]["Invent_nomer"] = dataGridView1.Rows[r].Cells["Invent_nomer"].Value;
                        dataSet.Tables["Inventorizacia"].Rows[r]["Tip"] = dataGridView1.Rows[r].Cells["Tip"].Value;
                        dataSet.Tables["Inventorizacia"].Rows[r]["Nazvanie"] = dataGridView1.Rows[r].Cells["Nazvanie"].Value;
                        dataSet.Tables["Inventorizacia"].Rows[r]["Nomer_Kabineta"] = dataGridView1.Rows[r].Cells["Nomer_Kabineta"].Value;
                        //замена текста на 6 ячейки на Delete
                        dataGridView1.Rows[e.RowIndex].Cells[5].Value = "Delete";
                        //убрает необходимость нажимать на enter для Update после вводы новых данных
                        this.Validate();
                        this.dataGridView1.EndEdit();
                        //обновление данных в бд
                        sqlDataAdapter.Update(dataSet, "Inventorizacia");
                    }
                    //обновление Бд
                    ReloadData();                       
                }
            }
            catch
            {

            }
        }




        //поиск данных с помощью TexBox
        private void textBox11_TextChanged(object sender, EventArgs e)
        {  
            //получение DataSoucre из dataGridView  //фильтр с присвоинной строкой
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = $"Invent_nomer LIKE '%{textBox11.Text}%'";
        }

        //поиск данных с помощью TexBox
        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            //получение DataSoucre из dataGridView  //фильтр с присвоинной строкой
             (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = $"Tip LIKE '%{textBox12.Text}%'";
        }

        //кнопка обновления
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            //вызов метода обновления
            ReloadData2();
        }
       
        //Insert/Update/Delete
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //блок  обработки исключения
            try
            {  //проверка нажатия на 4 ячейку
                if (e.ColumnIndex ==4 )

                {
                    //получение текста из linkLabel 
                    string task = dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
                    //проверка какую команды хотел выполнить пользователь
                    if (task == "Delete")
                    {
                        //вывод MessageBox с вопросом удаление строки
                        if (MessageBox.Show("Удалить эту строку?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                            == DialogResult.Yes)
                        {
                            //создание переменной
                            int rowIndex = e.RowIndex;
                            //вызов метода RemoveAt
                            dataGridView2.Rows.RemoveAt(rowIndex);
                            //удаление этой строки из dataSet
                            dataSet2.Tables["History"].Rows[rowIndex].Delete();
                            //обновление данных в бд
                            sqlDataAdapter2.Update(dataSet2, "History");

                        }
                    }

                    //проверка какую команды хотел выполнить пользователь
                    else if (task == "Insert")
                    {
                        //созданние int переменой с индексом стрроки
                        int rowIndex = dataGridView2.Rows.Count - 2;
                        //создание переменой куда запишим ссылку на новую строку которую мы создадим в DataSet в таблице History
                        DataRow row = dataSet2.Tables["History"].NewRow();
                        //занесение в новые строки данные из dataGridView 
                        row["Invent_nomer"] = dataGridView2.Rows[rowIndex].Cells["Invent_nomer"].Value;                       
                        row["Nomera_Kabenetov"] = dataGridView2.Rows[rowIndex].Cells["Nomera_Kabenetov"].Value;
                        row["Data"] = dataGridView2.Rows[rowIndex].Cells["Data"].Value;
                        //добавление новой строки в dataSet
                        dataSet2.Tables["History"].Rows.Add(row);
                        dataSet2.Tables["History"].Rows.RemoveAt(dataSet2.Tables["History"].Rows.Count - 1);
                        dataGridView2.Rows.RemoveAt(dataGridView2.Rows.Count - 2);
                        //установка для 5 ячейки текс Delete
                        dataGridView2.Rows[e.RowIndex].Cells[4].Value = "Delete";
                        //обновление данных в бд и занесение строки из dataGridView и dataSet
                        sqlDataAdapter2.Update(dataSet2, "History");
                        //установка neRowAdding в значение false
                        neRowAdding = false;
                    }

                    //проверка какую команды хотел выполнить пользователь
                    else if (task == "Update")
                    {
                        //полученние индекса выделенной строки
                        int r = e.RowIndex;
                        //обновление всех данных в dataSet
                        dataSet2.Tables["History"].Rows[r]["Invent_nomer"] = dataGridView2.Rows[r].Cells["Invent_nomer"].Value;
                        dataSet2.Tables["History"].Rows[r]["Nomera_Kabenetov"] = dataGridView2.Rows[r].Cells["Nomera_Kabenetov"].Value;
                        dataSet2.Tables["History"].Rows[r]["Data"] = dataGridView2.Rows[r].Cells["Data"].Value;
                        //замена текста на 5 ячейки на Delete
                        dataGridView2.Rows[e.RowIndex].Cells[4].Value = "Delete";
                      


                        //list box
                        // BindingSource BS = new BindingSource();
                        // DataTable DT = new DataTable();
                        // DT.Columns.Clear();
                        // DT.Columns.Add("field1_name");
                        // DT.Columns.Add("field2_name");
                        // BS.DataSource = DT;
                        // dataGridView2.AutoGenerateColumns = false;
                        //dataGridView2.DataSource = BS;

                        // Добавляем выпадающий список
                        // DataGridViewComboBoxColumn col2 = new DataGridViewComboBoxColumn();
                        // col2.Items.AddRange("Значение1", "Значение2");
                        // col2.DataPropertyName = "field2_name";
                        // col2.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
                        // col2.FlatStyle = FlatStyle.Flat;
                        //  dataGridView2.Columns.Add(col2);

                        //убрает необходимость нажимать на enter для Update после вводы новых данных
                        this.Validate();
                        this.dataGridView2.EndEdit();
                        //обновление данных в бд
                        sqlDataAdapter2.Update(dataSet2, "History");




                    }
                    //обновление Бд
                    ReloadData2();
                }
            }
            catch
            {

            }
        }
       
        //обработка выбора команды Insert
        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                // провенрка neRowAdding == false что бы не путались переменные при Insert и Update
                if (neRowAdding == false)
                {
                    //добавление новой строки
                    neRowAdding = true;
                    //добавление строки в последнию строку
                    int lastRow = dataGridView1.Rows.Count - 2;
                    //используя индекс последний строки в которой добавляем новую ячейку для создания класса DataGridViewRow
                    DataGridViewRow row = dataGridView1.Rows[lastRow];
                    //создание linkCell
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    //устновка LinkCell в 4 ячейку
                    dataGridView1[5, lastRow] = linkCell;
                    //переменование строки Delete в Insert
                    row.Cells["Delete"].Value = "Insert";
                }
            }
            //вывод сообщения о ошибки
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //обработчик события для Update
        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //neRowAdding == false проверка на редактирование данных
                if (neRowAdding == false)
                {
                    //получение индекса выделенной строки
                    int rowIndex = dataGridView2.SelectedCells[0].RowIndex;
                    //созданние экземпляра класса dataGridViewRow с присвоеннием строки по индексу по индексу которой мы присвоили переменную
                    DataGridViewRow editingRow = dataGridView2.Rows[rowIndex];
                    //создание linkCell
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    //переименование 5 ячейки Delete в Update
                    dataGridView2[4, rowIndex] = linkCell;
                    //переименование 5 ячейки Delete в Update
                    editingRow.Cells["Delete"].Value = "Update";
                }
            }
            //вывод сообщения о ошибки
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
       
        //Открытие БД Inventorizacia в excel
        private void button3_Click(object sender, EventArgs e)
        {
            //Открытие БД Inventorizacia в excel
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            int j,i;            
            for (i=0;i<=dataGridView2.RowCount-2;i++)
            {
               for ( j= 0;j< dataGridView2.ColumnCount - 1;j++)
                {
                    wsh.Cells[i+1, j+1] = dataGridView2[j, i].Value.ToString();
                }

            }                          
            exApp.Visible = true;
        }

        //Открытие БД Inventorizacia в excel
        private void button5_Click(object sender, EventArgs e)
        {
            //Открытие БД History в excel
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            int j, i;
            for (i = 0; i <= dataGridView1.RowCount - 2; i++)
            {
                for (j = 0; j < dataGridView1.ColumnCount - 1; j++)
                {
                    wsh.Cells[i + 1, j + 1] = dataGridView1[j, i].Value.ToString();
                }

            }
            exApp.Visible = true;
        }

        //Ограничение на ввод букв в поля для цифр
        private void dataGridView2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        { //Ограничение на ввод букв в поля для цифр
            e.Control.KeyPress -= new KeyPressEventHandler(Colum_KeyPress);
            if (dataGridView2.CurrentCell.ColumnIndex == 1)
            {
                TextBox textBox = e.Control as TextBox;
                if(textBox !=null)
                {
                    textBox.KeyPress += new KeyPressEventHandler(Colum_KeyPress);

                }
            }       
        }

        //Ограничение на ввод букв в поля для цифр
        private void Colum_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar)&& !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }

        }

        //Ограничение на ввод букв в поля для цифр
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        { //Ограничение на ввод букв в поля для цифр
            e.Control.KeyPress -= new KeyPressEventHandler(Colum_KeyPress);
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                TextBox textBox = e.Control as TextBox;
                if (textBox != null)
                {
                    textBox.KeyPress += new KeyPressEventHandler(Colum_KeyPress);

                }
            }
        }

        
    }
SELEC * FROM table ORDER BY fio;//выбрать всё из таблицы 
SELECT fio AS 'ФИО', phone AS 'Телефон' FROM usersTable Order by fio;//переименование строк
}
