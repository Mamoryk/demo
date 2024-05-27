using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

using System.IO;
using System.Windows.Controls;

using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

using System.Diagnostics;

namespace OutSourse
{
    public partial class Main : Form
    {
        
        DB DB = new DB();
        public Main()
        {
            InitializeComponent();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close(); 
        }

        private void pictureBox2_Click_2(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            if (GV.role == 1)
            {



            }
            else if (GV.role == 2)
            {
                button15.Visible = false;
               
            }
            else if (GV.role == 3)
            {
                button15.Visible = false;
                button8.Visible = false;
                button11.Visible = false;
                button5.Visible = false;
                button6.Visible = false;
                button9.Visible = false;
            }
            try
            {
                comboBox6.Items.Clear();
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"select * from role";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(table);
                foreach (DataRow row in table.Rows)
                {
                    comboBox6.Items.Add(row["rolename"]);
                }

            }
            catch
            {
                MessageBox.Show("Ошибка заполнения списка!", "Ошибка");
            }
            try
            {
                comboBox1.Items.Clear();
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"select * from tasktype";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(table);
                foreach (DataRow row in table.Rows)
                {
                    comboBox1.Items.Add(row["tasktypename"]);
                }

            }
            catch
            {
                MessageBox.Show("Ошибка заполнения списка!", "Ошибка");
            }
            try
            {
                comboBox2.Items.Clear();
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"select * from taskpriority";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(table);
                foreach (DataRow row in table.Rows)
                {
                    comboBox2.Items.Add(row["taskpriorityname"]);
                }

            }
            catch
            {
                MessageBox.Show("Ошибка заполнения списка!", "Ошибка");
            }
            try
            {
                comboBox3.Items.Clear();
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"select * from taskstatus";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(table);
                foreach (DataRow row in table.Rows)
                {
                    comboBox3.Items.Add(row["taskstatusname"]);
                }

            }
            catch
            {
                MessageBox.Show("Ошибка заполнения списка!", "Ошибка");
            }
            label1.Text = GV.USERNAME;

            try
            {
                comboBox10.Items.Clear();
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"select * from projectstatus";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(table);
                foreach (DataRow row in table.Rows)
                {
                    comboBox10.Items.Add(row["projectstatusname"]);
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"select * from role where roleid="+GV.role+"";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(table);
                label2.Text = table.Rows[0][1].ToString();
            }
            catch
            {
                MessageBox.Show("Ошибка заполнения таблицы!", "Ошибка");
            }
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"select projectid as [Номер проекта],projectname as [Название],projecttext as[Информация],createdate as[Дата],client as[Заказчик],projectstatusname as [Статус] from project,projectstatus where projectstatus = projectstatusid";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(table);
                dataGridView1.DataSource = table;
            }
            catch
            {
                MessageBox.Show("Ошибка заполнения таблицы!", "Ошибка");
            }

            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"select projectid as [Номер проекта],projectname as [Название] from project";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(table);
                dataGridView2.DataSource = table;
            }
            catch
            {
                MessageBox.Show("Ошибка заполнения таблицы!", "Ошибка");
            }
            
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView2_Click(object sender, EventArgs e)
        {
            try
            {

                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable dt = new DataTable();
                string querystring = @"select taskid as [Номер задачи],project as [Проект],tasktypename as [тип задачи],tasktext as [Текст],Taskpriorityname as [Приоритет],deadline as [Дэдлайн],taskstatusname as[Статус задачи] from task,taskpriority,taskstatus,tasktype where task.tasktype = tasktype.tasktypeid and task.taskpriority =taskpriority.taskpriorityid and task.taskstatus = taskstatus.taskstatusid and  project = " + dataGridView2.SelectedRows[0].Cells[0].Value.ToString() + "";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(dt);
                dataGridView3.DataSource = dt;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"delete from project where projectid = " + dataGridView1.SelectedRows[0].Cells[0].Value.ToString() + "";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                DB.openConnection();

                if (command.ExecuteNonQuery() == 1)
                {


                    MessageBox.Show("Проект Удален!");



                }
                else
                {

                }
                DB.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();

                string querystring = $"INSERT INTO project " +
                    $"(projectname,projecttext,createdate,client,projectstatus)VALUES ('{textBox5.Text}','{textBox4.Text}',getdate(),'{textBox2.Text}',"+1+") ;";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                DB.openConnection();

                if (command.ExecuteNonQuery() == 1)
                {


                    MessageBox.Show("Проект создан! ");



                }
                else
                {

                }
                DB.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"select projectid as [Номер проекта],projectname as [Название],projecttext as[Информация],createdate as[Дата],client as[Заказчик],projectstatusname as [Статус] from project,projectstatus where projectstatus = projectstatusid";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(table);
                dataGridView1.DataSource = table;
            }
            catch
            {
                MessageBox.Show("Ошибка заполнения таблицы!", "Ошибка");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"update project set projectstatus = (select projectstatusid from projectstatus where projectstatusname  = '"+comboBox10.Text+"')where projectid = " + dataGridView1.SelectedRows[0].Cells[0].Value.ToString() + "";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                DB.openConnection();

                if (command.ExecuteNonQuery() == 1)
                {


                    MessageBox.Show("Статус изменен!");



                }
                else
                {

                }
                DB.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage1;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage2;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage3;
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"INSERT INTO task (project, tasktype,tasktext,taskpriority,deadline,taskstatus)
                    VALUES (" + dataGridView2.SelectedRows[0].Cells[0].Value.ToString() + "," + Convert.ToInt32(comboBox1.SelectedIndex + 1) + ",'" + textBox3.Text + "'," + Convert.ToInt32(comboBox2.SelectedIndex + 1) + ",'" + dateTimePicker3.Value.ToString() + "'," + Convert.ToInt32(comboBox3.SelectedIndex + 1) + ")";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                DB.openConnection();

                if (command.ExecuteNonQuery() == 1)
                {


                    MessageBox.Show("Задача добавлена!");



                }
                else
                {

                }
                DB.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"delete from task where taskid = " + dataGridView3.SelectedRows[0].Cells[0].Value.ToString() + "";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                DB.openConnection();

                if (command.ExecuteNonQuery() == 1)
                {


                    MessageBox.Show("Задача Удалена!");



                }
                else
                {

                }
                DB.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView3_Click(object sender, EventArgs e)
        {
            try
            {

                textBox1.Text = dataGridView3.SelectedRows[0].Cells[1].Value.ToString();
                comboBox1.Text = dataGridView3.SelectedRows[0].Cells[2].Value.ToString();
                textBox3.Text = dataGridView3.SelectedRows[0].Cells[3].Value.ToString();
                comboBox2.Text = dataGridView3.SelectedRows[0].Cells[4].Value.ToString();
                dateTimePicker3.Text = dataGridView3.SelectedRows[0].Cells[5].Value.ToString();
                comboBox3.Text = dataGridView3.SelectedRows[0].Cells[6].Value.ToString();
                if (comboBox2.Text == "Низкий")
                {
                    panel1.BackColor = Color.Green;
                    comboBox2.ForeColor = Color.Green;
                }
                else if (comboBox2.Text == "Средний")
                {
                    panel1.BackColor = Color.CornflowerBlue;
                    comboBox2.ForeColor = Color.CornflowerBlue;
                }
                else if (comboBox2.Text == "Высокий")
                {
                    panel1.BackColor = Color.Orange;
                    comboBox2.ForeColor = Color.Orange;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {

                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable dt = new DataTable();
                string querystring = @"select userid as [Код работника],FIO as [ФИО],rolename as[Роль],login as[Логин],password as[Пароль] from users,role where users.role=role.roleid";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(dt);
                dataGridView4.DataSource = dt;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"delete from users where userid = " + dataGridView4.SelectedRows[0].Cells[0].Value.ToString() + "";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                DB.openConnection();

                if (command.ExecuteNonQuery() == 1)
                {


                    MessageBox.Show("Запись Удалена!");



                }
                else
                {

                }
                DB.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"INSERT INTO users (
                        FIO,
                        role,  
                        login,
                        password
                       )
                    VALUES ('" + textBox8.Text + "',(select roleid from role where rolename = '" + comboBox6.Text + "'),'" + textBox6.Text + "','" + textBox24.Text + "')";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                DB.openConnection();

                if (command.ExecuteNonQuery() == 1)
                {


                    MessageBox.Show("Запись добавлена!");



                }
                else
                {

                }
                DB.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage4;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"update  task set  taskstatus = (select taskstatusid from taskstatus where taskstatusname = '" + comboBox3.Text +"') where taskid = " + dataGridView3.SelectedRows[0].Cells[0].Value.ToString() + "";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                DB.openConnection();

                if (command.ExecuteNonQuery() == 1)
                {


                    MessageBox.Show("Задача изменена!");



                }
                else
                {

                }
                DB.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage5;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {

                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable dt = new DataTable();
                string querystring = @"select * From files";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(dt);
                dataGridView5.DataSource = dt;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"delete from files where fileid = " + dataGridView5.SelectedRows[0].Cells[0].Value.ToString() + "";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                DB.openConnection();

                if (command.ExecuteNonQuery() == 1)
                {


                    MessageBox.Show("Запись Удалена!");



                }
                else
                {

                }
                DB.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            application.Visible = true;
            Microsoft.Office.Interop.Word.Document document = application.Documents.Open(@""+ dataGridView5.SelectedRows[0].Cells[2].Value.ToString()  +"",
                                                                            Type.Missing, Type.Missing,
                                                                            Type.Missing, Type.Missing,
                                                                            Type.Missing, Type.Missing,
                                                                            Type.Missing, Type.Missing);

         
        }

        private void button20_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text Files (*.docx)|*.docx|All Files (*.*)|*.*";
            openFileDialog.Multiselect = false; // выбрать один файл

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                var fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(filePath);
               
                try
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    DataTable table = new DataTable();
                    string querystring = @"INSERT INTO files (
                        filename,
                        directory
                       )
                    VALUES ('" + fileNameWithoutExtension+ "','" + filePath + "')";
                    SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                    DB.openConnection();

                    if (command.ExecuteNonQuery() == 1)
                    {


                        MessageBox.Show("Запись добавлена!");



                    }
                    else
                    {

                    }
                    DB.closeConnection();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Word Document (*.docx)|*.docx";
            saveDialog.Title = "Save as Word Document";
            saveDialog.FileName = "ExportedData";

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveDialog.FileName;

                // Создаем новый экземпляр Microsoft Word
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = false;

                // Создаем новый документ
                var doc = wordApp.Documents.Add();

                // Добавляем заголовок
                var para = doc.Paragraphs.Add();
                para.Range.Text = "Exported Data from DataGridView";
                para.Range.InsertParagraphAfter();

                // Добавляем таблицу и заполняем ее данными из DataGridView
                var table = doc.Tables.Add(para.Range, dataGridView1.Rows.Count + 1, dataGridView1.Columns.Count);
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    table.Cell(1, i + 1).Range.Text = dataGridView1.Columns[i].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        table.Cell(i + 2, j + 1).Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

                // Сохраняем документ и закрываем Word
                doc.SaveAs(filePath);
                doc.Close();
                wordApp.Quit();

                // Открываем сохраненный файл
                Process.Start(filePath);
            }
        }
        
    }
}
