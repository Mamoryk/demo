using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutSourse
{
    public partial class Form1 : Form
    {
        DB DB = new DB();
        GV GV = new GV();
        public Form1()
        {
            InitializeComponent();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1 != null && textBox2 != null)
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable table = new DataTable();
                string querystring = @"select * from users where login=" + "'" + textBox1.Text + "'" + " and password =" + "'" + textBox2.Text + "'";
                SqlCommand command = new SqlCommand(querystring, DB.getConnection());
                adapter.SelectCommand = command;
                adapter.Fill(table);

                if (table.Rows.Count == 1)
                {
                    GV.role = Convert.ToInt32(table.Rows[0][2]);
                    GV.USERID = Convert.ToInt32(table.Rows[0][0]);
                    GV.USERNAME = Convert.ToString(table.Rows[0][1]);

                    if (GV.role == 1)
                    {
                        this.Hide();
                        Main main = new Main();
                        main.ShowDialog();
                        this.Show();


                    }
                    else if (GV.role == 2)
                    {
                        this.Hide();
                        Main main = new Main();
                        main.ShowDialog();
                        this.Show();
                    }
                    else if (GV.role == 3)
                    {
                        this.Hide();
                        Main main = new Main();
                        main.ShowDialog();
                        this.Show();
                    }
                }
                else
                {
                   


                        MessageBox.Show("Ошибка авторизации!", "Ошибка");
                    
                }

            }
            
        }
    }
}
