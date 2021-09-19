using System;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToSQL
{
    public partial class Mes_from : Form
    {
        private string message;
        private static Form1 _form;
        Thread t = new Thread(() => { _form = new Form1(); });
        Button bt = new Button();
        public Mes_from(string mes)
        {
            InitializeComponent();
            message = mes;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Mes_from_Load(object sender, EventArgs e)
        {
            MainTimer.Start();
            switch (message)
            {
                case "OK":
                        break;
                case "CREATE_OK":
                    label1.Text = "Данные успешно сохранены\nСтруктура таблицы успешно изменена";
                    this.Size = new Size(this.Size.Width + 50, this.Size.Height);
                    break;
                case "LOAD":
                    t.Start();
                    //pictureBox1.BringToFront();
                    //pictureBox1.Dock = DockStyle.Fill;
                    //this.Opacity = 0;
                    break;
                case "START":
                    //pictureBox1.BringToFront();
                    //pictureBox1.Dock = DockStyle.Fill;
                    //pictureBox1.Image = Properties.Resources._364;
                    //this.ControlBox = false;
                    break;
                case "EDIT":
                    button1.Text = "Пропустить";
                    label1.Text = "Настройки успешно сохранены\nРекомендуется перезапустить программу";
                    this.Size = new System.Drawing.Size(350, 150);
                    bt.Location = new Point(button1.Location.X + button1.Size.Width + 5, button1.Location.Y + 45);
                    button1.Location = new Point(button1.Location.X, button1.Location.Y + 45);
                    bt.Size = new Size(button1.Size.Width + 10, button1.Size.Height); 
                    bt.Text = "Перезапустить";
                    bt.Click += (s,ep) => { Application.Restart(); };
                    panel1.Controls.Add(bt);

                    break;
            }
        }


        private void MainTimer_Tick(object sender, EventArgs e)
        {
            //if (Properties.Settings.Default.open)
            //{
            //    MainTimer.Stop();
            //    //this.Close();
            //    _form.ShowDialog();
            //}
            //this.Close();
        }
    }
}
