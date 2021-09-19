using System;
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
    public partial class Test1 : Form
    {
        public Test1(int x,int y)
        {
            this.Location = new Point(x, y);
            this.Size = new Size(300, 300);
            InitializeComponent();
        }

        private void Test1_Load(object sender, EventArgs e)
        {

        }
    }
}
