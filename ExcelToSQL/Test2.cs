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
    public partial class Test2 : Form
    {
        public Test2(int x, int y)
        {
            this.Location = new Point(x, y);
            this.Size =new Size(300, 300);
            InitializeComponent();
        }

        private void Test2_Load(object sender, EventArgs e)
        {

        }
    }
}
