using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelDataReader
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure to Exit?","Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)== DialogResult.Yes)
            {
                Application.Exit();
                
            }
        }

        private void iToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Registration r= new Registration();
            r.Show();
        }

        private void importFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

            Files files = new Files();
            files.Show();
        }
    }
}
