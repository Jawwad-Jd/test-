using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ExcelDataReader
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ILMCVPK;Initial Catalog=Student_Information_System;Integrated Security=True");
            SqlDataAdapter sda = new SqlDataAdapter("Select Count(*) From Login where UserName='" + textBox1.Text +"' and Password ='" + textBox2.Text + "'", conn);
            DataTable tbl = new DataTable();
            sda.Fill(tbl);
            if (tbl.Rows[0][0].ToString() == "1")
            {


               
                FrmMain fm = new FrmMain();
                fm.Show();
            }
            else
                MessageBox.Show("You have entered wrong Password");
        }
    }
}
