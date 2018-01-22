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
    public partial class Spring : Form
    {
        DataSet ds;
        SqlCommand com = new SqlCommand();
        Excel.IExcelDataReader xlReader = null;

        public Spring()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.IO.FileInfo fi = new System.IO.FileInfo(Application.StartupPath + @"\2014_Spring.xlsx");
            int fileid = 1;
            this.ReadFile(fi, fileid);
        }

        public void ReadFile(System.IO.FileInfo fi, int fileID)
        {

            using (System.IO.FileStream fs = new System.IO.FileStream(fi.FullName, System.IO.FileMode.Open))
            {
                if (fi.Extension.ToLower() == ".xls")
                {
                    xlReader = Excel.ExcelReaderFactory.CreateBinaryReader(fs);
                }
                else if (fi.Extension.ToLower() == ".xlsx")
                {
                    xlReader = Excel.ExcelReaderFactory.CreateOpenXmlReader(fs);
                }
            }
            if (xlReader == null)
                return;
            xlReader.IsFirstRowAsColumnNames = true;
            if (xlReader.ExceptionMessage != null)
            {
                //this._dtFiles.Rows[fileID - 1]["Status"] = "Failed";
                //this.dgv.Rows[fileID - 1].DefaultCellStyle.ForeColor = Color.Red;

                // this.writeExceptionToFile(Application.StartupPath + @"\" + fi.Name + "_Exceptions.txt"
                //      , xlReader.ExceptionMessage, true);
                MessageBox.Show(xlReader.ExceptionMessage);
                xlReader.Close();
                return;
            }



            DataSet ds = xlReader.AsDataSet();
            xlReader.Close();
            DataTable dt = ds.Tables[0];
            this.dataGridView1.DataSource = dt;
            _dtStudent = dt;
            this.ds = ds;

            //DataTable dt2 = ds.Tables[1];
            //this.dataGridView2.DataSource = dt2;

        }
        DataTable _dtStudent;

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < _dtStudent.Rows.Count; i++)

            {


                DataRow Student = _dtStudent.Rows[i];

                decimal Serial_No = Student["Serial_No"] == DBNull.Value ? 0 : decimal.Parse(Student["Serial_No"].ToString());

                String Form_No = Student["Form_No"] == DBNull.Value ? "" : Student["Form_No"].ToString();

                String Student_Name = Student["Student_Name"] == DBNull.Value ? "" : Student["Student_Name"].ToString();

                String Gender = Student["Gender"] == DBNull.Value ? "" : Student["Gender"].ToString();

                String Student_Contact = Student["Student_Contact"] == DBNull.Value ? "" : Student["Student_Contact"].ToString();

                String Father_Name = Student["Father_Name"] == DBNull.Value ? "" : Student["Father_Name"].ToString();

                String Father_Contact = Student["Father_Contact"] == DBNull.Value ? "" : Student["Father_Contact"].ToString();

                String Degree = Student["Degree"] == DBNull.Value ? "" : Student["Degree"].ToString();

                String Address = Student["Address"] == DBNull.Value ? "" : Student["Address"].ToString();

                String Fee_Status = Student["Fee_Status"] == DBNull.Value ? "" : Student["Fee_Status"].ToString();

                String Last_Degree = Student["Last_Degree"] == DBNull.Value ? "" : Student["Last_Degree"].ToString();

                String Obtained_Marks = Student["Obtained_Marks"] == DBNull.Value ? "" : Student["Obtained_Marks"].ToString();

                decimal Percentage = Student["Percentage"] == DBNull.Value ? 0 : decimal.Parse(Student["Percentage"].ToString());



                SqlConnection con = new SqlConnection();

                con.ConnectionString = "Data Source=DESKTOP-ILMCVPK;Initial Catalog=Student_Information_System;Integrated Security=True ";

                //SqlCommand com = new SqlCommand();

                com.Connection = con;

                com.CommandText = "Select * From Student";

                com.CommandType = CommandType.Text;

                SqlDataAdapter da = new SqlDataAdapter();

                da.SelectCommand = com;

                DataTable dt = new DataTable();

                con.Open();

                da.Fill(dt);

                da.Dispose();

                con.Close();


                dataGridView1.DataSource = dt;






            }





        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            Files files = new Files();
            files.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection();

            con.ConnectionString = "Data Source=DESKTOP-ILMCVPK;Initial Catalog=Student_Information_System;Integrated Security=True ";



            con.Open();


            for (int i = 0; i < ds.Tables.Count; i++)

            {




                DataTable dt = ds.Tables[i];

                dt.Columns.Add("File_Name");


                for (int r = 0; r < dt.Rows.Count; r++)

                {

                    SqlCommand com = new SqlCommand();

                    com.Connection = con;



                    DataRow row = dt.Rows[r];

                    String File_Name = row["File_Name"] == DBNull.Value ? "" : row["File_Name"].ToString();
                    decimal Serial_No = row["Serial_No"] == DBNull.Value ? 0 : decimal.Parse(row["Serial_No"].ToString());
                    String Form_No = row["Form_No"] == DBNull.Value ? "" : row["Form_No"].ToString();
                    String Student_Name = row["Student_Name"] == DBNull.Value ? "" : row["Student_Name"].ToString();
                    String Gender = row["Gender"] == DBNull.Value ? "" : row["Gender"].ToString();
                    String Student_Contact = row[7] == DBNull.Value ? "" : row[7].ToString();               
                    String Father_Name = row["Father_Name"] == DBNull.Value ? "" : row["Father_Name"].ToString();
                    String Father_Contact = row["Father_Contact"] == DBNull.Value ? "" : row["Father_Contact"].ToString();
                    String Degree = row["Degree"] == DBNull.Value ? "" : row["Degree"].ToString();
                    String Address = row["Address"] == DBNull.Value ? "" : row["Address"].ToString();
                    String Fee_Status = row["Fee_Status"] == DBNull.Value ? "" : row["Fee_Status"].ToString();
                    String Last_Degree = row["Last_Degree"] == DBNull.Value ? "" : row["Last_Degree"].ToString();
                    string Obtained_Marks = row["Obtained_Marks"] == DBNull.Value ? "" : row["Obtained_Marks"].ToString();
                    string Percentage = row["Percentage"] == DBNull.Value ? "" : row["Percentage"].ToString();



                    /* com.CommandText = "insert into Student(File_Name,Serial_No, Form_No,Student_Name,Gender,Student_Contact,Father_Name," +

                 "Father_Contact,Degree,Address,Fee_Status,Last_Degree,Obtained_Marks,Percentage ) Values("

                  + "'" + File_Name + "','" + Serial_No + "','" + Form_No + "','" + Student_Name + "','" + Gender + "','" + Student_Contact +

                               "','" + Father_Name + "','" + Father_Contact + "','" + Degree + "','" + Address + "','" + Fee_Status +

                               "','" + Last_Degree + "','" + Obtained_Marks + "','" + Percentage + "' )";*/



  //                  string a = "";

  //                  string b = ""; string c = ""; string d = ""; string f = ""; string g = ""; string h = "";

  //                  string j = "";

  //                  string k = ""; string l = ""; string m = ""; string q = ""; string o = "";

  //                  string p = "";


  //                  com.CommandText = "insert into Student(File_Name,Serial_No, Form_No,Student_Name,Gender,Student_Contact,Father_Name," +


  //"Father_Contact,Degree,Address,Fee_Status,Last_Degree,Obtained_Marks,Percentage )" + Environment.NewLine


  //                       + "Values(" + "'" + a + "','" + b + "','" + c + "','" + d.Trim() + "','" + f + "','" + g +

  //                       "','" + h.Trim() + "','" + j + "','" + k + "','" + l + "','" + m +

  //                       "','" + q + "','" + o + "','" + p + "' )";




  //                  com.CommandText = "insert into Student(File_Name,Serial_No, Form_No,Student_Name,Gender,Student_Contact,Father_Name," +

  //                     "Father_Contact,Degree,Address,Fee_Status,Last_Degree,Obtained_Marks,Percentage )" + Environment.NewLine

  //                      + "Values(@File_Name, @Serial_No, @Form_No,@Student_Name,@Gender,@Student_Contact,@Father_Name," +

  //                      "@Father_Contact,@Degree,@Address,@Fee_Status,@Last_Degree,@Obtained_Marks,@Percentage )";



  //                  com.CommandType = CommandType.Text;

  //                  com.Parameters.AddWithValue("@File_Name", a);

  //                  com.Parameters.AddWithValue("@Serial_No", b);

  //                  com.Parameters.AddWithValue("@Form_No", c);

  //                  com.Parameters.AddWithValue("@Student_Name", d.Trim());

  //                  com.Parameters.AddWithValue("@Gender", f);

  //                  com.Parameters.AddWithValue("@Student_Contact", g);

  //                  com.Parameters.AddWithValue("@Father_Name", h.Trim());

  //                  com.Parameters.AddWithValue("@Father_Contact", j);

  //                  com.Parameters.AddWithValue("@Degree", k);

  //                  com.Parameters.AddWithValue("@Address", l);

  //                  com.Parameters.AddWithValue("@Fee_Status", m);

  //                  com.Parameters.AddWithValue("@Last_Degree", q);

  //                  com.Parameters.AddWithValue("@Obtained_Marks", o);

  //                  com.Parameters.AddWithValue("@Percentage", p);


  //                  int n = com.ExecuteNonQuery();

                    com.CommandText = "insert into Student(File_Name,Serial_No, Form_No,Student_Name,Gender,Student_Contact,Father_Name," +
                        "Father_Contact,Degree,Address,Fee_Status,Last_Degree,Obtained_Marks,Percentage ) Values("
                        + "'" + File_Name + "','" + Serial_No + "','" + Form_No + "','" + Student_Name + "','" + Gender + "','" + Student_Contact +
                        "','" + Father_Name + "','" + Father_Contact + "','" + Degree + "','" + Address + "','" + Fee_Status +
                        "','" + Last_Degree + "','" + Obtained_Marks + "','" + Percentage + "' )";
                    int n = com.ExecuteNonQuery();

                }

                con.Close();

                MessageBox.Show("Completed");

            }

        }
    }
}
