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
    /// <summary>hhfjhjgjhlhgfdefghj;.,m
    /// hayeeeeeeeeeeeeeeeeeeee yooooyoooooyooooyooooooyooooo
    /// </summary>you dangroooooooooooooooo
    public delegate void Handler();	// a simple delegate for marshalling calls from event handlers to the GUI thread


    // This delegate enables asynchronous calls for setting  
    // the text property on a TextBox control.  
    delegate void StringArgReturningVoidDelegate(string text);
    // This delegate enables asynchronous calls for setting  
    // the text property on a TextBox control.  
    delegate void VoidDelegate();


    public partial class Form1 : Form
    {

        DataSet ds;
        SqlCommand com = new SqlCommand();
        int lastcolindex = 17;
        

        Excel.IExcelDataReader xlReader = null;
        // DataSet ds = xlReader.ToDataset();


        public Form1()
        {
            InitializeComponent();


        }

        private void button1_Click(object sender, EventArgs e)
        {
            dlgOpen.RestoreDirectory = true;
            string filename = Application.StartupPath + @"\C:\Users\Hira Mukhtar\Documents\Data BIMS";
            if (DialogResult.OK == (new Invoker(dlgOpen).Invoke()))
            { /*handle result*/
                filename = this.dlgOpen.FileName;

                System.IO.FileInfo fi = new System.IO.FileInfo(filename);
                int fileid = 1;
                this.ReadFile(fi, fileid);
                //this.dataGridView1.DataSource = dtMaster;

            }
            //dlgOpen.ShowDialog();


            //System.IO.FileInfo fi = new System.IO.FileInfo(Application.StartupPath + @"\BSCS Fall 2014.xlsx");
            //int fileid = 1;
            //this.ReadFile(fi, fileid);
        }



        DataTable _dtStudent2;


        private void button2_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < _dtStudent2.Rows.Count; i++)
            {

                DataRow Student = _dtStudent2.Rows[i];
                string Serial_No = Student["Serial_No"] == DBNull.Value ? "" : Student["Serial_No"].ToString();
                String Registration_No = Student["Registration_No"] == DBNull.Value ? "" : Student["Registration_No"].ToString();
                String Student_Name = Student["Student_Name"] == DBNull.Value ? "" : Student["Student_Name"].ToString();
                //  String Gender = Student["Gender"] == DBNull.Value ? "" : Student["Gender"].ToString();
                //  String Student_Contact = Student["Student_Contact"] == DBNull.Value ? "" : Student["Student_Contact"].ToString();
                String Father_Name = Student["Father_Name"] == DBNull.Value ? "" : Student["Father_Name"].ToString();
                String Contact = Student["Contact"] == DBNull.Value ? "" : Student["Contact"].ToString();
                String CNIC = Student["CNIC"] == DBNull.Value ? "" : Student["CNIC"].ToString();
                // String Birth_Date = Student["Birth_Date"] == DBNull.Value ? "" : Student["Birth_Date"].ToString();
                String Birth_Date = Student[12] == DBNull.Value ? "" : Student[12].ToString();


                String Domicile = Student["Domicile"] == DBNull.Value ? "" : Student["Domicile"].ToString();
                String Permanent_Address = Student["Permanent_Address"] == DBNull.Value ? "" : Student["Permanent_Address"].ToString();
                String Admission_Date = Student["Admission_Date"] == DBNull.Value ? "" : Student["Admission_Date"].ToString();
                string Passing_Year = Student["Passing_Year"] == DBNull.Value ? "" : Student["Passing_Year"].ToString();
                String Remarks = Student["Remarks"] == DBNull.Value ? "" : Student["Remarks"].ToString();


                SqlConnection con = new SqlConnection();
                con.ConnectionString = "Data Source=DESKTOP-ILMCVPK;Initial Catalog=Student_Information_System;Integrated Security=True ";
                //SqlCommand com = new SqlCommand();
                com.Connection = con;
                com.CommandText = "Select * From Student2";
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





        private void button3_Click(object sender, EventArgs e)
        {

            SqlConnection con = new SqlConnection();
            con.ConnectionString = "Data Source=DESKTOP-41ku449;Initial Catalog=Student_Information_System;Integrated Security=True ";

            con.Open();
            //ds = xlReader.AsDataSet();

            for (int i = 0; i < ds.Tables.Count; i++)
            {

                DataTable dt = ds.Tables[i];

                if (i == 0)
                    _dtStudent2 = dt.Clone();
                //        dt.Columns.Add("File_Name");

                for (int r = 2; r < dt.Rows.Count; r++)
                {
                    SqlCommand com = new SqlCommand();
                    com.Connection = con;


                    DataRow row = dt.Rows[r];

                    DataRow MasterRow = _dtStudent2.NewRow();


                    for (int ch = 0; ch <= lastcolindex; ch++)
                    {
                        MasterRow[ch] = row[ch];

                    }
                    _dtStudent2.Rows.Add(MasterRow);



                    //String File_Name = row["File_Name"]== DBNull.Value ? "" : row["File_Name"].ToString(); 
                    string Serial_No = row[0] == DBNull.Value ? "" : row[0].ToString();
                    String Column1 = row[1] == DBNull.Value ? "" : row[1].ToString();
                    String Column3 = row[3] == DBNull.Value ? "" : row[3].ToString();
                    //            String Gender = row["Gender"] == DBNull.Value ? "" : row["Gender"].ToString();
                    //            String Student_Contact = row["Student_Contact"] == DBNull.Value ? "" : row["Student_Contact"].ToString();
                    String Column5 = row[5] == DBNull.Value ? "" : row[5].ToString();
                    String Column7 = row[7] == DBNull.Value ? "" : row[7].ToString();
                    String Column9 = row[9] == DBNull.Value ? "" : row[9].ToString();
                    String Column11 = row[11] == DBNull.Value ? "" : row[11].ToString();
                    String Column13 = row[13] == DBNull.Value ? "" : row[13].ToString();
                    String Column15 = row[15] == DBNull.Value ? "" : row[15].ToString();
                    string Column17 = row[17] == DBNull.Value ? "" : row[17].ToString();
                    string Column19 = row[19] == DBNull.Value ? "" : row[19].ToString();
                    string Column21 = row[21] == DBNull.Value ? "" : row[21].ToString();

                    /*    com.CommandText = "insert into Student2(Serial_No,Registration_No,Student_Name,Father_Name,Contact," +
                                   "CNIC,Birth_Date,Domicile,Permanent_Address,Admission_Date,Passing_Year,Remarks ) Values("
                       //            + "'" + Serial_No + "','" + Registration_No + "','" + Student_Name + "','" + Father_Name + "','" +Contact + "','" + CNIC +
                       //            "','" + Birth_Date + "','" + Domicile + "','" + Permanent_Address + "','" + Admission_Date + 
                       //            "','" + Passing_Year + "','" + Remarks + "' )";*/

                    /*  string a = "";
                      string b = ""; string c = ""; string d = ""; string f = ""; string g = ""; string h = "";
                      string j = "";
                      string k = ""; string l = ""; string m = ""; string q = ""; //string o = "";
                                                                                  // string p = "";

                      com.CommandText = "insert into Student2(Serial_No,Registration_No,Student_Name,Father_Name,Contact," +
            "CNIC,Birth_Date,Domicile,Permanent_Address,Admission_Date,Passing_Year,Remarks )" + Environment.NewLine
                            + "Values(" + "'" + a + "','" + b + "','" + c + "','" + d + "','" + f + "','" + g +
                           "','" + h + "','" + j + "','" + k + "','" + l + "','" + m +
                            "','" + q + "' )";


                      com.CommandText = "insert into Student2(Serial_No,Registration_No,Student_Name,Father_Name,Contact," +
          "CNIC,Birth_Date,Domicile,Permanent_Address,Admission_Date,Passing_Year,Remarks )" + Environment.NewLine
                        + "Values(@Serial_No,@Registration_No,@Student_Name,@Father_Name,@Contact," +
          "@CNIC,@Birth_Date,@Domicile,@Permanent_Address,@Admission_Date,@Passing_Year,@Remarks )";

                      com.CommandType = CommandType.Text;
                      com.Parameters.AddWithValue("@Serial_No", a);
                      com.Parameters.AddWithValue("@Registration_No", b);
                      com.Parameters.AddWithValue("@Student_Name", c);
                      com.Parameters.AddWithValue("@Father_Name", d);
                      com.Parameters.AddWithValue("@Contact", f);
                      com.Parameters.AddWithValue("@CNIC", g);
                      com.Parameters.AddWithValue("@Birth_Date", h);
                      com.Parameters.AddWithValue("@Domicile", j);
                      com.Parameters.AddWithValue("@Permanent_Address", k);
                      com.Parameters.AddWithValue("@Admission_Date", l);
                      com.Parameters.AddWithValue("@Passing_Year", m);
                      com.Parameters.AddWithValue("@Remarks", q);
                      //            com.Parameters.AddWithValue("@Obtained_Marks", o);
                      //            com.Parameters.AddWithValue("@Percentage", p);
                      */


                    com.CommandText = "insert into Student2(Serial_No,Registration_No,Student_Name,Father_Name,Contact," +
          "CNIC,Birth_Date,Domicile,Permanent_Address,Admission_Date,Passing_Year,Remarks  ) Values("
                    + "'" + Serial_No + "','" + Column1 + "','" + Column3 + "','" + Column5 + "','" + Column7 + "','" + Column9 +
                    "','" + Column11 + "','" + Column13 + "','" + Column15 + "','" + Column17 + "','" + Column19 +
                    "','" + Column21 + "' )";

                    int n = com.ExecuteNonQuery();


                }

            }
            con.Close();


            //string filename="";
            //String[] Tokens = filename.Split(new char[] { ' ' });
            //String Program = Tokens[0];
            //String SemesterNmae = Tokens[1] + " " + Tokens[2];
            //int Year = int.Parse(Tokens[2]);
            //int Semtype = Tokens[1] == "Spring" ? 1 : 3;
            //int Semester_ID = Year * 10 + Semtype;


            MessageBox.Show("Completed");

        }


        private void button4_Click(object sender, EventArgs e)
        {
            /*DataTable dtMaster2 = dtMaster.Clone();
            for(int i=0; i<dtMaster.Rows.Count; i++)
            {

            }*/

            Application.Exit();

        }
        
        private void Remove_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection();
            con.ConnectionString = "Data Source=DESKTOP-41ku449;Initial Catalog=Student_Information_System;Integrated Security=True ";

            con.Open();
            //com.CommandText= "Select Distinct Column1 From Student2 where FileName ='" + "BSCS Fall 2014.xlsx'";

            com.CommandText = "Select Distinct Registration_No From Student2";
            ////DataTable dt = new DataTable();
            ////da.Fill(dt);

            //Select Distinct Column1 From Student2 where FileName ="BSCS Fall 2014.xlsx";
            com.Connection = con;

            SqlDataAdapter da = new SqlDataAdapter();
            _dtStudent2 = new DataTable();
            da.SelectCommand = com;
            da.Fill(_dtStudent2);


            for (int i = 0; i < _dtStudent2.Rows.Count; i++)
            {
                if (_dtStudent2.Rows[i][0].ToString() == "")
                    continue;

                string Contact1 = "";
                string Contact2 = "";
                string Contact3 = "";
                
                DataRow r = _dtStudent2.Rows[i];

                String RegNo = r[0].ToString();
                com.CommandText = "Select * from Student2 where Registration_No='" + RegNo + "'";

                DataTable dt2 = new DataTable();
                SqlDataAdapter da2 = new SqlDataAdapter();
                da2.SelectCommand = com;
                da2.Fill(dt2);
                da2.Dispose();


                if (dt2.Rows.Count > 0)
                {
                    Contact1 = dt2.Rows[0]["Contact"] == DBNull.Value ? "" : dt2.Rows[0]["Contact"].ToString();
                }
                if (dt2.Rows.Count > 1)
                {
                    Contact2 = dt2.Rows[1]["Contact"] == DBNull.Value ? "" : dt2.Rows[1]["Contact"].ToString();
                }

                if (dt2.Rows.Count > 2)
                {

                    Contact3 = dt2.Rows[2]["Contact"] == DBNull.Value ? "" : dt2.Rows[2]["Contact"].ToString();
                }



                DataRow src = dt2.Rows[0];


                string Serial_No = src[0] == DBNull.Value ? "" : src[0].ToString();
                String File_Name = src[1] == DBNull.Value ? "" : src[1].ToString();
                String Registration_No = src[2] == DBNull.Value ? "" : src[2].ToString();
                String Student_Name = src[3] == DBNull.Value ? "" : src[3].ToString();
                //            String Gender = row["Gender"] == DBNull.Value ? "" : row["Gender"].ToString();
                //            String Student_Contact = row["Student_Contact"] == DBNull.Value ? "" : row["Student_Contact"].ToString();
                String Father_Name = src[4] == DBNull.Value ? "" : src[4].ToString();
                // String Contact = src[8] == DBNull.Value ? "" : src[8].ToString();
                String CNIC = src[6] == DBNull.Value ? "" : src[6].ToString();
                String Birth_Date = src[7] == DBNull.Value ? "" : src[7].ToString();
                String Domicile = src[8] == DBNull.Value ? "" : src[8].ToString();
                String Permanent_Address = src[9] == DBNull.Value ? "" : src[9].ToString();
                string Admission_Date = src[10] == DBNull.Value ? "" : src[10].ToString();
                string Passing_Year = src[11] == DBNull.Value ? "" : src[11].ToString();
                string Remarks = src[12] == DBNull.Value ? "" : src[12].ToString();

                SqlCommand com2 = new SqlCommand();
                com2.Connection = con;
                //com.CommandText=""
                com2.CommandText = "insert into Student3(Serial_No,Registration_No,Student_Name,Father_Name,Contact1,Contact2,Contact3," +
     "CNIC,Birth_Date,Domicile,Permanent_Address,Admission_Date,Passing_Year,Remarks  ) Values("
               + "'" + Serial_No + "','" + Registration_No + "','" + Student_Name + "','" + Father_Name + "','" + Contact1 + "','" + Contact2 + "','" + Contact3 + "','" + CNIC +
               "','" + Birth_Date + "','" + Domicile + "','" + Permanent_Address + "','" + Admission_Date + "','" + Passing_Year +
               "','" + Remarks + "' )";


                int n = com2.ExecuteNonQuery();
            }
            //SqlDataAdapter da = new SqlDataAdapter();
            //DataTable dt = new DataTable();
            //da.Fill(dt);
            con.Close();

            MessageBox.Show("Completed");

        }

        private void dlgOpen_FileOk(object sender, CancelEventArgs e)
        {
            //string filename = this.dlgOpen.FileName;

            //System.IO.FileInfo fi = new System.IO.FileInfo(filename);
            //int fileid = 1;
            //this.ReadFile(fi, fileid);
        }
        public void ReadFile(System.IO.FileInfo fi, int fileID)
        {

            string filename = fi.Name;
            string[] tokens = filename.Split(new char[] { ' ', '.' });
            string program = tokens[0];
            int year = int.Parse(tokens[2]);
            string semester = tokens[1] + " " + tokens[2];

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
            //DataTable dt = ds.Tables[0];

            DataTable dtMaster = null;
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                DataTable dt = ds.Tables[i];

                if (i == 0)
                    dtMaster = dt.Clone();
                for (int r = 2; r < dt.Rows.Count; r++)
                {
                    DataRow row = dtMaster.NewRow();
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        row[c] = dt.Rows[r][c];
                    }
                    dtMaster.Rows.Add(row);
                }
            }

            //this.Invoke(new Handler(delegate ()
            //{
                this.dataGridView1.DataSource = dtMaster;
                
            //}));

            _dtStudent2 = dtMaster;

            _dtStudent2.Columns.Add("Program", typeof(string));
            _dtStudent2.Columns.Add("Semester", typeof(string));
            _dtStudent2.Columns.Add("Year", typeof(string));
            foreach (DataRow row in _dtStudent2.Rows)
            {
                row["Program"] = program;
                row["Semester"] = semester;
                row["Year"] = year;

            }

            this.ds = ds;

            //DataTable dt2 = ds.Tables[1];
            //this.dataGridView2.DataSource = dt2;


        }

        private void button5_Click(object sender, EventArgs e)
        {
            //kis semester ma kon say courses offer kiyaa?
            //or kis semester ko kon sa course offer kiya?


            SqlConnection con = new SqlConnection();
            con.ConnectionString = "Data Source=DESKTOP-ILMCVPK;Initial Catalog=Student_Information_System;Integrated Security=True ";
            SqlCommand com = new SqlCommand();
            com.Connection = con;
            con.Open();
            com.CommandText = ("Select top 8 Semester_Name from Semesters where Semester_ID <= (select MAX(Semester_ID) from Semesters where Active = 1) and Semester_Type_ID != 2 order by Semester_ID Desc");

            

            SqlDataReader rdr = com.ExecuteReader();
            DataTable dtSemesters = new DataTable();
            dtSemesters.Columns.Add("Semester_No");
            dtSemesters.Columns.Add("Semester_Name");
            int i = 1;
            while(rdr.Read())
            {
                DataRow row = dtSemesters.NewRow();
                row[0] = i++;
                row[1] = rdr["Semester_Name"];
                dtSemesters.Rows.Add(row);

            }

            comboBox1.DataSource = dtSemesters;
            comboBox1.DisplayMember = "Semester_No";
            comboBox1.ValueMember = "Semester_Name";


            con.Close();






        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedValue == null)
                return;
            //if (comboBox1.SelectedValue.GetType() == typeof(string))
            //{
                MessageBox.Show(comboBox1.SelectedValue.ToString());


            }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
    }
