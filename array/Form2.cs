using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace array
{
    public partial class Form2 : Form
    {
        Getname name;
        private string studname;
        public Form2(string name)
        {
            InitializeComponent();
            LOcalFile();
            studname = name;
        }

        public void LOcalFile()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();

            dgvData.DataSource = dt;
        }


        public void Insert(int ID, string name, string gender, string hobbies, string favcolor, string saying)
        {
            int i = dgvData.Rows.Add(ID, name, gender, hobbies, favcolor, saying);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            int row = dgvData.CurrentCell.RowIndex + 2;
            sheet.Range[row, 9].Value = "0";

            book.SaveToFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);

            DataTable dt = sheet.ExportDataTable();
            dgvData.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach(DataGridViewRow search in dgvData.Rows)
            {
                search.Visible = false;

                for (int i = 0;  i < dgvData.Columns.Count; i++)
                {
                    if (search.Cells[i].Value != null && search.Cells[i].Value.ToString().ToLower().Contains(textBox1.Text))
                    {
                        search.Visible = true;
                    }
                }
            }
        }

        private void dgvData_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dgvData_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            frmAddStudent form = new frmAddStudent(studname);
            int r = dgvData.CurrentCell.RowIndex;
            form.lblID.Text = r.ToString();
            form.txtName.Text = dgvData.Rows[r].Cells[0].Value.ToString();
            string gender = dgvData.Rows[r].Cells[1].Value.ToString();
            if(gender == "Male")
            {
                form.rdbMale.Checked = true;
                form.rdbFemale.Checked = false;
            }
            else if(gender == "Female")
            {
                form.rdbFemale.Checked = true;
                form.rdbMale.Checked= false;
            }

            string hobbies = dgvData.Rows[r].Cells[2].Value.ToString();
            string[] hARRAY = hobbies.Split(',');

            form.chkBasketball.Checked = false;
            form.chkVolleyball.Checked = false;
            form.chkBadminton.Checked = false;

            foreach (string s in hARRAY)
            {
                string hobby = s.Trim();

                if (hobby == "Basketball")
                {
                    form.chkBasketball.Checked = true;
                }
                if (hobby == "Volleyball")
                {
                    form.chkVolleyball.Checked = true;
                }
                if (hobby == "Badminton")
                {
                    form.chkBadminton.Checked = true;
                }

            }

            string favcolor = dgvData.Rows[r].Cells[3].Value.ToString();

            if(favcolor == "Blue")
            {
                form.cmbColor.SelectedIndex = 0;
            }
            if(favcolor == "Red")
            {
                form.cmbColor.SelectedIndex = 1;
            }
            if(favcolor == "Pink")
            {
                form.cmbColor.SelectedIndex = 2;
            }
            
            string saying = dgvData.Rows[r].Cells[4].Value.ToString();

            string Username  = dgvData.Rows[r].Cells[5].Value.ToString();

            string Password = dgvData.Rows[r].Cells[6].Value.ToString();

            string Course = dgvData.Rows[r].Cells[7].Value.ToString();

            form.txtSaying.Text = saying;
            form.txtUname.Text = Username;
            form.txtPword.Text = Password;
            form.cmbCourse.Text = Course;

            form.btnSubmit.Visible = false;
            form.btnUpdate.Visible = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
