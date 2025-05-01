using Guna.UI2.AnimatorNS;
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
    public partial class frmInactive : Form
    {
        Getname name;
        private string studname;
        frmLags lags = new frmLags();
        public frmInactive(string name)
        {
            InitializeComponent();
            ShowInactiveStud("0");
            studname = name;
            btnStuName.Text = name;
        }

        public void ShowInactiveStud(string Status)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            DataRow[] i = dt.Select("STATUS = " + Status);

            foreach (DataRow row in i)
            {
                dgvActive.Rows.Add
                (
                    row[0].ToString(), row[1].ToString(), row[2].ToString(), row[3].ToString(), row[4].ToString(),
                    row[5].ToString(), row[6].ToString(), row[7].ToString(), row[8].ToString()
                );
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            btnTIme.Text = DateTime.Now.ToString("hh: mm: ss tt");
            btnDate.Text = DateTime.Now.ToString("MM/ dd/ yyyy");
        }

        private void frmInactive_Load(object sender, EventArgs e)
        {
            timer1.Start();
            lags.Lags(btnStuName.Text, "Inactive Student");
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            lags.Lags(btnStuName.Text, "Log-Out");
            LogIn logIn = new LogIn();
            logIn.ShowDialog();
            this.Hide();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            frmDashboard dashboard = new frmDashboard(studname);
            dashboard.ShowDialog();
            this.Hide();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            frmAddStudent student = new frmAddStudent(studname);
            student.ShowDialog();
            this.Hide();
        }

        private void btnActSud_Click(object sender, EventArgs e)
        {
            frmActive student = new frmActive(studname);
            student.ShowDialog();
            this.Hide();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            frmHistoryLags history = new frmHistoryLags(studname);
            history.ShowDialog();
            this.Hide();
        }

        private void dgvAvtive_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            frmAddStudent form = new frmAddStudent(studname);
            int r = dgvActive.CurrentCell.RowIndex;
            form.lblID.Text = r.ToString();
            form.txtName.Text = dgvActive.Rows[r].Cells[0].Value.ToString();
            string gender = dgvActive.Rows[r].Cells[1].Value.ToString();
            if (gender == "Male")
            {
                form.rdbMale.Checked = true;
                form.rdbFemale.Checked = false;
            }
            else if (gender == "Female")
            {
                form.rdbFemale.Checked = true;
                form.rdbMale.Checked = false;
            }

            string hobbies = dgvActive.Rows[r].Cells[2].Value.ToString();
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

            string favcolor = dgvActive.Rows[r].Cells[3].Value.ToString();

            if (favcolor == "Blue")
            {
                form.cmbColor.SelectedIndex = 0;
            }
            if (favcolor == "Red")
            {
                form.cmbColor.SelectedIndex = 1;
            }
            if (favcolor == "Pink")
            {
                form.cmbColor.SelectedIndex = 2;
            }

            string saying = dgvActive.Rows[r].Cells[4].Value.ToString();
            string Username = dgvActive.Rows[r].Cells[5].Value.ToString();
            string Password = dgvActive.Rows[r].Cells[6].Value.ToString();
            string Course = dgvActive.Rows[r].Cells[7].Value.ToString();

            form.txtSaying.Text = saying;
            form.txtUname.Text = Username;
            form.txtPword.Text = Password;
            form.cmbCourse.Text = Course;

            form.btnSubmit.Visible = false;
            form.btnUpdate.Visible = true;

            form.Show();
            this.Hide();
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            int row = dgvActive.CurrentCell.RowIndex + 2;
            sheet.Range[row, 9].Value = "0";

            book.SaveToFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);

            DataTable dt = sheet.ExportDataTable();
            dgvActive.DataSource = dt;
        }

        //private void guna2Button4_Click(object sender, EventArgs e)
        //{
        //    Workbook book = new Workbook();
        //    book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");

        //    Worksheet sheet = book.Worksheets[0];

        //    frmAddStudent addstudent = (frmAddStudent)Application.OpenForms["frmAddStudent"];
        //    int r = dgvActive.CurrentCell.RowIndex;

        //    if(!string.IsNullOrWhiteSpace(addstudent.txtName.Text))
        //    {
        //        dgvActive.Rows[r].Cells[0].Value = addstudent.txtName.Text;
        //    }

        //    string gender = "";
        //    if(addstudent.rdbMale.Checked == true)
        //    {
        //        dgvActive.Rows[r].Cells[1].Value = addstudent.rdbMale.Text;
        //    }
        //    if(addstudent.rdbFemale.Checked == true)
        //    {
        //        dgvActive.Rows[r].Cells[1].Value = addstudent.rdbFemale.Text;
        //    }

        //    string hobbies = "";
        //    if (addstudent.chkBasketball.Checked == true) hobbies += addstudent.chkBasketball.Text + " , ";
        //    if (addstudent.chkVolleyball.Checked == true) hobbies += addstudent.chkVolleyball.Text + " , ";
        //    if (addstudent.chkBadminton.Checked == true) hobbies += addstudent.chkBadminton.Text + " , ";
        //    dgvActive.Rows[r].Cells[2].Value = hobbies.Trim();

        //    if (addstudent.cmbColor.SelectedIndex != null) dgvActive.Rows[r].Cells[3].Value = addstudent.cmbColor.Text;
        //    if (addstudent.cmbCourse.SelectedIndex != null) dgvActive.Rows[r].Cells[7].Value = addstudent.cmbCourse.Text;
        //    if (!string.IsNullOrWhiteSpace(addstudent.txtSaying.Text)) dgvActive.Rows[r].Cells[4].Value = addstudent.txtSaying.Text;

        //    if(!string.IsNullOrWhiteSpace(addstudent.txtUname.Text)) dgvActive.Rows[r].Cells[5].Value = addstudent.txtUname.Text;
        //    if(!string.IsNullOrWhiteSpace(addstudent.txtPword.Text)) dgvActive.Rows[r].Cells[6].Value = addstudent.txtPword.Text;

        //    int row = (Convert.ToInt32(addstudent.lblID.Text)) + 2;

        //    sheet.Range[row, 1].Value = addstudent.txtName.Text;
        //    sheet.Range[row, 2].Value = gender;
        //    sheet.Range[row, 3].Value = hobbies;
        //    sheet.Range[row, 4].Value = addstudent.cmbColor.Text;
        //    sheet.Range[row, 5].Value = addstudent.txtSaying.Text;
        //    sheet.Range[row, 6].Value = addstudent.txtUname.Text;
        //    sheet.Range[row, 7].Value = addstudent.txtPword.Text;
        //    sheet.Range[row, 8].Value = addstudent.cmbCourse.Text;
        //}

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string searchValue = txtSearch.Text.Trim().ToLower();

            if (string.IsNullOrWhiteSpace(searchValue))
            {
                MessageBox.Show("Please enter a search keyword.");
                return;
            }

            foreach (DataGridViewRow row in dgvActive.Rows)
            {
                row.Visible = false;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value.ToString().ToLower().Contains(searchValue))
                    {
                        row.Visible = true;
                        break;
                    }
                }
            }
        }
    }
}
