using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;

namespace array
{
    public partial class frmActive : Form
    {
        Getname name;
        private string studname;
        frmLags lags = new frmLags();
        public frmActive(string name)
        {
            InitializeComponent();
            LoadActiveStudents();
            studname = name;
            btnStuName.Text = name;
        }
        public void LoadActiveStudents()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            DataTable dt = sheet.ExportDataTable();

            DataRow[] activeRows = dt.Select("STATUS = '1'");
            DataTable filtered = dt.Clone(); 

            foreach (DataRow row in activeRows)
            {
                filtered.ImportRow(row);
            }

            dgvAvtive.DataSource = filtered;
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            btnTIme.Text = DateTime.Now.ToString("hh: mm: ss tt");
            btnDate.Text = DateTime.Now.ToString("MM/ dd/ yyyy");
        }

        private void frmActive_Load(object sender, EventArgs e)
        {
            timer1.Start();
            lags.Lags(btnStuName.Text, "Active Student");
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            frmDashboard dashboard = new frmDashboard(studname);
            dashboard.Show();
            this.Hide();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            frmAddStudent student = new frmAddStudent(studname);
            student.Show();
            this.Hide();
        }

        private void btnActSud_Click(object sender, EventArgs e)
        {

        }

        private void btnInactStud_Click(object sender, EventArgs e)
        {
            frmInactive student = new frmInactive(studname);
            student.Show();
            this.Hide();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            frmHistoryLags history = new frmHistoryLags(studname);
            history.Show();
            this.Hide();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            lags.Lags(btnStuName.Text, "Log-Out");
            LogIn logIn = new LogIn();
            logIn.Show();
            this.Hide();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string searchValue = txtSearch.Text.Trim().ToLower();

            if (string.IsNullOrWhiteSpace(searchValue))
            {
                MessageBox.Show("Please enter a search keyword.");
                return;
            }

            // Reload the active students to get the original DataTable
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();

            DataRow[] activeRows = dt.Select("STATUS = '1'");
            DataTable filtered = dt.Clone();

            foreach (DataRow row in activeRows)
            {
                // Check all columns if any cell contains the search term
                if (row.ItemArray.Any(cell => cell != null && cell.ToString().ToLower().Contains(searchValue)))
                {
                    filtered.ImportRow(row);
                }
            }

            dgvAvtive.DataSource = filtered;
        }


        private void dgvAvtive_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            frmUpdate form = new frmUpdate();
            int r = dgvAvtive.CurrentCell.RowIndex;
            form.lblID.Text = r.ToString();
            form.txtName.Text = dgvAvtive.Rows[r].Cells[0].Value.ToString();
            form.txtEmail.Text = dgvAvtive.Rows[r].Cells[10].Value.ToString();
            string gender = dgvAvtive.Rows[r].Cells[1].Value.ToString();
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

            string hobbies = dgvAvtive.Rows[r].Cells[2].Value.ToString();
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

            string favcolor = dgvAvtive.Rows[r].Cells[3].Value.ToString();

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

            string saying = dgvAvtive.Rows[r].Cells[4].Value.ToString();
            string Username = dgvAvtive.Rows[r].Cells[5].Value.ToString();
            string Password = dgvAvtive.Rows[r].Cells[6].Value.ToString();
            string Course = dgvAvtive.Rows[r].Cells[7].Value.ToString();

            form.txtSaying.Text = saying;
            form.txtUname.Text = Username;
            form.txtPword.Text = Password;
            form.cmbCourse.Text = Course;

            form.Show();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you want to delete this student?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
                return;

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            int row = dgvAvtive.CurrentCell.RowIndex + 2; 
            sheet.Range[row, 11].Value = "0"; 

            book.SaveToFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);

            LoadActiveStudents(); 
            MessageBox.Show("Student marked as inactive.");

        }
    }
}
