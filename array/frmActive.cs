using Spire.Xls;
using System;
using System.Data;
using System.Drawing;
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

            if (!string.IsNullOrEmpty(Getname.ProfileImagePath) && System.IO.File.Exists(Getname.ProfileImagePath))
            {
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                pictureBox1.Image = Image.FromFile(Getname.ProfileImagePath);
            }
            else
            {
                pictureBox1.Image = null;
            }
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
            btnTIme.Text = DateTime.Now.ToString("hh:mm:ss tt");
            btnDate.Text = DateTime.Now.ToString("MM/dd/yyyy");
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
            frmActive student = new frmActive(studname);
            student.Show();
            this.Hide();
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

            foreach (DataGridViewRow row in dgvAvtive.Rows)
            {
                bool matchFound = false;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value.ToString().ToLower().Contains(searchValue))
                    {
                        matchFound = true;
                        break;
                    }
                }

                row.DefaultCellStyle.BackColor = matchFound ? Color.Yellow : Color.White;
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            lags.Lags(btnStuName.Text, "Deleted a student");

            if (dgvAvtive.CurrentCell == null)
            {
                MessageBox.Show("Please select a student to delete.", "No selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult result = MessageBox.Show("Do you want to delete this student?", "Confirm Restore", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
                return;

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            string selectedName = dgvAvtive.CurrentRow.Cells[0].Value.ToString();

            int excelRow = -1;
            for (int i = 2; i <= sheet.LastRow; i++)
            {
                if (sheet.Range[i, 1].Value == selectedName)
                {
                    excelRow = i;
                    break;
                }
            }

            if (excelRow == -1)
            {
                MessageBox.Show("Student not found in Excel.");
                return;
            }

            sheet.Range[excelRow, 10].Value = "0";

            book.SaveToFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);

            LoadActiveStudents();

            dgvAvtive.ClearSelection();
            if (dgvAvtive.Rows.Count > 0)
            {
                dgvAvtive.CurrentCell = dgvAvtive.Rows[0].Cells[0];
                dgvAvtive.Rows[0].Selected = true;
            }

            MessageBox.Show("Student marked as inactive.");
        }
        private void dgvAvtive_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            frmUpdate form = new frmUpdate(studname);
            int r = dgvAvtive.CurrentCell.RowIndex;
            form.lblID.Text = r.ToString();
            form.txtName.Text = dgvAvtive.Rows[r].Cells[0].Value.ToString();
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
            string Age = dgvAvtive.Rows[r].Cells[8].Value.ToString();
            string EMail = dgvAvtive.Rows[r].Cells[10].Value.ToString();
            string pict = dgvAvtive.Rows[r].Cells[11].Value.ToString();

            form.txtSaying.Text = saying;
            form.txtUname.Text = Username;
            form.txtPword.Text = Password;
            form.cmbCourse.Text = Course;
            form.lblAge.Text = Age;
            form.txtEmail.Text = EMail;
            form.txtpfp.Text = pict;

            form.Show();
        }

    }
}
