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
            LoadInActiveStudents();
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

        public void LoadInActiveStudents()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            DataTable dt = sheet.ExportDataTable();

            DataRow[] activeRows = dt.Select("STATUS = '0'");
            DataTable filtered = dt.Clone(); 

            foreach (DataRow row in activeRows)
            {
                filtered.ImportRow(row);
            }

            dgvActive.DataSource = filtered;
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
            logIn.Show();
            this.Hide();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
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

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            frmHistoryLags history = new frmHistoryLags(studname);
            history.Show();
            this.Hide();
        }

        private void dgvAvtive_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            frmUpdateInActive form = new frmUpdateInActive();
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
            string Age = dgvActive.Rows[r].Cells[8].Value.ToString();
            string EMail = dgvActive.Rows[r].Cells[10].Value.ToString();
            string pict = dgvActive.Rows[r].Cells[11].Value.ToString();

            form.txtSaying.Text = saying;
            form.txtUname.Text = Username;
            form.txtPword.Text = Password;
            form.cmbCourse.Text = Course;         
            form.lblAge.Text = Age;
            form.txtEmail.Text = EMail;
            form.txtpfp.Text = pict;

            form.Show();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string searchValue = txtSearch.Text.Trim().ToLower();

            if (string.IsNullOrWhiteSpace(searchValue))
            {
                MessageBox.Show("Please enter a search keyword.");
                return;
            }
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();

            DataRow[] activeRows = dt.Select("STATUS = '1'");
            DataTable filtered = dt.Clone();

            foreach (DataRow row in activeRows)
            {
                if (row.ItemArray.Any(cell => cell != null && cell.ToString().ToLower().Contains(searchValue)))
                {
                    filtered.ImportRow(row);
                }
            }

            dgvActive.DataSource = filtered;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvActive.Rows.Count <= 1)
            {
                MessageBox.Show("You cannot restore the last remaining student.", "Action Not Allowed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult result = MessageBox.Show("Do you want to restore this student?", "Confirm Restore", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
                return;

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            int row = dgvActive.CurrentCell.RowIndex + 2;
            sheet.Range[row, 11].Value = "1";

            book.SaveToFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);

            LoadInActiveStudents();
            MessageBox.Show("Student marked as active.");
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
