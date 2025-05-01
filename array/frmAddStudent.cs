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
    public partial class frmAddStudent : Form
    {
        Getname name;
        private string studname;
        frmInactive inactive;
        frmLags lags = new frmLags();
        string[] Student = new string[5];
        int i = 0;
        string data = "";
        string gender = "";
        string hobbies = "";
        string favorite = "";
        string saying = "";
        string Username = "";
        string Password = "";
        string Course = "";
        public frmAddStudent(string name)
        {
            InitializeComponent();
            studname = name;
            btnStuName.Text = name;
            dateTimePicker1_ValueChanged(null, null);
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            frmDashboard frmDashboard = new frmDashboard(studname);
            frmDashboard.Show();
            this.Hide();
        }

        public void Errors()
        {
            if(string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(txtEmail.Text)
               || string.IsNullOrWhiteSpace(txtPword.Text) || string.IsNullOrWhiteSpace(txtSaying.Text)
               || string.IsNullOrWhiteSpace(txtUname.Text) || string.IsNullOrWhiteSpace(txtPfp.Text) || string.IsNullOrWhiteSpace(txtPfp.Text))
            {
                MessageBox.Show("Please Input the empty fields", "ERROR", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                return;
            }

            if(!rdbMale.Checked && !rdbFemale.Checked)
            {
                MessageBox.Show("Please select a gender", "ERROR", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                return;
            }

            if(!chkBadminton.Checked && !chkBadminton.Checked && !chkVolleyball.Checked)
            {
                MessageBox.Show("Please select your hobby", "ERROR", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            lags.Lags(btnStuName.Text, "Add a student");
            Form2 frm = new Form2(studname);
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];
            data = "";
            gender = "";
            hobbies = "";
            favorite = "";
            saying = "";
            Username = "";
            Password = "";
            Course = "";

            if (!Getname.EmailValid(txtEmail.Text))
            {
                MessageBox.Show("Please input a valid email (e.g., user@example.com)", "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                return;
            }

            if (!string.IsNullOrEmpty(txtName.Text))
            {
                data += txtName.Text;
            }
            if (rdbFemale.Checked)
            {
                gender += rdbFemale.Text;
            }
            else if (rdbMale.Checked)
            {
                gender += rdbMale.Text;
            }
            if (chkBadminton.Checked)
            {
                hobbies += chkBadminton.Text + ", ";
            }
            if (chkBasketball.Checked)
            {
                hobbies += chkBasketball.Text + " , ";
            }
            if (chkVolleyball.Checked)
            {
                hobbies += chkVolleyball.Text + " , ";
            }
            if (cmbColor.SelectedIndex == 0)
            {
                favorite += cmbColor.Text;
            }
            if (cmbColor.SelectedIndex == 1)
            {
                favorite += cmbColor.Text;
            }
            if (cmbColor.SelectedIndex == 2)
            {
                favorite += cmbColor.Text;
            }
            if (!string.IsNullOrEmpty(txtSaying.Text))
            {
                saying += txtSaying.Text;
            }
            if (!string.IsNullOrEmpty(txtUname.Text))
            {
                Username = txtUname.Text;
            }
            if (!string.IsNullOrEmpty(txtPword.Text))
            {
                Password = txtPword.Text;
            }
            if (cmbCourse.SelectedIndex == 1)
            {
                Course += cmbCourse.Text;
            }
            else if (cmbCourse.SelectedIndex == 2)
            {
                Course += cmbCourse.Text;
            }
            else if (cmbCourse.SelectedIndex == 3)
            {
                Course += cmbCourse.Text;
            }

            int row = sheet.Rows.Length + 1;

            sheet.Range[row, 1].Value = txtName.Text;
            sheet.Range[row, 2].Value = gender;
            sheet.Range[row, 3].Value = hobbies;
            sheet.Range[row, 4].Value = favorite;
            sheet.Range[row, 5].Value = saying;
            sheet.Range[row, 6].Value = txtUname.Text;
            sheet.Range[row, 7].Value = txtPword.Text;
            sheet.Range[row, 8].Value = cmbCourse.Text;
            sheet.Range[row, 9].Value = "1";
            book.SaveToFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);
            DataTable dt = sheet.ExportDataTable();
            frm.dgvData.DataSource = dt;

            txtName.Clear();
            rdbFemale.Checked = false;
            rdbMale.Checked = false;
            chkBadminton.Checked = false;
            chkBasketball.Checked = false;
            chkVolleyball.Checked = false;
            cmbColor.SelectedIndex = -1;
            txtSaying.Clear();
            txtUname.Clear();
            txtPword.Clear();
            cmbCourse.SelectedIndex = -1;
        }

        private void btnDisplayAll_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2(studname);
            form2.Show();
        }

        private void btnActSud_Click(object sender, EventArgs e)
        {
            frmActive frmActive = new frmActive(studname);

            frmActive.Show();
            this.Hide();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            frmHistoryLags frmHistoryLags = new frmHistoryLags(studname);
            frmHistoryLags.Show();
            this.Hide();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            btnTIme.Text = DateTime.Now.ToString("hh: mm: ss tt");
            btnDate.Text = DateTime.Now.ToString("MM/ dd/ yyyy");
        }

        private void frmAddStudent_Load(object sender, EventArgs e)
        {
            timer1.Start();
            lags.Lags(btnStuName.Text, "Add Student");
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            lags.Lags(btnStuName.Text, "Log-Out");
            LogIn logIn = new LogIn();
            logIn.ShowDialog();
            this.Hide();
        }

        private void btnInactStud_Click(object sender, EventArgs e)
        {
            frmInactive inactive = new frmInactive(studname);
            inactive.Show();
            this.Hide();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if(dialog.ShowDialog() == DialogResult.OK)
            {
                txtPfp.Text = dialog.FileName;
            }
        }

        private int CalculateAge(DateTime birthDate)
        {
            DateTime today = DateTime.Today;
            int age = today.Year - birthDate.Year;

            if (birthDate.Date > today.AddYears(-age))
            {
                age--;
            }

            return age;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime birthDate = dtpAge.Value;
            int age = CalculateAge(birthDate);
            lblAge.Text = age.ToString();
        }
    }
}
