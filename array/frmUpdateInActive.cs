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
using DrawingImage = System.Drawing.Image;
using System.Windows.Forms;

namespace array
{
    public partial class frmUpdateInActive : Form
    {
        Getname name;
        private string studname;
        frmLags lags = new frmLags();
        public frmUpdateInActive(string name)
        {
            InitializeComponent();
            studname = name;
            btnStuName.Text = studname;
            if (!string.IsNullOrEmpty(Getname.ProfileImagePath) && System.IO.File.Exists(Getname.ProfileImagePath))
            {
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                pictureBox1.Image = DrawingImage.FromFile(Getname.ProfileImagePath);
            }
            else
            {
                pictureBox1.Image = null;
            }
        }

        public bool Errors()
        {
            if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(txtEmail.Text)
               || string.IsNullOrWhiteSpace(txtPword.Text) || string.IsNullOrWhiteSpace(txtSaying.Text)
               || string.IsNullOrWhiteSpace(txtUname.Text) || string.IsNullOrWhiteSpace(txtpfp.Text))
            {
                MessageBox.Show("Please Input the empty fields", "ERROR", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                return false;
            }

            if (!rdbMale.Checked && !rdbFemale.Checked)
            {
                MessageBox.Show("Please select a gender", "ERROR", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                return false;
            }

            if (!chkBadminton.Checked && !chkBadminton.Checked && !chkVolleyball.Checked)
            {
                MessageBox.Show("Please select your hobby", "ERROR", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
        }

        private void btnSubmit_Click_1(object sender, EventArgs e)
        {
            if (!Errors()) return;

            lags.Lags(btnStuName.Text, "Updated a student");

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            frmInactive form = (frmInactive)Application.OpenForms["frmInactive"];
            int r = form.dgvActive.CurrentCell.RowIndex;

            bool isUsernameTaken = false;
            bool isPasswordTaken = false;

            for (int i = 2; i <= sheet.LastRow; i++)
            {
                string existingUsername = sheet.Range[i, 6]?.Value?.Trim() ?? "";
                string existingPassword = sheet.Range[i, 7]?.Value?.Trim() ?? "";

                if (txtUname.Text.Trim().Equals(existingUsername, StringComparison.OrdinalIgnoreCase))
                {
                    isUsernameTaken = true;
                }

                if (txtPword.Text.Trim().Equals(existingPassword, StringComparison.Ordinal))
                {
                    isPasswordTaken = true;
                }

                if (isUsernameTaken || isPasswordTaken)
                    break;
            }

            if (isUsernameTaken && isPasswordTaken)
            {
                MessageBox.Show("Both the username and password are already taken. Please choose different credentials.");
                return;
            }
            else if (isUsernameTaken)
            {
                MessageBox.Show("The username is already taken. Please choose another one.");
                return;
            }
            else if (isPasswordTaken)
            {
                MessageBox.Show("The password is already taken. Please choose a different one.");
                return;
            }


            if (!Getname.EmailValid(txtEmail.Text))
            {
                MessageBox.Show("Please input a valid email (e.g., user@example.com)", "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                return;
            }

            string name = txtName.Text;
            string gender = rdbFemale.Checked ? rdbFemale.Text : (rdbMale.Checked ? rdbMale.Text : "");
            {
                form.dgvActive.Rows[r].Cells[0].Value = txtName.Text;
            }

            if (rdbFemale.Checked == true)
            {
                form.dgvActive.Rows[r].Cells[1].Value = rdbFemale.Text;
                gender = rdbFemale.Text;
            }
            if (rdbMale.Checked == true)
            {
                form.dgvActive.Rows[r].Cells[1].Value = rdbMale.Text;
                gender = rdbMale.Text;
            }

            string hobbies = "";
            if (chkBasketball.Checked == true) hobbies += chkBasketball.Text + ", ";
            if (chkVolleyball.Checked == true) hobbies += chkVolleyball.Text + ", ";
            if (chkBadminton.Checked == true) hobbies += chkBadminton.Text + ", ";
            hobbies = hobbies.TrimEnd(' ', ',');
            form.dgvActive.Rows[r].Cells[2].Value = hobbies;

            if (cmbColor.SelectedItem != null) form.dgvActive.Rows[r].Cells[3].Value = cmbColor.SelectedItem.ToString();
            if (cmbCourse.SelectedItem != null) form.dgvActive.Rows[r].Cells[7].Value = cmbCourse.SelectedItem.ToString();
            if (!string.IsNullOrEmpty(txtSaying.Text)) form.dgvActive.Rows[r].Cells[4].Value = txtSaying.Text;

            int age = DateTime.Now.Year - dtpAge.Value.Year;
            if (DateTime.Now < dtpAge.Value.AddYears(age)) age--;
            form.dgvActive.Rows[r].Cells[8].Value = age;

            string imagepath = txtpfp.Text;
            if (!string.IsNullOrEmpty(imagepath)) form.dgvActive.Rows[r].Cells[11].Value = imagepath;

            int rowIndex = (Convert.ToInt32(lblID.Text)) + 2;

            sheet.Range[rowIndex, 1].Value = txtName.Text;
            sheet.Range[rowIndex, 2].Value = gender;
            sheet.Range[rowIndex, 3].Value = hobbies;
            sheet.Range[rowIndex, 4].Value = cmbColor.Text;
            sheet.Range[rowIndex, 5].Value = txtSaying.Text;
            sheet.Range[rowIndex, 6].Value = txtUname.Text;
            sheet.Range[rowIndex, 7].Value = txtPword.Text;
            sheet.Range[rowIndex, 8].Value = cmbCourse.Text;
            sheet.Range[rowIndex, 9].Value = age.ToString();
            sheet.Range[rowIndex, 10].Value = "0";
            sheet.Range[rowIndex, 11].Value = txtEmail.Text;
            sheet.Range[rowIndex, 12].Value = imagepath;

            book.SaveToFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);

            DataTable dt = sheet.ExportDataTable();
            form.dgvActive.DataSource = dt;
            form.dgvActive.Refresh();

            MessageBox.Show("Student updated successfully!");
            ClearInputs();
        }
        private void ClearInputs()
        {
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
            txtEmail.Clear();
            cmbCourse.SelectedIndex = -1;
            txtpfp.Clear();
            lblAge.Text = "";
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.jfif";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtpfp.Text = dialog.FileName;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            btnTIme.Text = DateTime.Now.ToString("hh: mm: ss tt");
            btnDate.Text = DateTime.Now.ToString("MM/ dd/ yyyy");
        }

        private void frmUpdateInActive_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
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
            txtEmail.Clear();
            cmbCourse.SelectedIndex = -1;
            txtpfp.Clear();
            lblAge.Text = "";
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

        private void dtpAge_ValueChanged(object sender, EventArgs e)
        {
            DateTime birthDate = dtpAge.Value;
            int age = CalculateAge(birthDate);
            lblAge.Text = age.ToString();
        }
    }
}
