using Guna.UI2.AnimatorNS;
using Spire.Xls;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace array
{
    public partial class frmUpdate : Form
    {
        Getname name;
        private string studname;
        frmLags lags = new frmLags();

        public frmUpdate(string name)
        {
            InitializeComponent();
            studname = name;
            btnStuName.Text = studname;
            if (!string.IsNullOrEmpty(Getname.ProfileImagePath) && File.Exists(Getname.ProfileImagePath))
            {
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                pictureBox1.Image = Image.FromFile(Getname.ProfileImagePath);
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

            // Fixed your hobby check to include all three checkboxes properly
            if (!chkBadminton.Checked && !chkBasketball.Checked && !chkVolleyball.Checked)
            {
                MessageBox.Show("Please select your hobby", "ERROR", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
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

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            if (!Errors()) return;

            lags.Lags(btnStuName.Text, "Updated a student");

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            frmActive form = (frmActive)Application.OpenForms["frmActive"];
            int r = form.dgvAvtive.CurrentCell.RowIndex;

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
                form.dgvAvtive.Rows[r].Cells[0].Value = txtName.Text;
            }

            if (rdbFemale.Checked == true)
            {
                form.dgvAvtive.Rows[r].Cells[1].Value = rdbFemale.Text;
                gender = rdbFemale.Text;
            }
            if (rdbMale.Checked == true)
            {
                form.dgvAvtive.Rows[r].Cells[1].Value = rdbMale.Text;
                gender = rdbMale.Text;
            }

            string hobbies = "";
            if (chkBasketball.Checked == true) hobbies += chkBasketball.Text + ", ";
            if (chkVolleyball.Checked == true) hobbies += chkVolleyball.Text + ", ";
            if (chkBadminton.Checked == true) hobbies += chkBadminton.Text + ", ";
            hobbies = hobbies.TrimEnd(' ', ',');
            form.dgvAvtive.Rows[r].Cells[2].Value = hobbies;

            if (cmbColor.SelectedItem != null) form.dgvAvtive.Rows[r].Cells[3].Value = cmbColor.SelectedItem.ToString();
            if (cmbCourse.SelectedItem != null) form.dgvAvtive.Rows[r].Cells[7].Value = cmbCourse.SelectedItem.ToString();
            if (!string.IsNullOrEmpty(txtSaying.Text)) form.dgvAvtive.Rows[r].Cells[4].Value = txtSaying.Text;

            int age = DateTime.Now.Year - dtpAge.Value.Year;
            if (DateTime.Now < dtpAge.Value.AddYears(age)) age--;
            form.dgvAvtive.Rows[r].Cells[8].Value = age;

            string imagepath = txtpfp.Text;
            if (!string.IsNullOrEmpty(imagepath)) form.dgvAvtive.Rows[r].Cells[11].Value = imagepath;

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
            sheet.Range[rowIndex, 10].Value = "1";
            sheet.Range[rowIndex, 11].Value = txtEmail.Text;
            sheet.Range[rowIndex, 12].Value = imagepath;

            book.SaveToFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);

            DataTable dt = sheet.ExportDataTable();
            form.dgvAvtive.DataSource = dt;
            form.dgvAvtive.Refresh();

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

        private void guna2Button1_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            btnTIme.Text = DateTime.Now.ToString("hh:mm:ss tt");
            btnDate.Text = DateTime.Now.ToString("MM/dd/yyyy");
        }

        private void frmUpdate_Load_1(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void guna2Button4_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.jfif";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtpfp.Text = dialog.FileName;
            }
        }

        private void dtpAge_ValueChanged_1(object sender, EventArgs e)
        {
            int age = CalculateAge(dtpAge.Value);
            lblAge.Text = age.ToString();
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
    }
}
