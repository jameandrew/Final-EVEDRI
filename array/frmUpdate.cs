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
    public partial class frmUpdate : Form
    {
        public frmUpdate()
        {
            InitializeComponent();
        }

        public void Errors()
        {
            if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(txtEmail.Text)
               || string.IsNullOrWhiteSpace(txtPword.Text) || string.IsNullOrWhiteSpace(txtSaying.Text)
               || string.IsNullOrWhiteSpace(txtUname.Text) || string.IsNullOrWhiteSpace(txtpfp.Text))
               
            {
                MessageBox.Show("Please Input the empty fields", "ERROR", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                return;
            }

            if (!rdbMale.Checked && !rdbFemale.Checked)
            {
                MessageBox.Show("Please select a gender", "ERROR", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                return;
            }

            if (!chkBadminton.Checked && !chkBadminton.Checked && !chkVolleyball.Checked)
            {
                MessageBox.Show("Please select your hobby", "ERROR", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            frmActive form = (frmActive)Application.OpenForms["frmActive"];
            frmInactive formInactive = (frmInactive)Application.OpenForms["frmInactive"];
            string hobby = "";
            string gender = "";
            string favColor = "";
            Errors();

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            if (form.dgvAvtive.SelectedRows.Count > 0)
            {
                DataGridViewRow update = form.dgvAvtive.SelectedRows[0];
                
                if (rdbMale.Checked) gender = rdbMale.Text;
                else if (rdbFemale.Checked) gender = rdbFemale.Text;

                List<string> hobbies = new List<string>();
                if (chkVolleyball.Checked) hobbies.Add("Volleyball");
                if (chkBasketball.Checked) hobbies.Add("Basketball");
                if (chkBadminton.Checked) hobbies.Add("Badminton");

                if (cmbColor.SelectedIndex >= 0)
                    favColor = cmbColor.Text;
                int age = DateTime.Now.Year - dtpAge.Value.Year;
                if (DateTime.Now < dtpAge.Value.AddYears(age)) age--;

                update.Cells[0].Value = txtName.Text;
                update.Cells[1].Value = gender;
                update.Cells[2].Value = hobby;
                update.Cells[3].Value = favColor;
                update.Cells[4].Value = txtSaying.Text;
                update.Cells[5].Value = txtUname.Text;
                update.Cells[6].Value = txtPword.Text;

                int row = Convert.ToInt32(lblID.Text) + 2;
                sheet.Range[row, 1].Value = txtName.Text;
                sheet.Range[row, 2].Value = gender;
                sheet.Range[row, 3].Value = hobby;
                sheet.Range[row, 4].Value = favColor;
                sheet.Range[row, 5].Value = txtSaying.Text;
                sheet.Range[row, 6].Value = txtUname.Text;
                sheet.Range[row, 7].Value = txtPword.Text;
                sheet.Range[row, 10].Value = txtEmail.Text;
                sheet.Range[row, 12].Value = txtpfp.Text;   

                book.SaveToFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);
            }

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
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.jfif";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtpfp.Text = dialog.FileName;
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

        private void dtpAge_ValueChanged(object sender, EventArgs e)
        {
            DateTime birthDate = dtpAge.Value;
            int age = CalculateAge(birthDate);
            lblAge.Text = age.ToString();
        }
    }
    
}
