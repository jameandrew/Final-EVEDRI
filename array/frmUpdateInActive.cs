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
    public partial class frmUpdateInActive : Form
    {
        public frmUpdateInActive()
        {
            InitializeComponent();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            frmInactive form = (frmInactive)Application.OpenForms["frmInActive"];
            frmInactive formInactive = (frmInactive)Application.OpenForms["frmInactive"];
            string hobby = "";
            string gender = "";
            string favColor = "";

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            if (form.dgvActive.SelectedRows.Count > 0)
            {
                DataGridViewRow update = form.dgvActive.SelectedRows[0];

                if (rdbMale.Checked) gender = rdbMale.Text;
                else if (rdbFemale.Checked) gender = rdbFemale.Text;

                List<string> hobbies = new List<string>();
                if (chkVolleyball.Checked) hobbies.Add("Volleyball");
                if (chkBasketball.Checked) hobbies.Add("Basketball");
                if (chkBadminton.Checked) hobbies.Add("Badminton");
                hobby = string.Join(", ", hobbies);

                if (cmbColor.SelectedIndex >= 0)
                    favColor = cmbColor.Text;

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
    }
}
