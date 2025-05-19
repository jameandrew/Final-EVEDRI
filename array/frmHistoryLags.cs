using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace array
{
    public partial class frmHistoryLags : Form
    {
        Getname name;
        private string studname;
        frmLags lags = new frmLags();
        public frmHistoryLags(string name)
        {
            InitializeComponent();
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

        private void frmHistoryLags_Load(object sender, EventArgs e)
        {
            timer1.Start();
            lags.Lags(btnStuName.Text, "Lags");
            frmLags frmLags = new frmLags();
            frmLags.showLogs(dgvLags);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            btnTIme.Text = DateTime.Now.ToString("hh: mm: ss tt");
            btnDate.Text = DateTime.Now.ToString("MM/ dd/ yyyy");
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            lags.Lags(btnStuName.Text, "Log-Out");
            LogIn logIn = new LogIn();
            logIn.Show();
            this.Hide();
        }

        private void btnInactStud_Click(object sender, EventArgs e)
        {
            frmInactive inactive = new frmInactive(studname);
            inactive.Show();
            this.Hide();
        }

        private void btnActSud_Click(object sender, EventArgs e)
        {
            frmActive active = new frmActive(studname);
            active.Show();
            this.Hide();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            frmAddStudent student = new frmAddStudent(studname);
            student.Show();
            this.Hide();
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            frmDashboard student = new frmDashboard(studname);
            student.Show();
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

            foreach (DataGridViewRow row in dgvLags.Rows)
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
    }
}
