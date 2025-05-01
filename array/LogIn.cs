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
    public partial class LogIn : Form
    {
        frmLags lags = new frmLags();
        public LogIn()
        {
            InitializeComponent();
            this.AcceptButton = btnLOgin;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (txtPassword.UseSystemPasswordChar)
            {
                pictureBox1.Image = Image.FromFile("C:\\Users\\HF\\Downloads\\hide.png");
                txtPassword.UseSystemPasswordChar = false;
                pictureBox1.Text = "Show";
            }
            else
            {
                pictureBox1.Image = Image.FromFile("C:\\Users\\HF\\Downloads\\visibility.png");
                txtPassword.UseSystemPasswordChar = true;
                pictureBox1.Text = "Hide";
            }

            pictureBox1.Refresh();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnLOgin_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text) || string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                MessageBox.Show("Please input all fields", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            Getname nameGetter = new Getname(txtUsername.Text, txtPassword.Text);
            string studentName;

            bool isValid = nameGetter.showName(out studentName);

            if (isValid)
            {
                lags.Lags(txtUsername.Text, "Logged in");

                frmDashboard dashboard = new frmDashboard(studentName);
                dashboard.Show();
                this.Hide();
            }
        }

    }
}
