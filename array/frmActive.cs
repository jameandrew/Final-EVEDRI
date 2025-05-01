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
    public partial class frmActive : Form
    {
        Getname name;
        private string studname;
        frmLags lags = new frmLags();
        public frmActive(string name)
        {
            InitializeComponent();
            ShowActiveStud("1");
            studname = name;
            btnStuName.Text = name;
        }

        public void ShowActiveStud(string Status)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            DataRow[] i = dt.Select("STATUS = " + Status);

            foreach (DataRow row in i)
            {
                dgvAvtive.Rows.Add
                (
                    row[0].ToString(), row[1].ToString(), row[2].ToString(), row[3].ToString(), row[4].ToString(),
                    row[5].ToString(), row[6].ToString(), row[7].ToString(), row[8].ToString()
                );
            }
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
    }
}
