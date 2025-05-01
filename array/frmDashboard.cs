using Guna.UI2.AnimatorNS;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace array
{
    public partial class frmDashboard : Form
    {
        Getname name;
        private string studname;
        frmLags lags = new frmLags();
        Workbook workbook = new Workbook();
        public frmDashboard(string name)
        {
            InitializeComponent();
            lblTotalAct.Text = ShowTotal(9, "1").ToString();
            lblTotalInact.Text = ShowTotal(9, "0").ToString();
            lblTotalMale.Text = ShowTotal(2, "Male").ToString();
            lblTotalFemale.Text = ShowTotal(2, "Female").ToString();
            lblRed.Text = ShowTotal(4, "Red").ToString();
            lblBlue.Text = ShowTotal(4, "Blue").ToString();
            lblPink.Text = ShowTotal(4, "Pink").ToString();
            lblBasket.Text = ShowTotal(3, "Basketball").ToString();
            lblVolley.Text = ShowTotal(3, "Volleyball").ToString();
            lblBadminton.Text = ShowTotal(3, "Badminton").ToString();
            lblBSIT.Text = ShowTotal(8, "BSIT").ToString();
            lblBSCS.Text = ShowTotal(8, "BSCS").ToString();
            lblBSFM.Text = ShowTotal(8, "BSFM").ToString();
            studname = name;
            btnStuName.Text = name;
        }

        public int ShowTotal(int count, string values)
        {
            workbook.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            int row = sheet.Rows.Length;

            int total = 0;

            for(int i = 2; i <= row; i++)
            {
                if (sheet.Range[i, count].Value == values)
                {
                    total++;
                }
            }
            return total;
        }

        public void ShowName(string Name)
        {
            btnStuName.Text = Name;
        }

        private void timer1_Tick(object sender, EventArgs e) 
        {
            btnTIme.Text = DateTime.Now.ToString("hh: mm: ss tt");
            btnDate.Text = DateTime.Now.ToString("MM/ dd/ yyyy");
        }
       
        private void frmDashboard_Load(object sender, EventArgs e)
        {
            timer1.Start();
            lags.Lags(btnStuName.Text,"Dashboard");
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            frmAddStudent frmAddStudent = new frmAddStudent(studname);
            frmAddStudent.Show();
            this.Hide();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            frmHistoryLags historyForm = new frmHistoryLags(studname);
            historyForm.Show();
            this.Hide();
        }

        private void btnActSud_Click(object sender, EventArgs e)
        {
            frmActive ActStudent = new frmActive(studname);
            ActStudent.Show();
            this.Hide();
        }

        private void btnInactStud_Click(object sender, EventArgs e)
        {
            frmInactive inactive = new frmInactive(studname);
            inactive.Show();
            this.Hide();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            lags.Lags(btnStuName.Text, "Log-Out");
            LogIn logIn = new LogIn();
            logIn.ShowDialog();
            this.Hide();
        }
    }
}
