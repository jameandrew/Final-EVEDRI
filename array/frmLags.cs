using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace array
{
    public class frmLags
    {
        Workbook workbook = new Workbook();

        public void Lags(string user, string message)
        {
            workbook.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = workbook.Worksheets[1];
            int row = sheet.Rows.Length + 1;
            sheet.Range[row,1].Value = user;
            sheet.Range[row,2].Value = message;
            sheet.Range[row,3].Value = DateTime.Now.ToString("MM/dd/yyyy");
            sheet.Range[row,4].Value = DateTime.Now.ToString("hh:mm:ss:tt");
            workbook.SaveToFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);
        }

        public void showLogs(DataGridView dgv)
        {
            workbook.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = workbook.Worksheets[1];
            DataTable dt = sheet.ExportDataTable();
            dgv.DataSource = dt;
        }
    }
}
