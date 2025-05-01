using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace array
{
    public class Getname
    {
        Workbook workbook = new Workbook();
        public string Username { get; set; }
        public string Password { get; set; }
        public Getname() { }
        public Getname(string username, string password)
        {
            this.Password = password;
            this.Username = username;
        }

        public bool showName(out string name)
        {
            name = "";
            string picture = "";
            workbook.LoadFromFile(@"C:\Users\HF\Downloads\EVEDRI.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            int row = sheet.Rows.Length;

            for (int i = 2; ; i++)
            {
                var currentUsername = sheet.Range[i, 6].Value;
                var currentPassword = sheet.Range[i, 7].Value;

                if (string.IsNullOrWhiteSpace(currentUsername))
                    break;

                if (Username == currentUsername && Password == currentPassword)
                {
                    name = sheet.Range[i, 1].Value;
                    return true;
                }
            }


            MessageBox.Show("Invalid username or password", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
        public static bool EmailValid(string Email)
        {
            string Emailadd = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
            return Regex.IsMatch(Email, Emailadd);
        }
    }
}
