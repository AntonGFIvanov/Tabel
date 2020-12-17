using System;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Tabel
{
    public partial class FormStart : Form
    {
        public FormStart()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start("d:\\TABEL\\COPY\\tab.bat");
            FormGeneral f = new FormGeneral();
            f.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
                FormCorrection f = new FormCorrection();
                f.ShowDialog();
            
        }

        private void FormStart_Load(object sender, EventArgs e)
        {
            label1.Text = "ver. " + Assembly.GetExecutingAssembly().GetName().Version.ToString();
            label1.Text = "ver.   4.0.8.1" ;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FormGeneral.CopyFile("\\\\Fsmtm1\\Install\\!!!Programs!!!\\Pull\\TABEL\\COPY\\tab.bat", "d:\\TABEL\\COPY\\tab.bat");
            //File.Copy("\\\\Fsmtm1\\Install\\!!!Programs!!!\\Pull\\TABEL\\COPY\\tab.bat", "d:\\TABEL\\COPY\\tab.bat");
            FormGeneral.CopyFile("\\\\Fsmtm1\\Install\\!!!Programs!!!\\Pull\\TABEL\\COPY\\UpTABEL.bat", "d:\\TABEL\\COPY\\UpTABEL.bat");
            Process.Start("d:\\TABEL\\COPY\\UpTABEL.bat");
            Close();
        }       
    }
}
