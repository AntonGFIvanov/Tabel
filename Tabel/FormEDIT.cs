using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Tabel
{
    public partial class FormEDIT : Form
    {
        string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\TABEL\BASE\;Extended Properties=dBASE IV;User ID=Admin;Password=";

        public FormEDIT(double[] massiveTemp, string department, string FIO)
        {
            InitializeComponent();
            CultureInfo inf = new CultureInfo(System.Threading.Thread.CurrentThread.CurrentCulture.Name);
            System.Threading.Thread.CurrentThread.CurrentCulture = inf;
            inf.NumberFormat.NumberDecimalSeparator = ".";
            MassiveDouble(massiveTemp);
            textBoxDEP.Text = department;
            textBoxFIO.Text = FIO;
        }

        private void MassiveDouble(double[] massiveTemp)
        {
            int i = 0;
            textBoxTN.Text = Convert.ToString(massiveTemp[i++]);
            textBoxFIO.Text = Convert.ToString(massiveTemp[i++]);
            textBoxFIO.Text = Convert.ToString(massiveTemp[i++]);
            textBoxFIO.Text = Convert.ToString(massiveTemp[i++]);
            textBoxDH.Text = Convert.ToString(massiveTemp[i++]);
            textBoxDNF.Text = Convert.ToString(massiveTemp[i++]);
            textBoxDNP.Text = Convert.ToString(massiveTemp[i++]);
            textBoxDNO.Text = Convert.ToString(massiveTemp[i++]);
            textBoxDNR.Text = Convert.ToString(massiveTemp[i++]);
            textBoxDOU.Text = Convert.ToString(massiveTemp[i++]);
            textBoxDPRO.Text = Convert.ToString(massiveTemp[i++]);
            textBoxADM.Text = Convert.ToString(massiveTemp[i++]);
            textBoxBOLD.Text = Convert.ToString(massiveTemp[i++]);
            textBoxSHR.Text = Convert.ToString(massiveTemp[i++]);
            textBoxPRG.Text = Convert.ToString(massiveTemp[i++]);
            textBoxCAS.Text = Convert.ToString(massiveTemp[i++]);
            textBoxPRAZP.Text = Convert.ToString(massiveTemp[i++]);
            textBoxPRAZG.Text = Convert.ToString(massiveTemp[i++]);
            textBoxWD.Text = Convert.ToString(massiveTemp[i++]);
            textBoxNOC.Text = Convert.ToString(massiveTemp[i++]);
            textBoxKBN.Text = Convert.ToString(massiveTemp[i++]);
            textBoxPRZ.Text = Convert.ToString(massiveTemp[i++]);
            textBoxGOS.Text = Convert.ToString(massiveTemp[i++]);
            textBoxD_SCET.Text = Convert.ToString(massiveTemp[i++]);
            textBoxCAS_PR.Text = Convert.ToString(massiveTemp[i++]);
            textBoxNOC2.Text = Convert.ToString(massiveTemp[i++]);
            textBoxNOWSW.Text = Convert.ToString(massiveTemp[i++]);
            textBoxNOWPR.Text = Convert.ToString(massiveTemp[i++]);
            textBoxNOWWH.Text = Convert.ToString(massiveTemp[i++]);
            textBoxMED.Text = Convert.ToString(massiveTemp[i++]);
            textBoxDONS.Text = Convert.ToString(massiveTemp[i++]);
            textBoxDK.Text = Convert.ToString(massiveTemp[i++]);
            textBoxDRZ.Text = Convert.ToString(massiveTemp[i++]);
            textBoxCAS7.Text = Convert.ToString(massiveTemp[i++]);
            textBoxNP.Text = Convert.ToString(massiveTemp[i++]);
        }        

        private void buttonUPDATE_Click(object sender, EventArgs e)
        {
            string strQueryUpdate = $@"UPDATE TABEL SET DNF={Convert.ToInt32(textBoxDNF.Text)}, 
                                                         WD={Convert.ToInt32(textBoxWD.Text)}, 
                                                        DRZ={Convert.ToInt32(textBoxDRZ.Text)}, 
                                                        DNP={Convert.ToInt32(textBoxDNP.Text)},                                                      
                                                        PRG={Convert.ToInt32(textBoxPRG.Text)},
                                                       BOLD={Convert.ToInt32(textBoxBOLD.Text)},
                                                        DNO={Convert.ToInt32(textBoxDNO.Text)},
                                                        DNR={Convert.ToInt32(textBoxDNR.Text)},
                                                        DOU={Convert.ToInt32(textBoxDOU.Text)},
                                                        ADM={Convert.ToInt32(textBoxADM.Text)},
                                                     D_SCET={Convert.ToInt32(textBoxD_SCET.Text)},
                                                       DONS={Convert.ToInt32(textBoxDONS.Text)},
                                                     CAS_PR={Convert.ToDouble(textBoxCAS_PR.Text.Replace(',', '.'))},
                                                        MED={Convert.ToDouble(textBoxMED.Text.Replace(',', '.'))},
                                                       DPRO={Convert.ToDouble(textBoxDPRO.Text.Replace(',', '.'))},
                                                        PRZ={Convert.ToDouble(textBoxPRZ.Text.Replace(',', '.'))},
                                                         DH={Convert.ToDouble(textBoxDH.Text.Replace(',', '.'))},
                                                        GOS={Convert.ToDouble(textBoxGOS.Text.Replace(',', '.'))},
                                                        SHR={Convert.ToDouble(textBoxSHR.Text.Replace(',', '.'))},
                                                        CAS={Convert.ToDouble(textBoxCAS.Text.Replace(',', '.'))},
                                                      PRAZP={Convert.ToDouble(textBoxPRAZP.Text.Replace(',', '.'))},
                                                      PRAZG={Convert.ToDouble(textBoxPRAZG.Text.Replace(',', '.'))},
                                                        NOC={Convert.ToDouble(textBoxNOC.Text.Replace(',', '.'))},
                                                       NOC2={Convert.ToDouble(textBoxNOC2.Text.Replace(',', '.'))},
                                                      NOWPR={Convert.ToDouble(textBoxNOWPR.Text.Replace(',', '.'))},
                                                        KBN={Convert.ToDouble(textBoxKBN.Text.Replace(',', '.'))},
                                                         DK={Convert.ToDouble(textBoxDK.Text.Replace(',', '.'))},
                                                      NOWSW={Convert.ToDouble(textBoxNOWSW.Text.Replace(',', '.'))},
                                                      NOWWH={Convert.ToDouble(textBoxNOWWH.Text.Replace(',', '.'))}, 
                                                       CAS7={Convert.ToInt32(textBoxCAS7.Text)},
                                                         NP={Convert.ToInt32(textBoxNP.Text)}    
        
                                                   WHERE TN={Convert.ToInt32(textBoxTN.Text)}";
            try
            {
                using (OleDbConnection cn = new OleDbConnection(con))
                {                    
                    cn.Open();
                    OleDbCommand command = new OleDbCommand(strQueryUpdate, cn);
                    command.ExecuteNonQuery();
                    MessageBox.Show(" \nOK!\n!" );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(" ERROR: \nОшибка обновления!\n" + ex.Message);
            }
        }

        private void buttonCLOSE_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FormEDIT_Load(object sender, EventArgs e)
        {

        }
    }
}
