using Microsoft.Office.Interop.Excel;
using System;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Tabel
{
    public partial class FormCorrection : Form
    {
        static string connectionStringBase = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\TABEL\BASE\;Extended Properties=dBASE IV;User ID=Admin;Password=";
        string department;

        //Конструктор

        public FormCorrection()
        {
            InitializeComponent();
        }


        // Функции

        #region Обновление dataGridView1

        public static void UpDataGrid(ComboBox cmbBox, DataGridView dgview)
        {
            string strQuery = $@"Select distinct tabel.tn as ТН, sprrab.fam as Фамилия, sprrab.imq as Имя, sprrab.otc as Отчество,
                                           tabel.dh as 'дни хоз',tabel.dnf as 'дни факт',tabel.dnp as 'дни простоя',tabel.dno as 'дни отп',
                                           tabel.dnr as 'родовые',tabel.dou as 'уч отп',tabel.dpro as 'прочие',tabel.adm as 'адм отп',
                                           tabel.bold as 'дни больн',tabel.shr as 'по среднему ПД',tabel.prg as 'прогул',tabel.cas as 'часы факт',
                                           tabel.prazp as 'праз приказ',tabel.prazg as 'праз граф',tabel.wd as 'вых',tabel.noc as 'ночн 1',
                                           tabel.kbn as 'сверхнормы',tabel.prz as 'презид',tabel.gos as 'гос обяз',tabel.d_scet as 'дни св счет',
                                           tabel.cas_pr as 'часы простоя',tabel.noc2 as 'ночн 2',tabel.nowsw as 'сверхур 1 опл',tabel.nowpr as 'праз 1 опл вых',
                                           tabel.nowwh as 'доп день отд',tabel.med as 'мед спр',tabel.kolh as 'донор б/о',tabel.DK as 'ком служ',
                                           tabel.DRZ as 'в др цехах',tabel.CAS7 as 'доп час 7+1', tabel.NP as 'по среднему НПД'
                                           from tabel,sprrab where tabel.kc={Convert.ToInt32(cmbBox.Text)} and sprrab.kc={Convert.ToInt32(cmbBox.Text)} and tabel.tn=sprrab.tn and (sprrab.puvl=0 or sprrab.puvl=1 or sprrab.puvl=5 or sprrab.puvl=9) order by sprrab.fam , sprrab.imq , sprrab.otc";
            /*puvl - признак увольнения/перевода
            // 0 - текущее место работы
            // 1 - перевод
            // 5 - декрет
            // 9 - уволен
            */
            dgview.DataSource = FormGeneral.DTselect(strQuery, connectionStringBase);

            for (int k = 0; k < dgview.ColumnCount; k++)
            {
                dgview.Columns[k].Width = 40;
            }
            dgview.Columns["ТН"].Width = 45;
            dgview.Columns["'часы факт'"].Width = 50;
            dgview.Columns["Фамилия"].Width = 120;
            dgview.Columns["Имя"].Width = 100;
            dgview.Columns["Отчество"].Width = 110;

            for (int i = 0; i < dgview.RowCount; i++)
            {
                for (int j = 4; j < dgview.ColumnCount; j++)
                {
                    if (!(dgview.Rows[i].Cells[j].Value.ToString() == "0") && !(dgview.Rows[i].Cells[j].Value.ToString() == ""))
                        dgview.Rows[i].Cells[j].Style.BackColor = Color.LightCoral;
                }
            }
        }

        #endregion

        #region Запросы к базам данных

        public static string stringSELECT(string strQUERY)
        {
            string str = "";
            using (OleDbConnection CN1 = new OleDbConnection(connectionStringBase))
            {
                CN1.Open();
                OleDbCommand command = new OleDbCommand(strQUERY, CN1);
                if (command.ExecuteScalar() is string)
                {
                    str = (string)(command.ExecuteScalar());
                    return str;
                }
                return str;
            }
        }

        #endregion

        #region Получение даты

        public static DateTime DateTimeSELECT(string strQUERY)
        {
            using (OleDbConnection CN1 = new OleDbConnection(connectionStringBase))
            {
                DateTime temp = new DateTime(2000, 01, 01);
                CN1.Open();
                OleDbCommand command = new OleDbCommand(strQUERY, CN1);
                if (command.ExecuteScalar() is DateTime)
                {
                    temp = Convert.ToDateTime(command.ExecuteScalar());
                    return temp;
                }
                return temp;
            }
        }

        #endregion

        #region Формирование документа Excel

        public void DataGridToExcel(DataGridView GridSourse)
        {
            try
            {
                string fileExcel = $"d:\\TABEL\\Табель{comboBoxDEP.Text}{DateTime.Now.Month}.xls";

                CopyFile("d:\\TABEL\\COPY\\MyFile.xls", fileExcel);
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Workbooks.Open(fileExcel);
                Workbook ExcelWorkBook;
                Worksheet ExcelWorkSheet;
                //Книга.           
                ExcelWorkBook = ExcelApp.Workbooks.Open(fileExcel);
                //Таблица.
                ExcelWorkSheet = (Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        ExcelApp.Cells[i + 3, 1] = i + 1;
                        ExcelApp.Cells[i + 3, 1].Borders.Color = BorderStyle.FixedSingle;

                        if (!(dataGridView1.Rows[i].Cells[j].Value.ToString() == "0"))
                        {
                            ExcelApp.Cells[i + 3, j + 2] = dataGridView1.Rows[i].Cells[j].Value;
                            ExcelApp.Cells[i + 3, j + 2].Borders.Color = BorderStyle.FixedSingle;
                            if (j == 2 || j == 3)
                            {
                                ExcelApp.Cells[i + 3, j + 2] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 1);
                                ExcelApp.Cells[i + 3, j + 2].Borders.Color = BorderStyle.FixedSingle;
                            }
                        }
                        else
                        {
                            ExcelApp.Cells[i + 3, j + 2] = "";
                            ExcelApp.Cells[i + 3, j + 2].Borders.Color = BorderStyle.FixedSingle;
                        }
                    }
                }
                //Вызываем эксель.
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
            catch
            {
                MessageBox.Show("Ошибка выгрузки данных! \nDataGridToExcel!");
            }
        }

        #endregion

        #region Копирование файлов

        void CopyFile(string sourcefn, string destinfn)
        {
            //sourcefn - файл,который копируем,с путем
            //destinfn - имя с путем куда копируем
            FileInfo fn = new FileInfo(sourcefn);
            fn.CopyTo(destinfn, true);
        }

        #endregion

        #region Печать контрольного ярлыка
        void PrintControlLabel(DataGridView datagrid)
        {
            decimal dh = 0, dk = 0, shr = 0, cas = 0, prazp = 0, prazg = 0, noc = 0, kbn = 0, prz = 0, gos = 0, cas_pr = 0, noc2 = 0, nowsw = 0, nowpr = 0, nowwh = 0, CAS7 = 0, NP = 0;
            int dnf = 0, dnp = 0, dno = 0, dnr = 0, dou = 0, dpro = 0, adm = 0, bold = 0, prg = 0, wd = 0, d_scet = 0, med = 0, kolh = 0;

            for (int i = 0; i <= datagrid.RowCount - 1; i++)
            {
                for (int j = 3; j <= datagrid.ColumnCount - 4; j++)
                {
                    if (datagrid.Rows[i].Cells[j].Value == null || datagrid.Rows[i].Cells[j].Value.ToString() == "")
                    {
                        Convert.ToInt32(datagrid.Rows[i].Cells[j].Value = 0);
                        //(int)(datagrid.Rows[i].Cells[j].Value ?? 0);
                    }
                    if (datagrid.Rows[i].Cells[j].Value == DBNull.Value)
                    {
                        Convert.ToInt32(datagrid.Rows[i].Cells[j].Value.ToString());
                    }
                }
            }

            for (int i = 0; i <= datagrid.RowCount - 1; i++)
            {
                // переменные в часах
                dh += Convert.ToDecimal(datagrid.Rows[i].Cells["'дни хоз'"].Value);
                dk += Convert.ToDecimal(datagrid.Rows[i].Cells["'ком служ'"].Value);
                shr += Convert.ToDecimal(datagrid.Rows[i].Cells["'по среднему ПД'"].Value);
                cas += Convert.ToDecimal(datagrid.Rows[i].Cells["'часы факт'"].Value);
                prazp += Convert.ToDecimal(datagrid.Rows[i].Cells["'праз приказ'"].Value);
                prazg += Convert.ToDecimal(datagrid.Rows[i].Cells["'праз граф'"].Value);
                noc += Convert.ToDecimal(datagrid.Rows[i].Cells["'ночн 1'"].Value);
                kbn += Convert.ToDecimal(datagrid.Rows[i].Cells["'сверхнормы'"].Value);
                prz += Convert.ToDecimal(datagrid.Rows[i].Cells["'презид'"].Value);
                gos += Convert.ToDecimal(datagrid.Rows[i].Cells["'гос обяз'"].Value);
                cas_pr += Convert.ToDecimal(datagrid.Rows[i].Cells["'часы простоя'"].Value);
                noc2 += Convert.ToDecimal(datagrid.Rows[i].Cells["'ночн 2'"].Value);
                nowsw += Convert.ToDecimal(datagrid.Rows[i].Cells["'сверхур 1 опл'"].Value);
                nowpr += Convert.ToDecimal(datagrid.Rows[i].Cells["'праз 1 опл вых'"].Value);
                nowwh += Convert.ToDecimal(datagrid.Rows[i].Cells["'доп день отд'"].Value);
                CAS7 += Convert.ToDecimal(datagrid.Rows[i].Cells["'доп час 7+1'"].Value);
                NP += Convert.ToDecimal(datagrid.Rows[i].Cells["'по среднему НПД'"].Value);

                //переменные в днях
                dnf += Convert.ToInt32(datagrid.Rows[i].Cells["'дни факт'"].Value);
                dnp += Convert.ToInt32(datagrid.Rows[i].Cells["'дни простоя'"].Value);
                dno += Convert.ToInt32(datagrid.Rows[i].Cells["'дни отп'"].Value);
                dnr += Convert.ToInt32(datagrid.Rows[i].Cells["'родовые'"].Value);
                dou += Convert.ToInt32(datagrid.Rows[i].Cells["'уч отп'"].Value);
                dpro += Convert.ToInt32(datagrid.Rows[i].Cells["'прочие'"].Value);
                adm += Convert.ToInt32(datagrid.Rows[i].Cells["'адм отп'"].Value);
                bold += Convert.ToInt32(datagrid.Rows[i].Cells["'дни больн'"].Value);
                prg += Convert.ToInt32(datagrid.Rows[i].Cells["'прогул'"].Value);
                wd += Convert.ToInt32(datagrid.Rows[i].Cells["'вых'"].Value);
                d_scet += Convert.ToInt32(datagrid.Rows[i].Cells["'дни св счет'"].Value);
                med += Convert.ToInt32(datagrid.Rows[i].Cells["'мед спр'"].Value);
                kolh += Convert.ToInt32(datagrid.Rows[i].Cells["'донор б/о'"].Value);
            }
            string fileExcel = "d:\\TABEL\\COPY\\ControlLabel.xls";

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Workbooks.Open(fileExcel);
            Workbook ExcelWorkBook;
            Worksheet ExcelWorkSheet;
            //Книга.           
            ExcelWorkBook = ExcelApp.Workbooks.Open(fileExcel);
            //Таблица.
            ExcelWorkSheet = (Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            ExcelApp.Cells[1, 2] = comboBoxDEP.Text;

            ExcelApp.Cells[6, 3] = dnf;
            ExcelApp.Cells[7, 3] = cas;
            ExcelApp.Cells[8, 3] = wd;
            ExcelApp.Cells[9, 3] = prazp;
            ExcelApp.Cells[10, 3] = nowpr;
            ExcelApp.Cells[11, 3] = nowsw;
            ExcelApp.Cells[12, 3] = prazg;
            ExcelApp.Cells[13, 3] = noc;
            ExcelApp.Cells[14, 3] = noc2;
            ExcelApp.Cells[15, 3] = bold;
            ExcelApp.Cells[16, 3] = dno;
            ExcelApp.Cells[17, 3] = dk;
            ExcelApp.Cells[18, 3] = dh;
            ExcelApp.Cells[19, 3] = prz;
            ExcelApp.Cells[20, 3] = dnp;
            ExcelApp.Cells[21, 3] = cas_pr;
            ExcelApp.Cells[22, 3] = d_scet;
            ExcelApp.Cells[23, 3] = kolh;
            ExcelApp.Cells[24, 3] = prg;
            ExcelApp.Cells[25, 3] = NP;
            ExcelApp.Cells[26, 3] = shr;
            ExcelApp.Cells[27, 3] = 000;
            ExcelApp.Cells[28, 3] = med;
            ExcelApp.Cells[29, 3] = CAS7;
            ExcelApp.Cells[30, 3] = nowwh;
            ExcelApp.Cells[31, 3] = dnr;
            ExcelApp.Cells[32, 3] = adm;
            ExcelApp.Cells[33, 3] = dpro;
            ExcelApp.Cells[34, 3] = gos;
            ExcelApp.Cells[35, 3] = dou;

            //Вызываем эксель.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
            //ExcelWorkSheet.PrintOutEx(); //печать ярлыка в фоновом режиме
        }
        #endregion

        // События

        #region Загрузка формы

        private void FormCorrection_Load(object sender, EventArgs e)
        {
            try
            {
                //Проверка табеля на актуальность данных
                int monthNOW = DateTime.Now.Month;
                if (DateTime.Now.Day < 5)
                {
                    //monthNOW = monthNOW - 1;
                    monthNOW--;
                }
                string strQueryDate = "select distinct data from tabel";
                int monthFROMbase = DateTimeSELECT(strQueryDate).Month;
                if (monthNOW != monthFROMbase)
                {
                    MessageBox.Show("Текущая версия табеля устарела!\nВыполните загрузку табеля для корректировки!");
                    Close();
                }
                else
                {
                    string strQueryKC = "select DISTINCT kc  from tabel";
                    System.Data.DataTable dataTable = FormGeneral.DTselect(strQueryKC, connectionStringBase);
                    comboBoxDEP.DataSource = dataTable.DefaultView;
                    comboBoxDEP.DisplayMember = "kc";
                    comboBoxDEP.ValueMember = "KC";
                    comboBoxDEP.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Выполните загрузку табеля для корректировки! \n" + ex.Message);
                Close();
            }
        }

        #endregion

        #region Выбор значения из комбобокса
        private void comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                UpDataGrid(comboBoxDEP, dataGridView1);

                //string strQueryNAME = $"select knz  from sp where  kc={Convert.ToInt32(comboBoxDEP.Text)}";
                //label2.Text = stringSELECT(strQueryNAME);
                //department = $"{comboBoxDEP.Text}  |  {label2.Text}";
                department = $"{comboBoxDEP.Text}";
            }
            catch
            {
                return;
            }
        }

        #endregion

        #region Нажатие кнопки 1

        private void button1_Click(object sender, EventArgs e)
        {
            double[] massiveTemp = new double[dataGridView1.ColumnCount];
            int indexRow = dataGridView1.CurrentRow.Index;

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                if (i == 1 || i == 2 || i == 3)
                    massiveTemp[i] = 0;
                else if (!(dataGridView1.Rows[indexRow].Cells[i].Value.ToString() == "0") && !(dataGridView1.Rows[indexRow].Cells[i].Value.ToString() == ""))
                    massiveTemp[i] = Convert.ToDouble(dataGridView1.Rows[indexRow].Cells[i].Value);
                else
                    massiveTemp[i] = 0;
            }
            string FIO = $"{dataGridView1.Rows[indexRow].Cells[1].Value}  {dataGridView1.Rows[indexRow].Cells[2].Value}  {dataGridView1.Rows[indexRow].Cells[3].Value}";
            FormEDIT formTemp = new FormEDIT(massiveTemp, department, FIO);
            formTemp.ShowDialog();
        }

        #endregion

        #region Активация формы

        private void FormCorrection_Activated(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource == null)
                return;
            else
                UpDataGrid(comboBoxDEP, dataGridView1);
        }

        #endregion

        #region Закрытие формы

        private void FormCorrection_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (dataGridView1.DataSource == null)
                return;
            else
            {
                /*
                DialogResult dialogResult = MessageBox.Show("Если были произведены измения, необходима выгрузка данных.\nВыгрузить данные для бухгалтерии?", "Выгрузка данных", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    string strQueryDate = "select distinct data from tabel";
                    int monthFROMbase = DateTimeSELECT(strQueryDate).Month;
                    string sourcefn, sourcefnPer;  //Имя файла,который копируем
                    int kcname = FormGeneral.INTSelectDBF("select kc from tabel order by kc", connectionStringBase);
                    sourcefn = "T" + kcname;
                    sourcefnPer = "P" + kcname + monthFROMbase.ToString();
                    FormGeneral.CopyFile("d:\\TABEL\\BASE\\TABEL.DBF", "\\\\Mztm\\Trmash_Data\\Maz\\NEWTABEL\\" + sourcefn + ".DBF");
                    FormGeneral.CopyFile("d:\\TABEL\\BASE\\period.DBF", "\\\\Mztm\\Trmash_Data\\Maz\\NEWTABEL\\BaseP\\" + sourcefnPer + ".DBF");
                    FormGeneral.InsertCommonFile(sourcefn, kcname);
                    //PrintControlLabel(dataGridView1);
                }
                else if (dialogResult == DialogResult.No)
                {
                    return;
                }
                */
                return;
            }
        }

        #endregion

        #region Нажатие кнопки Выход

        private void buttonExitSavePrint_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion

        #region Нажатие кнопки 2

        private void button2_Click(object sender, EventArgs e)
        {
            //DataGridToExcel(dataGridView1);
        }

        #endregion

        #region Нажатие кнопки Печать ярлыка

        private void buttonPrint_Click(object sender, EventArgs e)
        {
            PrintControlLabel(dataGridView1);
        }

        #endregion

    }
}
