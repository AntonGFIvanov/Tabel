using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Globalization;
using System.IO;

namespace Tabel
{
    public partial class Form1 : Form
    {
        string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\TABEL\;Extended Properties=dBASE IV;User ID=Admin;Password=";
        string bd;
        int monthGlobal;

        public Form1()
        {
            InitializeComponent();

            CultureInfo inf = new CultureInfo(System.Threading.Thread.CurrentThread.CurrentCulture.Name);
            System.Threading.Thread.CurrentThread.CurrentCulture = inf;
            inf.NumberFormat.NumberDecimalSeparator = ".";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button2.Text = "В ОБЩИЙ ФАЙЛ (2)";
            button3.Text = "ПЕРЕСЧИТАТЬ (3)";
            button4.Text = "ВЫГРУЗИТЬ В СЕТЬ (4)";
            // Displays an OpenFileDialog so the user can select a Cursor.  
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "d:\\TABEL\\tab\\";
            openFileDialog1.Filter = "Файл БД| *.DBF";
            openFileDialog1.Title = "Выберите табель цеха";

            //string filePath = @"d:\TABEL\tab\";
            //Path.GetDirectoryName(filePath);

            // Show the Dialog.  
            // If the user clicked OK in the dialog and  
            // a .dbf file was selected, open it. 
            try
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        //Копирование из сети на D диск пользователя
                        //CopyFile( openFileDialog1.FileName , "d:\\TABEL\\"+ openFileDialog1.FileName);

                        string strSQL = "SELECT * FROM " + openFileDialog1.FileName;
                        dataGridView1.DataSource = DTselect(strSQL, con);
                        label2.Text = openFileDialog1.FileName;
                        bd = openFileDialog1.FileName;
                    }
                button1.Text = "ВЫПОЛНЕНО";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки файла с диска. ERROR: " + ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Перевод в структуру табеля МТМ
            try
            {
                int rowcount = dataGridView1.RowCount;
                string delt = "delete * from TABEL ";
                string INSERT = "INSERT INTO TABEL (TN, kc_maz, ku_maz, plan_f)  SELECT DISTINCT tn, ceh ,uch, plan_priv FROM " + bd + " where vo=0 and plan_priv>0";

                SelectUpdate(delt, con);
                SelectUpdate(INSERT, con);

                button2.Text = "ВЫПОЛНЕНО";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки! " + ex);
            }
        }

        DateTime date1 = new DateTime(0, 0);
        private void timer1_Tick(object sender, EventArgs e)
        {
            date1 = date1.AddSeconds(1);
            //label3.Text = date1.ToString("mm:ss");
            timer1.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {               
                int rowcount = dataGridView1.RowCount;
                int tn = 0;
                int bold = 0, dnf = 0, dno = 0, dnp = 0, dou = 0, prg = 0, wd = 0, adm = 0, dnr = 0, med = 0;
                decimal cas = 0, dk = 0, kbn = 0, prz = 0, dons = 0, cas_pr = 0, nowpr = 0, nowsw = 0, nowwh = 0, prazp = 0, d_scet = 0, gos = 0, cas_scet = 0;
                decimal noc = 0, noc2 = 0;
                
                //Переменные для работы с периодами
                #region Переменные периодов 

                DateTime dateTimeStart; //Дата начала периода
                DateTime dateTimeEnd;  //Дата конца периода
                string cursor1;  //Указатель на тип 
                string cursor2;  //Указатель на тип следущего дня
                int superDay;  //Счетчик числа
                bool marker;  //Маркер сброса начального указателя

                #endregion

                for (int i = 0; i < rowcount-1; i++)
                {
                    int MounthTemp = Convert.ToInt32(dataGridView1.Rows[i].Cells["mes"].Value);
                    monthGlobal = MounthTemp; //Глобальная переменная для месяца
                    int YearTemp= Convert.ToInt32(dataGridView1.Rows[i].Cells["god"].Value);
                    int tn1 = Convert.ToInt32(dataGridView1.Rows[i].Cells["tn"].Value);
                    int kodCex = Convert.ToInt32(dataGridView1.Rows[i].Cells["ceh"].Value);

                    if (tn1 != tn)
                    {
                        bold = 0; dnf = 0; dno = 0; dnp = 0; dou = 0; prg = 0; wd = 0; adm = 0; dnr = 0; med = 0;
                        cas = 0; dk = 0; kbn = 0; prz = 0; dons = 0; cas_pr = 0; nowpr = 0; nowsw = 0; nowwh = 0; prazp = 0; d_scet = 0; gos = 0; cas_scet = 0;
                        noc = 0; noc2 = 0;
                        tn = tn1;
                    }

                    decimal testing = 0; //Переменная используется для перевода минут в дясятичный вид
                    cursor1 ="";
                    marker = false;
                    superDay = 0;
                    dateTimeStart = new DateTime(YearTemp, MounthTemp, 01);

                    if (Convert.ToInt32(dataGridView1.Rows[i].Cells["vo"].Value) == 0)
                    {
                        for (int j = 18; j < 204; j++)
                        {
                            if (dataGridView1.Rows[i].Cells[j].Value is string)
                            {
                                string vidd = dataGridView1.Rows[i].Cells[j].Value.ToString();

                                // Подсчет периодов
                                #region Подсчет периодов
                                if (j % 6 == 0)
                                {
                                    #region Обработка месяцев с 31 числом

                                    if (monthGlobal == 1 || monthGlobal == 3 || monthGlobal == 5 || monthGlobal == 7 || monthGlobal == 8 || monthGlobal == 10 || monthGlobal == 12)
                                    {
                                        if (marker == false)
                                        {
                                            cursor1 = vidd;
                                            marker = true;
                                        }
                                        dateTimeEnd = DateNew(j, MounthTemp, YearTemp, ref superDay);
                                        cursor2 = vidd;
                                        dateTimeStart = WorkBase(ref cursor1, cursor2, dateTimeStart, dateTimeEnd, tn1, kodCex);
                                    }
                                    #endregion

                                    #region Обработка месяцев с 30 числом

                                    if (monthGlobal == 4 || monthGlobal == 6 || monthGlobal == 9 || monthGlobal == 11)
                                    {
                                        if (marker == false)
                                        {
                                            cursor1 = vidd;
                                            marker = true;
                                        }
                                        dateTimeEnd = DateNew(j, MounthTemp, YearTemp, ref superDay);
                                        cursor2 = vidd;
                                        dateTimeStart = WorkBaseTwo(ref cursor1, cursor2, dateTimeStart, dateTimeEnd, tn1, kodCex);
                                    }
                                    #endregion

                                    #region Обработка февраля

                                    if (monthGlobal == 2)
                                    {
                                        if (marker == false)
                                        {
                                            cursor1 = vidd;
                                            marker = true;
                                        }
                                        dateTimeEnd = DateNew(j, MounthTemp, YearTemp, ref superDay);
                                        cursor2 = vidd;
                                        dateTimeStart = WorkBaseFEB(ref cursor1, cursor2, dateTimeStart, dateTimeEnd, tn1, kodCex);
                                    }
                                    #endregion
                                }
                                #endregion

                                switch (vidd)
                                {
                                    #region Командировка                           
                                    case "К*":
                                    case "К":
                                        testing = 0;
                                        if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                            testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                        Ref(ref testing);
                                        dk += testing;
                                        break;
                                    #endregion

                                    #region Часы фактические  
                                    case "1*":
                                    case "2":
                                    case "2*":
                                    case "3":
                                    case "3*":
                                    case "1":
                                        dnf++;
                                        testing = 0;
                                        if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                            testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                        Ref(ref testing);
                                        cas += testing;
                                        break;
                                    #endregion

                                    #region Часы в выходной без доплат
                                    /*case "1В":
                                    case "2В":
                                    case "3В":
                                        if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                            nowpr += Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                        break;
                                        */
                                    #endregion

                                    #region Часы праздничные/выходные за 1 оплату
                                    case "1П":
                                    case "2П":
                                    case "3П":
                                        testing = 0;
                                        if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                            testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                        Ref(ref testing);
                                        nowpr += testing;
                                        break;

                                    #endregion

                                    #region Часы 2-ая оплата (13 вид)
                                    /*
                                     case "1\"":
                                     case "2\"":
                                     case "3\"":
                                         if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                             prazp += Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                         break;*/
                                    #endregion

                                    #region Дни труд-го отпуска
                                    case "О*":
                                    case "О":
                                    case "ПО*":
                                    case "ПО":
                                        dno++;
                                        break;
                                    #endregion

                                    #region Дни больничного                            
                                    case "Б*":
                                    case "5Б":
                                    case "5Б*":
                                    case "Б":
                                        bold++;
                                        break;
                                    #endregion

                                    #region Дни заводского простоя
                                    case "5*":
                                    case "5":
                                        dnp++;
                                        break;
                                    #endregion

                                    #region Часы цехового простоя
                                    case "П*":
                                    case "П":
                                        testing = 0;
                                        if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                            testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                        Ref(ref testing);
                                        cas_pr += testing;
                                        break;
                                    #endregion

                                    #region Часы свой счет (административный)
                                    case "А*":
                                        // дни за свой счет
                                        if (Convert.ToInt32(dataGridView1.Rows[i].Cells[j + 2].Value) == 7)
                                            d_scet++;
                                        else
                                        //часы за свой счет
                                        {
                                            testing = 0;
                                            if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                                testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                            Ref(ref testing);
                                            cas_scet += testing;
                                        }
                                        break;
                                    case "А":
                                        if (Convert.ToInt32(dataGridView1.Rows[i].Cells[j + 2].Value) == 8)
                                            d_scet++;
                                        else
                                        {
                                            testing = 0;
                                            if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                                testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                            Ref(ref testing);
                                            cas_scet += testing;
                                        }
                                        break;
                                    #endregion

                                    #region Часы гос. обяз.
                                    case "Г*":
                                    case "Г":
                                        testing = 0;
                                        if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                            testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                        Ref(ref testing);
                                        gos += testing;
                                        break;
                                    #endregion

                                    #region Дни уч. отпуска без сохранения зп
                                    case "УБЗ*":
                                    case "УБЗ":
                                        dou++;
                                        break;
                                    #endregion

                                    #region Дни род. больничного
                                    case "Р*":
                                    case "Р":
                                        dnr++;
                                        break;
                                    #endregion

                                    #region Дни прогулов
                                    case "ПР*":
                                    case "ПР":
                                        prg++;
                                        break;
                                    #endregion

                                    #region Часы по указу президента(день матери)
                                    case "ДМ*":
                                    case "ДМ":
                                        testing = 0;
                                        if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                            testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                        Ref(ref testing);
                                        prz += testing;
                                        break;
                                    #endregion

                                    #region Часы донорсие с оплатой(донорский день)
                                    case "Д*":
                                    case "Д":
                                        testing = 0;
                                        if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                            testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                        Ref(ref testing);
                                        dons += testing;
                                        break;
                                    #endregion

                                    #region Выходной день
                                    case "ПД":
                                    case "В":
                                        wd++;
                                        break;
                                    #endregion

                                    #region Часы сверхурочные(8 вид)
                                    case "СВ":
                                        testing = 0;
                                        if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                            testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                        Ref(ref testing);
                                        kbn += testing;
                                        break;
                                    #endregion

                                    #region Отпуск по уходу за ребенком до 3 лет
                                    case "ОЖ*":
                                    case "ОЖ":
                                        adm++;
                                        break;
                                    #endregion

                                    #region Часы выходных за ранее отработанное время(неопл день отдыха)
                                    case "О1^*":
                                    case "О2^":
                                    case "О2^*":
                                    case "О3^":
                                    case "О3^*":
                                    case "О1^":
                                        testing = 0;
                                        if (dataGridView1.Rows[i].Cells[j + 2].Value is double)
                                            testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j + 2].Value);
                                        Ref(ref testing);
                                        nowwh += testing;
                                        break;
                                    #endregion

                                    #region Неявка по невыясненным причинам(прогул)
                                    case "НН*":
                                    case "НН":
                                        prg++;
                                        break;
                                    #endregion

                                    #region Дни мед. справки
                                    case "Х":
                                    case "Х*":
                                        med++;
                                        break;
                                    #endregion

                                    default:
                                        break;
                                }
                            }
                            else
                                continue;
                        }
                        string update = "update tabel set d_scet=" + d_scet + ",nowwh=" + nowwh + ",nowpr=" + nowpr + ",cas_pr=" + cas_pr + ",dons=" + dons + ",prz=" + prz + ",kbn=" + kbn + ",adm=" + adm + ",wd=" + wd + ",prg=" + prg + ",dk=" + dk + ", dnf=" + dnf + ", dno=" + dno + ", cas=" + cas + ", bold= " + bold + ", dnp=" + dnp + ", dou=" + dou + ", dnr=" + dnr + ", med=" + med + " where tn=" + Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value);

                        SelectUpdate(update, con);
                    }

                    #region Расчет 7 вида оплаты (ночные)
                    else if (Convert.ToInt32(dataGridView1.Rows[i].Cells["vo"].Value) == 7)
                    {
                        noc = Convert.ToDecimal(dataGridView1.Rows[i].Cells["chas_mes"].Value);
                        /*
                        for (int j = 18; j < 204; j++)
                        {
                            testing = 0;
                            if (dataGridView1.Rows[i].Cells[j].Value is double)
                                testing = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j].Value);
                            Ref(ref testing);
                            noc += testing;
                        }
                        */
                        SelectUpdate("update tabel set noc=" + noc + " where tn=" + Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value), con);
                    }
                    #endregion

                    #region Расчет 28 вида оплаты (ночные)
                    else if (Convert.ToInt32(dataGridView1.Rows[i].Cells["vo"].Value) == 28)
                    {
                        noc2 = Convert.ToDecimal(dataGridView1.Rows[i].Cells["chas_mes"].Value);
                        
                        SelectUpdate("update tabel set  noc2=" + noc2 + " where tn=" + Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value), con);
                    }
                    #endregion

                    #region Расчет 13 вида оплаты (двойная)
                    else if (Convert.ToInt32(dataGridView1.Rows[i].Cells["vo"].Value) == 13)
                    {
                        prazp = Convert.ToDecimal(dataGridView1.Rows[i].Cells["chas_mes"].Value);
                       
                        SelectUpdate("update tabel set  prazp=" + prazp + " where tn=" + Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value), con);
                    }
                    #endregion

                    #region Расчет 12 вида оплаты (дежурство в ПД)
                    else if (Convert.ToInt32(dataGridView1.Rows[i].Cells["vo"].Value) == 12)
                    {
                        decimal prazg = 0;
                        prazg = Convert.ToDecimal(dataGridView1.Rows[i].Cells["chas_mes"].Value);

                        SelectUpdate("update tabel set  prazg=" + prazg + " where tn=" + Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value), con);
                    }
                    #endregion

                    #region Расчет 8 вида оплаты (сверхурочные)
                    else if (Convert.ToInt32(dataGridView1.Rows[i].Cells["vo"].Value) == 8)
                    {
                        nowsw = Convert.ToDecimal(dataGridView1.Rows[i].Cells["chas_mes"].Value);
                        /*for (int j = 18; j < 204; j++)
                        {
                            if (dataGridView1.Rows[i].Cells[j].Value is double)
                                nowsw += Convert.ToDecimal(dataGridView1.Rows[i].Cells[j].Value);
                        */
                        SelectUpdate("update tabel set nowsw=" + nowsw + " where tn=" + Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value), con);
                    }
                    #endregion
                }
                button3.Text = "ВЫПОЛНЕНО";
                Zam(dataGridView1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки! " + ex.Message);
            }
        }
        #region Замена цехов и участков МАЗ
        public static void Zam(DataGridView temp)
        {
            string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\TABEL\;Extended Properties=dBASE IV;User ID=Admin;Password=";
            OleDbConnection cn;
            cn = new OleDbConnection(con);

            //замена кода цеха МАЗ на код цеха МТМ
            string uch = null;
            int rowcount = temp.RowCount;
            int kc_maz, kc2 = 0;
                         
            for (int i = 0; i < rowcount; i++)
            {
                kc_maz = INTSelectDBF("SELECT kc_maz FROM tabel where tn=" + Convert.ToInt32(temp.Rows[i].Cells["tn"].Value), con);
                int tn_work = Convert.ToInt32(temp.Rows[i].Cells["tn"].Value);

                #region Добавление кода цеха МТМ  
                switch (kc_maz)
                {
                    case 210:
                        if (temp.Rows[i].Cells["uch"].Value.ToString() != "")
                        {
                            uch = temp.Rows[i].Cells["uch"].Value.ToString();
                            kc2 = Zam_uch(210, uch, cn);
                        }
                        break;
                    case 280:
                        kc2 = 201;
                        break;
                    case 282:
                        kc2 = 206;
                        break;
                    case 284:
                        kc2 = 200;
                        break;                    
                    case 288:
                        kc2 = 203;
                        break;                   
                    case 292:
                        kc2 = 205;
                        break;
                    case 294:
                        kc2 = 207;
                        break;
                    case 296:
                        kc2 = 303;
                        break;
                    case 298:
                        kc2 = 302;
                        break;                    
                    case 303:
                        kc2 = 301;
                        break;
                    case 305:
                        kc2 = 304;
                        break;
                    case 310:
                        if (temp.Rows[i].Cells["uch"].Value.ToString() != "")
                        {
                            uch = temp.Rows[i].Cells["uch"].Value.ToString();
                            kc2 = Zam_uch(310, uch, cn);
                        }
                        break;                        
                }
                #endregion
                string monthDate = MonthDate();

                if (kc_maz == 210 || kc_maz == 310)
                {
                    SelectUpdate("update TABEL set kc=" + kc2 + ", data='" + Convert.ToDateTime(monthDate) + "' where kc_maz=" + kc_maz + " and ku_maz='" + uch + "'", con);
                    SelectUpdate("update period set kc=" + kc2 + " where tn=" + tn_work, con);
                }

                else
                {
                    SelectUpdate("update TABEL set kc=" + kc2 + ", data='" + Convert.ToDateTime(monthDate) + "' where kc_maz=" + kc_maz, con);
                    SelectUpdate("update period set kc=" + kc2 + " where tn=" + tn_work, con);
                }
            }
        }
        public static int Zam_uch(int kc, string uch, OleDbConnection cn)
        {           
            string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\TABEL\;Extended Properties=dBASE IV;User ID=Admin;Password=";
            string kcu;
            int kcmtm;

            #region Перевод участков 210 и 310 цеха МАЗ в отделы МТМ
            String substring = uch.Substring(0, 2);
            kcu = Convert.ToString(kc) + substring;
            kcmtm = INTSelectDBF("SELECT kc FROM sp where kskmaz=" + Convert.ToInt32(kcu), con);
            return kcmtm;
            #endregion
        }
        #endregion

        public static void Ref(ref decimal test)
        {
            decimal test2, test3;
            test2 = test - (int)test;
            test3 = test2 * 100 / 60;
            test = (int)test + test3;
        }

        public static string MonthDate()
        {
            DateTime temp;
            int data1 = DateTime.Now.Day;
            int data2 = DateTime.Now.Month;
            int data3 = DateTime.Now.Year;

            #region Определение последнего числа отчетного месяца
            if (data1 < 10)
            {
                if (data2 > 1)
                    data2 -= 1;
                else
                {
                    data2 = 12;
                    data3 -= 1;
                }
            }
            switch (data2)
            {
                case 1:
                    {
                        temp = new DateTime(data3, data2, 31);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 2:
                    {
                        temp = new DateTime(data3, data2, 28);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 3:
                    {
                        temp = new DateTime(data3, data2, 31);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 4:
                    {
                        temp = new DateTime(data3, data2, 30);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 5:
                    {
                        temp = new DateTime(data3, data2, 31);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 6:
                    {
                        temp = new DateTime(data3, data2, 30);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 7:
                    {
                        temp = new DateTime(data3, data2, 31);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 8:
                    {
                        temp = new DateTime(data3, data2, 31);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 9:
                    {
                        temp = new DateTime(data3, data2, 30);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 10:
                    {
                        temp = new DateTime(data3, data2, 31);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 11:
                    {
                        temp = new DateTime(data3, data2, 30);
                        return temp.ToString("dd,MM,yyyy");
                    }
                case 12:
                    {
                        temp = new DateTime(data3, data2, 31);
                        return temp.ToString("dd,MM,yyyy");
                    }
            }
            temp = DateTime.Now;
            return temp.ToString();
            #endregion
        }

        public static void SelectUpdate(string str, string con)
        {
            OleDbConnection cn;
            cn = new OleDbConnection(con);
            cn.Open();
            OleDbCommand command = new OleDbCommand(str, cn);
            command.ExecuteNonQuery();
            cn.Close();
        }

        public static int INTSelectDBF(string str, string con)
        {
            OleDbConnection cn;
            cn = new OleDbConnection(con);
            cn.Open();
            OleDbCommand command = new OleDbCommand(str, cn);
            int SeleINT = Convert.ToInt32(command.ExecuteScalar());
            cn.Close();
            return SeleINT;
        }

        public static DataTable DTselect(string str, string con)
        {
            OleDbConnection cn;
            cn = new OleDbConnection(con);
            cn.Open();
            OleDbCommand command = new OleDbCommand(str, cn);
            DataTable temp = new DataTable();
            temp.Load(command.ExecuteReader());
            cn.Close();
            return temp;
        }
        #region Работа кнопки Выгрузка
        private void button4_Click(object sender, EventArgs e)
        {            
            string sourcefn, sourcefnPer;  //Имя файла,который копируем
            int kcname = INTSelectDBF("select kc from tabel order by kc", con);
            sourcefn = "T" + kcname;
            sourcefnPer = "P" + kcname+monthGlobal.ToString();
            CopyFile("d:\\TABEL\\TABEL.DBF", "\\\\Mztm\\Trmash_Data\\Maz\\NEWTABEL\\" + sourcefn + ".DBF");
            CopyFile("d:\\TABEL\\period.DBF", "\\\\Mztm\\Trmash_Data\\Maz\\NEWTABEL\\BasePeriod\\" + sourcefnPer + ".DBF");
            InsertCommonFile(sourcefn, kcname);
            CopyFile("d:\\TABEL\\COPY\\TABEL.DBF", "d:\\TABEL\\TABEL.DBF");
            CopyFile("d:\\TABEL\\COPY\\period.DBF", "d:\\TABEL\\period.DBF");
            button4.Text = "ВЫПОЛНЕНО";
        }
        #endregion

        #region Функция INSERT

        private void InsertCommonFile(string soursefile, int kcname)
        {
            //добавление записей в общую БД
            string conSRV = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Mztm\Trmash_Data\Maz\NEWTABEL\;Extended Properties=dBASE IV;User ID=Admin;Password=";

            string ClearingFile = "Delete * from TABEL_OK where kc=" + kcname;
            SelectUpdate(ClearingFile, conSRV);
            string INSERTCommon = "INSERT INTO TABEL_OK  SELECT * FROM " + soursefile;
            SelectUpdate(INSERTCommon, conSRV);
        }
        #endregion

        #region Функция копирования файлов

        void CopyFile(string sourcefn, string destinfn)
        {
            //sourcefn - файл,который копируем,с путем
            //destinfn - имя с путем куда копируем
            FileInfo fn = new FileInfo(sourcefn);
            fn.CopyTo(destinfn, true);
        }
        #endregion

        #region Кнопка выхода

        private void button5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        #endregion

        private void label3_Click(object sender, EventArgs e)
        {

        }

        #region Определение даты записи

        
        public static DateTime DateNew(int schet,int MounthBase,int YearBase,ref int superDay)
        {
 
            int Mounth = MounthBase;
            int Year = YearBase;
            if (schet % 6 == 0)
                superDay++;
            DateTime dateTimeEnd = new DateTime(Year, Mounth, superDay);
            return dateTimeEnd;
        }
        #endregion

        #region Функция обработки месяцев с 31 днем

        public static DateTime WorkBase(ref string cursor,string cursorTemp,DateTime dateTimeStartTemp, DateTime dateTimeEndTemp,int tabNumber,int kodCex)
        {
            //добавление записей о периодах
            string conSRV = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\TABEL\;Extended Properties=dBASE IV;User ID=Admin;Password=";

            DateTime dateTimeRes;

            //Обработка одинаковых дней и последнего числа месяца
            if (cursorTemp == cursor)
            {
                if (dateTimeEndTemp.Day == 31)
                {
                    string INSERTCommons = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeStartTemp}','{dateTimeEndTemp}','{cursor}',0,0,0)";
                    SelectUpdate(INSERTCommons, conSRV);
                    cursor = cursorTemp;
                    return dateTimeEndTemp;
                }
                return dateTimeStartTemp;
            }
            //Обработка последнего числа месяца
            if (dateTimeEndTemp.Day == 31)
            {
                dateTimeRes = new DateTime(dateTimeEndTemp.Year, dateTimeEndTemp.Month, dateTimeEndTemp.Day - 1);
                string INSERTCommons = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeStartTemp}','{dateTimeRes}','{cursor}',0,0,0)";
                string INSERTCommons2 = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeEndTemp}','{dateTimeEndTemp}','{cursorTemp}',0,0,0)";
                SelectUpdate(INSERTCommons, conSRV);
                SelectUpdate(INSERTCommons2, conSRV);
                cursor = cursorTemp;
                return dateTimeEndTemp;
            }

            //Основная зона обработки
            dateTimeRes = new DateTime(dateTimeEndTemp.Year, dateTimeEndTemp.Month, dateTimeEndTemp.Day-1);

            string INSERTCommon = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeStartTemp}','{dateTimeRes}','{cursor}',0,0,0)";
            cursor = cursorTemp;
            SelectUpdate(INSERTCommon, conSRV);
            return new DateTime(dateTimeEndTemp.Year, dateTimeEndTemp.Month, dateTimeEndTemp.Day);
        }

        #endregion

        #region Функция обработки месяцев с 30 днями

        public static DateTime WorkBaseTwo(ref string cursor, string cursorTemp, DateTime dateTimeStartTemp, DateTime dateTimeEndTemp, int tabNumber, int kodCex)
        {
            //добавление записей о периодах
            string conSRV = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\TABEL\;Extended Properties=dBASE IV;User ID=Admin;Password=";

            DateTime dateTimeRes;

            //Обработка одинаковых дней и последнего числа месяца
            if (cursorTemp == cursor)
            {
                if (dateTimeEndTemp.Day == 30)
                {
                    string INSERTCommons = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeStartTemp}','{dateTimeEndTemp}','{cursor}',0,0,0)";
                    SelectUpdate(INSERTCommons, conSRV);
                    cursor = cursorTemp;
                    return dateTimeEndTemp;
                }
                return dateTimeStartTemp;
            }
            //Обработка последнего числа месяца
            if (dateTimeEndTemp.Day == 30)
            {
                dateTimeRes = new DateTime(dateTimeEndTemp.Year, dateTimeEndTemp.Month, dateTimeEndTemp.Day - 1);
                string INSERTCommons = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeStartTemp}','{dateTimeRes}','{cursor}',0,0,0)";
                string INSERTCommons2 = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeEndTemp}','{dateTimeEndTemp}','{cursorTemp}',0,0,0)";
                SelectUpdate(INSERTCommons, conSRV);
                SelectUpdate(INSERTCommons2, conSRV);
                cursor = cursorTemp;
                return dateTimeEndTemp;
            }

            //Основная зона обработки
            dateTimeRes = new DateTime(dateTimeEndTemp.Year, dateTimeEndTemp.Month, dateTimeEndTemp.Day - 1);

            string INSERTCommon = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeStartTemp}','{dateTimeRes}','{cursor}',0,0,0)";
            cursor = cursorTemp;
            SelectUpdate(INSERTCommon, conSRV);
            return new DateTime(dateTimeEndTemp.Year, dateTimeEndTemp.Month, dateTimeEndTemp.Day);
        }

        #endregion

        #region Функция обработки февраля

        public static DateTime WorkBaseFEB(ref string cursor, string cursorTemp, DateTime dateTimeStartTemp, DateTime dateTimeEndTemp, int tabNumber, int kodCex)
        {
            //добавление записей о периодах
            string conSRV = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\TABEL\;Extended Properties=dBASE IV;User ID=Admin;Password=";

            DateTime dateTimeRes;

            //Обработка одинаковых дней и последнего числа месяца
            if (cursorTemp == cursor)
            {
                if (dateTimeEndTemp.Day == 30)
                {
                    string INSERTCommons = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeStartTemp}','{dateTimeEndTemp}','{cursor}',0,0,0)";
                    SelectUpdate(INSERTCommons, conSRV);
                    cursor = cursorTemp;
                    return dateTimeEndTemp;
                }
                return dateTimeStartTemp;
            }
            //Обработка последнего числа месяца
            if (dateTimeEndTemp.Day == 30)
            {
                dateTimeRes = new DateTime(dateTimeEndTemp.Year, dateTimeEndTemp.Month, dateTimeEndTemp.Day - 1);
                string INSERTCommons = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeStartTemp}','{dateTimeRes}','{cursor}',0,0,0)";
                string INSERTCommons2 = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeEndTemp}','{dateTimeEndTemp}','{cursorTemp}',0,0,0)";
                SelectUpdate(INSERTCommons, conSRV);
                SelectUpdate(INSERTCommons2, conSRV);
                cursor = cursorTemp;
                return dateTimeEndTemp;
            }

            //Основная зона обработки
            dateTimeRes = new DateTime(dateTimeEndTemp.Year, dateTimeEndTemp.Month, dateTimeEndTemp.Day - 1);

            string INSERTCommon = $"INSERT INTO PERIOD VALUES ({kodCex},{tabNumber},'{dateTimeStartTemp}','{dateTimeRes}','{cursor}',0,0,0)";
            cursor = cursorTemp;
            SelectUpdate(INSERTCommon, conSRV);
            return new DateTime(dateTimeEndTemp.Year, dateTimeEndTemp.Month, dateTimeEndTemp.Day);
        }

        #endregion
    }
}
