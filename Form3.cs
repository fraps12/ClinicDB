using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.IO;
using dbclass;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
namespace CLdb
{
    public partial class Form3 : Form
    {
       
        public Form3()
        {
            InitializeComponent();
        }
        String globalpath = @"C:\Users\" + Environment.UserName + @"\Documents\ClinicDB";
        private Excel.Application excelapp;
#region excelcellss definition
        private Excel.Range excelcells;
        private Excel.Range excelcells1;
        private Excel.Range excelcells2;
        private Excel.Range excelcells3;
        private Excel.Range excelcells4;
        private Excel.Range excelcells5;
        private Excel.Range excelcells6;
        private Excel.Range excelcells7;
        private Excel.Range excelcells8;
        private Excel.Range excelcells9;
        private Excel.Range excelcells10;
        private Excel.Range excelcells11;
        private Excel.Range excelcells12;
        private Excel.Range excelcells13;
        private Excel.Range excelcells14;
        private Excel.Range excelcells15;
        private Excel.Range excelcells16;
        private Excel.Range excelcells17;
        private Excel.Range excelcells18;
        private Excel.Range excelcells19;
        private Excel.Range excelcells20;
        private Excel.Range excelcells21;
        private Excel.Range excelcells22;
        private Excel.Range excelcells23;
        private Excel.Range excelcells24;
        private Excel.Range excelcells25;
        private Excel.Range excelcells26;
        private Excel.Range excelcells27;
        private Excel.Range excelcells28;
        private Excel.Range excelcells29;
        private Excel.Range excelcells30;
        private Excel.Range excelcells31;
        private Excel.Range excelcells32;
        private Excel.Range excelcells33;
        private Excel.Range excelcells34;
        private Excel.Range excelcells35;
        private Excel.Range excelcells36;
        private Excel.Range excelcells37;
        private Excel.Range excelcells38;
        private Excel.Range excelcells39;
        private Excel.Range excelcells40;
        private Excel.Range excelcells41;
        private Excel.Range excelcells42;
        private Excel.Range excelcells43;
        private Excel.Range excelcells44;
        private Excel.Range excelcells45;
        private Excel.Range excelcells46;
        private Excel.Range excelcells47;
        private Excel.Range excelcells48;
        private Excel.Range excelcells49;
        private Excel.Range excelcells50;
        private Excel.Range excelcells51;
        private Excel.Range excelcells52;
        private Excel.Range excelcells53;
        private Excel.Range excelcells54;
        private Excel.Range excelcells55;
        private Excel.Range excelcells56;
        private Excel.Range excelcells57;
        private Excel.Range excelcells58;
        private Excel.Range excelcells59;
        private Excel.Range excelcells60;
        private Excel.Range excelcells61;
        private Excel.Range excelcells62;
        private Excel.Range excelcells63;
        private Excel.Range excelcells64;
        private Excel.Range excelcells65;
        private Excel.Range excelcells66;
        private Excel.Range excelcells67;
        private Excel.Range excelcells68;
        private Excel.Range excelcells69;
        private Excel.Range excelcells70;
        private Excel.Range excelcells71;
        private Excel.Range excelcells72;
        private Excel.Range excelcells73;
        private Excel.Range excelcells74;
        private Excel.Range excelcells75;
        private Excel.Range excelcells76;
        private Excel.Range excelcells77;
        private Excel.Range excelcells78;
        private Excel.Range excelcells79;
        private Excel.Range excelcells80;
        private Excel.Range excelcells81;
        private Excel.Range excelcells82;
        private Excel.Range excelcells83;
        private Excel.Range excelcells84;
        private Excel.Range excelcells85;
        private Excel.Range excelcells86;
        private Excel.Range excelcells87;
        private Excel.Range excelcells88;
        private Excel.Range excelcells89;
        private Excel.Range excelcells90;
        private Excel.Range excelcells91;
        private Excel.Range excelcells92;
        #endregion





        private void Form3_Load(object sender, EventArgs e)
        {
            if (SelectedItemPath.FilePathExists.Contains(globalpath + @"\Biochimie\"))
            {
                tabControl1.SelectTab(0);
                excelapp = new Excel.Application();
                excelapp.Workbooks.Open(SelectedItemPath.FilePathExists);
                Excel.Workbook book = excelapp.ActiveWorkbook;
                Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[1];

                excelcells = sheet.get_Range("I12", "J12"); //Рэндж ячеек data recoltarii
                excelcells.Value = richTextBox4.Text;

                excelcells1 = sheet.get_Range("N12"); //Рэндж ячеек data receptionarii
                richTextBox7.Text = Convert.ToString(excelcells1.Value);

                excelcells2 = sheet.get_Range("K15"); //Рэндж для имени и фамилии, а так же имени файла
                richTextBox17.Text = Convert.ToString(excelcells2.Value);

                excelcells3 = sheet.get_Range("O15"); //Рэндж возраста
                richTextBox12.Text = Convert.ToString(excelcells3.Value);

                excelcells4 = sheet.get_Range("K17"); //Рэндж номера идентификации            
                richTextBox9.Text = Convert.ToString(excelcells4.Value);

                excelcells40 = sheet.get_Range("H97");//дата выдачи
                richTextBox99.Text = Convert.ToString(excelcells40.Value);

                excelcells5 = sheet.get_Range("O17"); //Рендж номера страховки(полица де асигураре)            
                richTextBox13.Text = Convert.ToString(excelcells5.Value);


                excelcells6 = sheet.get_Range("J19");// Рендж учреждения(институция)           
                richTextBox10.Text = Convert.ToString(excelcells6.Value);

                excelcells7 = sheet.get_Range("O19");//рендж отделения(секция)            
                richTextBox14.Text = Convert.ToString(excelcells7.Value);

                excelcells8 = sheet.get_Range("J21");//рендж участка            
                richTextBox11.Text = Convert.ToString(excelcells8.Value);

                excelcells9 = sheet.get_Range("N21");//рендж номера медкарты            
                richTextBox15.Text = Convert.ToString(excelcells9.Value);


                excelcells10 = sheet.get_Range("N9");//номер анализа
                richTextBox2.Text = Convert.ToString(excelcells10.Value);



                excelcells11 = sheet.get_Range("L35");
                richTextBox18.Text = Convert.ToString(excelcells11.Value);





                excelcells12 = sheet.get_Range("L37");
                richTextBox23.Text = Convert.ToString(excelcells12.Value);


                excelcells13 = sheet.get_Range("L39");
                richTextBox26.Text = Convert.ToString(excelcells13.Value);



                excelcells14 = sheet.get_Range("L41");
                richTextBox29.Text = Convert.ToString(excelcells14.Value);



                excelcells15 = sheet.get_Range("L43");
                richTextBox32.Text = Convert.ToString(excelcells15.Value);


                excelcells16 = sheet.get_Range("L45");
                richTextBox35.Text = Convert.ToString(excelcells16.Value);



                excelcells17 = sheet.get_Range("L47");
                richTextBox38.Text = Convert.ToString(excelcells17.Value);




                excelcells18 = sheet.get_Range("L49");
                richTextBox44.Text = Convert.ToString(excelcells18.Value);



                excelcells19 = sheet.get_Range("L51");
                richTextBox50.Text = Convert.ToString(excelcells19.Value);



                excelcells20 = sheet.get_Range("L53");
                richTextBox53.Text = Convert.ToString(excelcells20.Value);


                excelcells21 = sheet.get_Range("L55");
                richTextBox56.Text = Convert.ToString(excelcells21.Value);



                excelcells22 = sheet.get_Range("L57");
                richTextBox59.Text = Convert.ToString(excelcells22.Value);





                excelcells24 = sheet.get_Range("L61");
                richTextBox65.Text = Convert.ToString(excelcells24.Value);



                excelcells25 = sheet.get_Range("L63");
                richTextBox77.Text = Convert.ToString(excelcells25.Value);



                excelcells26 = sheet.get_Range("L65");
                richTextBox80.Text = Convert.ToString(excelcells26.Value);



                excelcells27 = sheet.get_Range("L67");
                richTextBox83.Text = Convert.ToString(excelcells27.Value);



                excelcells28 = sheet.get_Range("L69");
                richTextBox86.Text = Convert.ToString(excelcells28.Value);



                excelcells29 = sheet.get_Range("L71");
                richTextBox89.Text = Convert.ToString(excelcells29.Value);



                excelcells30 = sheet.get_Range("L73");
                richTextBox98.Text = Convert.ToString(excelcells30.Value);



                excelcells31 = sheet.get_Range("L77");
                richTextBox101.Text = Convert.ToString(excelcells31.Value);



                excelcells32 = sheet.get_Range("L81");
                richTextBox110.Text = Convert.ToString(excelcells32.Value);



                excelcells33 = sheet.get_Range("L83");
                richTextBox113.Text = Convert.ToString(excelcells33.Value);



                excelcells34 = sheet.get_Range("L59");
                richTextBox92.Text = Convert.ToString(excelcells34.Value);




                excelcells35 = sheet.get_Range("L75");
                richTextBox95.Text = Convert.ToString(excelcells35.Value);




                excelcells36 = sheet.get_Range("L79");
                richTextBox104.Text = Convert.ToString(excelcells36.Value);



                excelcells37 = sheet.get_Range("L87");
                richTextBox41.Text = Convert.ToString(excelcells37.Value);



                excelcells38 = sheet.get_Range("L89");
                richTextBox47.Text = Convert.ToString(excelcells38.Value);



                excelcells39 = sheet.get_Range("L91");
                richTextBox68.Text = Convert.ToString(excelcells39.Value);


            }
            if(SelectedItemPath.FilePathExists.Contains(globalpath + @"\Imunologie"))
            {
                tabControl1.SelectTab(1);
                excelapp = new Excel.Application();
                excelapp.Workbooks.Open(SelectedItemPath.FilePathExists);
                Excel.Workbook book = excelapp.ActiveWorkbook;
                Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[1];

                excelcells = sheet.get_Range("I12", "J12"); //Рэндж ячеек data recoltarii
                richTextBox31.Text = Convert.ToString(excelcells.Value);

                excelcells1 = sheet.get_Range("N12"); //Рэндж ячеек data receptionarii
                richTextBox28.Text = Convert.ToString(excelcells1.Value);

                excelcells2 = sheet.get_Range("K15"); //Рэндж для имени и фамилии, а так же имени файла
                richTextBox5.Text = Convert.ToString(excelcells2.Value);

                excelcells3 = sheet.get_Range("O15"); //Рэндж возраста
                richTextBox22.Text = Convert.ToString(excelcells3.Value);

                excelcells4 = sheet.get_Range("K17"); //Рэндж номера идентификации            
                richTextBox27.Text = Convert.ToString(excelcells4.Value);


                excelcells5 = sheet.get_Range("O17"); //Рендж номера страховки(полица де асигураре)            
                richTextBox21.Text = Convert.ToString(excelcells5.Value);

                excelcells80 = sheet.get_Range("H97");//дата выдачи
                richTextBox97.Text = Convert.ToString(excelcells80.Value);

                excelcells6 = sheet.get_Range("J19");// Рендж учреждения(институция)           
                richTextBox25.Text = Convert.ToString(excelcells6.Value);

                excelcells7 = sheet.get_Range("O19");//рендж отделения(секция)            
                richTextBox20.Text = Convert.ToString(excelcells7.Value);

                excelcells8 = sheet.get_Range("J21");//рендж участка            
                richTextBox24.Text = Convert.ToString(excelcells8.Value);

                excelcells9 = sheet.get_Range("N21");//рендж номера медкарты            
                richTextBox19.Text = Convert.ToString(excelcells9.Value);


                excelcells10 = sheet.get_Range("N9");//номер анализа
                richTextBox34.Text = Convert.ToString(excelcells10.Value);

                excelcells11 = sheet.get_Range("I29"); 
                comboBox1.Text = Convert.ToString(excelcells11.Value);

                excelcells12 = sheet.get_Range("I31");
                comboBox2.Text = Convert.ToString(excelcells12.Value);

                excelcells13 = sheet.get_Range("I33");
                comboBox3.Text = Convert.ToString(excelcells13.Value);

                excelcells14 = sheet.get_Range("I35");
                comboBox4.Text = Convert.ToString(excelcells14.Value);

                excelcells15 = sheet.get_Range("I37");
                comboBox5.Text = Convert.ToString(excelcells15.Value);

                excelcells16 = sheet.get_Range("I39");
                comboBox6.Text = Convert.ToString(excelcells16.Value);

                excelcells17 = sheet.get_Range("I41");
                comboBox7.Text = Convert.ToString(excelcells17.Value);

                excelcells18 = sheet.get_Range("I43");
                comboBox8.Text = Convert.ToString(excelcells11.Value);

                excelcells19 = sheet.get_Range("I45");
                comboBox9.Text = Convert.ToString(excelcells19.Value);

                excelcells20 = sheet.get_Range("I47");
                comboBox10.Text = Convert.ToString(excelcells20.Value);

                excelcells21 = sheet.get_Range("I49");
                comboBox11.Text = Convert.ToString(excelcells21.Value);

                excelcells22 = sheet.get_Range("I51");
                comboBox12.Text = Convert.ToString(excelcells22.Value);

                excelcells23 = sheet.get_Range("I53");
                comboBox24.Text = Convert.ToString(excelcells23.Value);

                excelcells25 = sheet.get_Range("I55");
                comboBox28.Text = Convert.ToString(excelcells25.Value);

                excelcells26 = sheet.get_Range("I57");
                comboBox30.Text = Convert.ToString(excelcells11.Value);

                excelcells27 = sheet.get_Range("I59");
                comboBox32.Text = Convert.ToString(excelcells27.Value);

                excelcells28 = sheet.get_Range("I61");
                comboBox34.Text = Convert.ToString(excelcells28.Value);



                //парсинг значений

                excelcells29 = sheet.get_Range("L29");
                richTextBox61.Text = Convert.ToString(excelcells29.Value);

                excelcells30 = sheet.get_Range("L31");
                richTextBox60.Text = Convert.ToString(excelcells30.Value);

                excelcells31 = sheet.get_Range("L33");
                richTextBox58.Text = Convert.ToString(excelcells31.Value);


                excelcells32 = sheet.get_Range("L35");
                richTextBox57.Text = Convert.ToString(excelcells32.Value);

                excelcells33 = sheet.get_Range("L37");
                richTextBox55.Text = Convert.ToString(excelcells33.Value);

                excelcells34 = sheet.get_Range("L39");
                richTextBox54.Text = Convert.ToString(excelcells34.Value);

                excelcells35 = sheet.get_Range("L41");
                richTextBox52.Text = Convert.ToString(excelcells35.Value);

                excelcells36 = sheet.get_Range("L43");
                richTextBox51.Text = Convert.ToString(excelcells36.Value);

                excelcells37 = sheet.get_Range("L45");
                richTextBox49.Text = Convert.ToString(excelcells37.Value);

                excelcells38 = sheet.get_Range("L47");
                richTextBox37.Text = Convert.ToString(excelcells38.Value);

                excelcells39 = sheet.get_Range("L49");
                richTextBox39.Text = Convert.ToString(excelcells39.Value);

                excelcells40 = sheet.get_Range("L51");
                richTextBox40.Text = Convert.ToString(excelcells40.Value);

                excelcells41 = sheet.get_Range("L53");
                richTextBox42.Text = Convert.ToString(excelcells41.Value);

                excelcells42 = sheet.get_Range("L55");
                richTextBox43.Text = Convert.ToString(excelcells42.Value);

                excelcells43 = sheet.get_Range("L57");
                richTextBox45.Text = Convert.ToString(excelcells29.Value);

                excelcells44 = sheet.get_Range("L59");
                richTextBox46.Text = Convert.ToString(excelcells44.Value);

                excelcells45 = sheet.get_Range("L61");
                richTextBox48.Text = Convert.ToString(excelcells29.Value);

                //парсинг нормы
                excelcells46 = sheet.get_Range("Q29");
                comboBox13.Text = Convert.ToString(excelcells46.Value);

                excelcells47 = sheet.get_Range("Q31");
                comboBox14.Text = Convert.ToString(excelcells47.Value);

                excelcells48 = sheet.get_Range("Q33");
                comboBox15.Text = Convert.ToString(excelcells48.Value);

                excelcells49 = sheet.get_Range("Q35");
                comboBox16.Text = Convert.ToString(excelcells49.Value);

                excelcells50 = sheet.get_Range("Q37");
                comboBox17.Text = Convert.ToString(excelcells50.Value);

                excelcells51 = sheet.get_Range("Q39");
                comboBox18.Text = Convert.ToString(excelcells51.Value);

                excelcells52 = sheet.get_Range("Q41");
                comboBox19.Text = Convert.ToString(excelcells52.Value);

                excelcells53 = sheet.get_Range("Q43");
                comboBox20.Text = Convert.ToString(excelcells53.Value);


                excelcells54 = sheet.get_Range("Q45");
                comboBox21.Text = Convert.ToString(excelcells54.Value);

                excelcells55 = sheet.get_Range("Q47");
                comboBox22.Text = Convert.ToString(excelcells55.Value);

                excelcells56 = sheet.get_Range("Q49");
                comboBox23.Text = Convert.ToString(excelcells56.Value);

                excelcells57 = sheet.get_Range("Q51");
                comboBox24.Text = Convert.ToString(excelcells57.Value);

                excelcells58 = sheet.get_Range("Q53");
                comboBox25.Text = Convert.ToString(excelcells58.Value);


                excelcells59 = sheet.get_Range("Q55");
                comboBox27.Text = Convert.ToString(excelcells59.Value);

                excelcells60 = sheet.get_Range("Q57");
                comboBox29.Text = Convert.ToString(excelcells60.Value);

                excelcells61 = sheet.get_Range("Q59");
                comboBox31.Text = Convert.ToString(excelcells61.Value);

                excelcells62 = sheet.get_Range("Q61");
                comboBox33.Text = Convert.ToString(excelcells62.Value);

                //Парсинг Интерпретации

                excelcells63 = sheet.get_Range("O29");
                comboBox40.Text = Convert.ToString(excelcells63.Value);

                excelcells64 = sheet.get_Range("O31");
                comboBox35.Text = Convert.ToString(excelcells64.Value);

                excelcells65 = sheet.get_Range("O33");
                comboBox36.Text = Convert.ToString(excelcells65.Value);

                excelcells66 = sheet.get_Range("O35");
                comboBox37.Text = Convert.ToString(excelcells66.Value);

                excelcells67 = sheet.get_Range("O37");
                comboBox42.Text = Convert.ToString(excelcells67.Value);

                excelcells68 = sheet.get_Range("O39");
                comboBox41.Text = Convert.ToString(excelcells68.Value);

                excelcells69 = sheet.get_Range("O41");
                comboBox39.Text = Convert.ToString(excelcells69.Value);

                excelcells70 = sheet.get_Range("O43");
                comboBox51.Text = Convert.ToString(excelcells70.Value);


                excelcells71 = sheet.get_Range("O45");
                comboBox38.Text = Convert.ToString(excelcells71.Value);

                excelcells72 = sheet.get_Range("O47");
                comboBox46.Text = Convert.ToString(excelcells72.Value);

                excelcells73 = sheet.get_Range("O49");
                comboBox45.Text = Convert.ToString(excelcells73.Value);

                excelcells74 = sheet.get_Range("O51");
                comboBox44.Text = Convert.ToString(excelcells74.Value);

                excelcells75 = sheet.get_Range("O53");
                comboBox43.Text = Convert.ToString(excelcells75.Value);


                excelcells76 = sheet.get_Range("O55");
                comboBox50.Text = Convert.ToString(excelcells76.Value);

                excelcells77 = sheet.get_Range("O57");
                comboBox49.Text = Convert.ToString(excelcells77.Value);

                excelcells78 = sheet.get_Range("O59");
                comboBox48.Text = Convert.ToString(excelcells78.Value);

                excelcells79 = sheet.get_Range("O61");
                comboBox47.Text = Convert.ToString(excelcells79.Value);


            }
            if (SelectedItemPath.FilePathExists.Contains(globalpath + @"\Reumo.Probe"))
            {
                tabControl1.SelectTab(2);
                excelapp = new Excel.Application();
                excelapp.Workbooks.Open(SelectedItemPath.FilePathExists);
                Excel.Workbook book = excelapp.ActiveWorkbook;
                Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[1];

                excelcells = sheet.get_Range("I12"); //Рэндж ячеек data recoltarii
                richTextBox75.Text = Convert.ToString(excelcells.Value);

                excelcells1 = sheet.get_Range("N12"); //Рэндж ячеек data receptionarii
                richTextBox73.Text = Convert.ToString(excelcells1.Value);

                excelcells2 = sheet.get_Range("J15"); //Рэндж для имени и фамилии, а так же имени файла
                richTextBox62.Text = Convert.ToString(excelcells2.Value);

                excelcells3 = sheet.get_Range("O15"); //Рэндж возраста
                richTextBox69.Text = Convert.ToString(excelcells3.Value);

                excelcells4 = sheet.get_Range("J17"); //Рэндж номера идентификации            
                richTextBox72.Text = Convert.ToString(excelcells4.Value);


                excelcells5 = sheet.get_Range("O17"); //Рендж номера страховки(полица де асигураре)            
                richTextBox67.Text = Convert.ToString(excelcells5.Value);


                excelcells6 = sheet.get_Range("J19");// Рендж учреждения(институция)           
                richTextBox71.Text = Convert.ToString(excelcells6.Value);

                excelcells7 = sheet.get_Range("O19");//рендж отделения(секция)            
                richTextBox66.Text = Convert.ToString(excelcells7.Value);

                excelcells8 = sheet.get_Range("J21");//рендж участка            
                richTextBox70.Text = Convert.ToString(excelcells8.Value);

                excelcells9 = sheet.get_Range("N21");//рендж номера медкарты            
                richTextBox64.Text = Convert.ToString(excelcells9.Value);


                excelcells10 = sheet.get_Range("Q9");//номер анализа
                richTextBox78.Text = Convert.ToString(excelcells10.Value);


                excelcells11 = sheet.get_Range("H29");
                comboBox61.Text = Convert.ToString(excelcells11.Value);

                excelcells12 = sheet.get_Range("H31");
                comboBox60.Text = Convert.ToString(excelcells12.Value);

                excelcells13 = sheet.get_Range("H33");
                comboBox59.Text = Convert.ToString(excelcells13.Value);

                excelcells14 = sheet.get_Range("H35");
                comboBox58.Text = Convert.ToString(excelcells14.Value);

                excelcells15 = sheet.get_Range("H37");
                comboBox57.Text = Convert.ToString(excelcells15.Value);

                excelcells83 = sheet.get_Range("H43");
                richTextBox96.Text = Convert.ToString(excelcells83.Value); //Парсинг в дату выдаче
                /////////////////////////////////////////////////////////


                excelcells16 = sheet.get_Range("K29");
                richTextBox81.Text = Convert.ToString(excelcells16.Value);

                excelcells17 = sheet.get_Range("K31");
                richTextBox82.Text = Convert.ToString(excelcells17.Value);

                excelcells18 = sheet.get_Range("K33");
                richTextBox84.Text = Convert.ToString(excelcells18.Value);

                excelcells19 = sheet.get_Range("K35");
                richTextBox85.Text = Convert.ToString(excelcells19.Value);

                excelcells20 = sheet.get_Range("K37");
                richTextBox87.Text = Convert.ToString(excelcells20.Value);
                ///////////////////////////////////////////////////////



                excelcells21 = sheet.get_Range("M29");
                richTextBox88.Text = Convert.ToString(excelcells21.Value);

                excelcells22 = sheet.get_Range("M31");
                richTextBox90.Text = Convert.ToString(excelcells22.Value);

                excelcells23 = sheet.get_Range("M33");
                richTextBox91.Text = Convert.ToString(excelcells23.Value);

                excelcells25 = sheet.get_Range("M35");
                richTextBox93.Text = Convert.ToString(excelcells25.Value);

                excelcells26 = sheet.get_Range("M37");
                richTextBox94.Text = Convert.ToString(excelcells26.Value);
                //////////////////////////////////////////////////////////


                excelcells27 = sheet.get_Range("N29");
                comboBox56.Text = Convert.ToString(excelcells27.Value);


                excelcells28 = sheet.get_Range("N31");
                comboBox55.Text = Convert.ToString(excelcells28.Value);

                excelcells29 = sheet.get_Range("N33");
                comboBox54.Text = Convert.ToString(excelcells29.Value);

                excelcells30 = sheet.get_Range("N35");
                comboBox53.Text = Convert.ToString(excelcells30.Value);

                excelcells31 = sheet.get_Range("N37");
                comboBox52.Text = Convert.ToString(excelcells31.Value);

            }
        }

        private void button1_Click(object sender, EventArgs e) //сохранить на биохимии
        {
            //excelapp = new Excel.Application();
            excelapp.Workbooks.Open(SelectedItemPath.FilePathExists);
            Excel.Workbook book = excelapp.ActiveWorkbook;
            Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[1];

            excelcells = sheet.get_Range("I12", "J12"); //Рэндж ячеек data recoltarii
            excelcells.Value = richTextBox4.Text;



            excelcells1 = sheet.get_Range("N12", "O12"); //Рэндж ячеек data receptionarii
            excelcells1.Value = richTextBox7.Text;

            excelcells2 = sheet.get_Range("K15"); //Рэндж для имени и фамилии, а так же имени файла
            excelcells2.Value = richTextBox17.Text;

            excelcells3 = sheet.get_Range("O15"); //Рэндж возраста
            excelcells3.Value = richTextBox12.Text;

            excelcells4 = sheet.get_Range("K17"); //Рэндж номера идентификации
            excelcells4.Value = Convert.ToString(richTextBox9.Text);

            excelcells88 = sheet.get_Range("H97", "M97");//дата выдачи
            excelcells88.Value = richTextBox99.Text;

            excelcells5 = sheet.get_Range("O17", "P17"); //Рендж номера страховки(полица де асигураре)
            excelcells5.Value = richTextBox13.Text;
            

            excelcells6 = sheet.get_Range("J19", "K19");// Рендж учреждения(институция)
            excelcells6.Value = richTextBox10.Text;

            excelcells7 = sheet.get_Range("O19", "P19");//рендж отделения(секция)
            excelcells7.Value = richTextBox14.Text;

            excelcells8 = sheet.get_Range("J21", "K21");//рендж участка
            excelcells8.Value = richTextBox11.Text;

            excelcells9 = sheet.get_Range("N21", "O21");//рендж номера медкарты
            excelcells9.Value = richTextBox15.Text;
            excelcells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            excelcells10 = sheet.get_Range("N9");//номер анализа
            excelcells10.Value = richTextBox2.Text;

            if (String.IsNullOrWhiteSpace(richTextBox18.Text)) //18 = proteina totala
            {
                var testRange1 = sheet.Range[sheet.Cells[35, 8], sheet.Cells[35, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[36, 8], sheet.Cells[36, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells11 = sheet.get_Range("L35", "M36");
            excelcells11.Value = richTextBox18.Text;



            if (String.IsNullOrWhiteSpace(richTextBox23.Text))//23 = Albumina
            {
                var testRange1 = sheet.Range[sheet.Cells[37, 8], sheet.Cells[37, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[38, 8], sheet.Cells[38, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells12 = sheet.get_Range("L37", "M38");
            excelcells12.Value = richTextBox23.Text;



            if (String.IsNullOrWhiteSpace(richTextBox26.Text))//26 = uree
            {
                var testRange1 = sheet.Range[sheet.Cells[39, 8], sheet.Cells[39, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[40, 8], sheet.Cells[40, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells13 = sheet.get_Range("L39", "M40");
            excelcells13.Value = richTextBox26.Text;

            if (String.IsNullOrWhiteSpace(richTextBox29.Text))// 29 = Kreatinina
            {
                var testRange1 = sheet.Range[sheet.Cells[41, 8], sheet.Cells[41, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[42, 8], sheet.Cells[42, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells14 = sheet.get_Range("L41", "M42");
            excelcells14.Value = richTextBox29.Text;

            if (String.IsNullOrWhiteSpace(richTextBox32.Text))//32 = accid uric
            {
                var testRange1 = sheet.Range[sheet.Cells[43, 8], sheet.Cells[43, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[44, 8], sheet.Cells[44, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells15 = sheet.get_Range("L43", "M44");
            excelcells15.Value = richTextBox32.Text;

            if (String.IsNullOrWhiteSpace(richTextBox35.Text))//35 = bilirubiina totala
            {
                var testRange1 = sheet.Range[sheet.Cells[45, 8], sheet.Cells[45, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[46, 8], sheet.Cells[46, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells16 = sheet.get_Range("L45", "M46");
            excelcells16.Value = richTextBox35.Text;

            if (String.IsNullOrWhiteSpace(richTextBox38.Text))//38 = bilirubina conjugata
            {
                var testRange1 = sheet.Range[sheet.Cells[47, 8], sheet.Cells[47, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[48, 8], sheet.Cells[48, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells17 = sheet.get_Range("L47", "M48");
            excelcells17.Value = richTextBox38.Text;


            if (String.IsNullOrWhiteSpace(richTextBox44.Text)) // 44 = glucoza
            {
                var testRange1 = sheet.Range[sheet.Cells[49, 8], sheet.Cells[49, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[50, 8], sheet.Cells[50, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells18 = sheet.get_Range("L49", "M50");
            excelcells18.Value = richTextBox44.Text;

            if (String.IsNullOrWhiteSpace(richTextBox50.Text))//50 = ALAT
            {
                var testRange1 = sheet.Range[sheet.Cells[51, 8], sheet.Cells[51, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[52, 8], sheet.Cells[52, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells19 = sheet.get_Range("L51", "M52");
            excelcells19.Value = richTextBox50.Text;

            if (String.IsNullOrWhiteSpace(richTextBox53.Text))//53 = ASAT
            {
                var testRange1 = sheet.Range[sheet.Cells[53, 8], sheet.Cells[53, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[54, 8], sheet.Cells[54, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells20 = sheet.get_Range("L53", "M54");
            excelcells20.Value = richTextBox53.Text;

            if (String.IsNullOrWhiteSpace(richTextBox56.Text)) //56 = Amilaza
            {
                var testRange1 = sheet.Range[sheet.Cells[55, 8], sheet.Cells[55, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[56, 8], sheet.Cells[56, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells21 = sheet.get_Range("L55", "M56");
            excelcells21.Value = richTextBox56.Text;

            if (String.IsNullOrWhiteSpace(richTextBox59.Text))//59 = lipaza
            {
                var testRange1 = sheet.Range[sheet.Cells[57, 8], sheet.Cells[57, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[58, 8], sheet.Cells[58, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells22 = sheet.get_Range("L57", "M58");
            excelcells22.Value = richTextBox59.Text;

            

            if (String.IsNullOrWhiteSpace(richTextBox65.Text))//65 = Lactat dehidrogenaza
            {
                var testRange1 = sheet.Range[sheet.Cells[61, 8], sheet.Cells[61, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[62, 8], sheet.Cells[62, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells24 = sheet.get_Range("L61", "M62");
            excelcells24.Value = richTextBox65.Text;


            if (String.IsNullOrWhiteSpace(richTextBox77.Text))// glutamil trans peptidaza
            {
                var testRange1 = sheet.Range[sheet.Cells[63, 8], sheet.Cells[63, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[64, 8], sheet.Cells[64, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells25 = sheet.get_Range("L63", "M64");
            excelcells25.Value = richTextBox77.Text;

            if (String.IsNullOrWhiteSpace(richTextBox80.Text))//80 = colesterol total
            {
                var testRange1 = sheet.Range[sheet.Cells[65, 8], sheet.Cells[65, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[66, 8], sheet.Cells[66, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells26 = sheet.get_Range("L65", "M66");
            excelcells26.Value = richTextBox80.Text;

            if (String.IsNullOrWhiteSpace(richTextBox83.Text))// Trigliceride
            {
                var testRange1 = sheet.Range[sheet.Cells[67, 8], sheet.Cells[67, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[68, 8], sheet.Cells[68, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells27 = sheet.get_Range("L67", "M68");
            excelcells27.Value = richTextBox83.Text;

            if (String.IsNullOrWhiteSpace(richTextBox86.Text))//Colesterol densitatea inalta
            {
                var testRange1 = sheet.Range[sheet.Cells[69, 8], sheet.Cells[69, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[70, 8], sheet.Cells[70, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells28 = sheet.get_Range("L69", "M70");
            excelcells28.Value = richTextBox86.Text;

            if (String.IsNullOrWhiteSpace(richTextBox89.Text))//colesterol densitatea joasa
            {
                var testRange1 = sheet.Range[sheet.Cells[71, 8], sheet.Cells[71, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[72, 8], sheet.Cells[72, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells29 = sheet.get_Range("L71", "M72");
            excelcells29.Value = richTextBox89.Text;

            if (String.IsNullOrWhiteSpace(richTextBox98.Text))//Potasiu
            {
                var testRange1 = sheet.Range[sheet.Cells[73, 8], sheet.Cells[73, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[74, 8], sheet.Cells[74, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells30 = sheet.get_Range("L73", "M74");
            excelcells30.Value = richTextBox98.Text;

            if (String.IsNullOrWhiteSpace(richTextBox101.Text))//Calciu
            {
                var testRange1 = sheet.Range[sheet.Cells[77, 8], sheet.Cells[77, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[78, 8], sheet.Cells[78, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells31 = sheet.get_Range("L77", "M78");
            excelcells31.Value = richTextBox101.Text;


            if (String.IsNullOrWhiteSpace(richTextBox110.Text))//Fier
            {
                var testRange1 = sheet.Range[sheet.Cells[81, 8], sheet.Cells[81, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[82, 8], sheet.Cells[82, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells32 = sheet.get_Range("L81", "M82");
            excelcells32.Value = richTextBox110.Text;

            if (String.IsNullOrWhiteSpace(richTextBox113.Text))//Magneziu
            {
                var testRange1 = sheet.Range[sheet.Cells[83, 8], sheet.Cells[83, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[84, 8], sheet.Cells[84, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells33 = sheet.get_Range("L83", "M84");
            excelcells33.Value = richTextBox113.Text;

            if (String.IsNullOrWhiteSpace(richTextBox92.Text))//fosfataza alcolina
            {
                var testRange1 = sheet.Range[sheet.Cells[59, 8], sheet.Cells[59, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[60, 8], sheet.Cells[60, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells34 = sheet.get_Range("L59", "M60");
            excelcells34.Value = richTextBox92.Text;


            if (String.IsNullOrWhiteSpace(richTextBox95.Text))//Sodiu
            {
                var testRange1 = sheet.Range[sheet.Cells[75, 8], sheet.Cells[75, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[76, 8], sheet.Cells[76, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells35 = sheet.get_Range("L75", "M76");
            excelcells35.Value = richTextBox95.Text;


            if (String.IsNullOrWhiteSpace(richTextBox104.Text))//Anorganic fosfor
            {
                var testRange1 = sheet.Range[sheet.Cells[79, 8], sheet.Cells[79, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[80, 8], sheet.Cells[80, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells36 = sheet.get_Range("L79", "M80");
            excelcells36.Value = richTextBox104.Text;

            if (String.IsNullOrWhiteSpace(richTextBox41.Text))//Antistreptolizina-O
            {
                var testRange1 = sheet.Range[sheet.Cells[87, 8], sheet.Cells[87, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[88, 8], sheet.Cells[88, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells37 = sheet.get_Range("L87", "M88");
            excelcells37.Value = richTextBox41.Text;

            if (String.IsNullOrWhiteSpace(richTextBox47.Text))//C-Protein reactiv
            {
                var testRange1 = sheet.Range[sheet.Cells[89, 8], sheet.Cells[89, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[90, 8], sheet.Cells[90, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells38 = sheet.get_Range("L89", "M90");
            excelcells38.Value = richTextBox47.Text;

            if (String.IsNullOrWhiteSpace(richTextBox68.Text))//latex test
            {
                var testRange1 = sheet.Range[sheet.Cells[91, 8], sheet.Cells[91, 16]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[92, 8], sheet.Cells[92, 16]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells39 = sheet.get_Range("L91", "M92");
            excelcells39.Value = richTextBox68.Text;
            String filename;
            if(File.Exists(globalpath + @"\Biochimie\" + SelectedItemPath.SelItemPath) == true) //сохранения в папку биохимия, тк, там файл изменялся там
            {
                String path = globalpath;
                filename = Convert.ToString(richTextBox17.Text);                
                book.SaveAs(Filename: path + @"\Biochimie\" + filename + ".xlsx");
                excelapp.Quit();
            }

            if (File.Exists(globalpath + @"\Imunologie\" + SelectedItemPath.SelItemPath) == true)//сохранения в папку имунология, тк, там файл изменялся там
            {
                String path1 = globalpath;
                filename = Convert.ToString(richTextBox17.Text);                
                book.SaveAs(Filename: path1 + @"\Imunologie\" + filename + ".xlsx");
                excelapp.Quit();
            }

            if (File.Exists(globalpath + @"\Reumo.Probe\" + SelectedItemPath.SelItemPath) == true)//сохранения в папку ревмо, тк, там файл изменялся там
            {
                String path2 = globalpath;                
                filename = Convert.ToString(richTextBox17.Text);
                book.SaveAs(Filename: path2 + @"\Reumo.Probe\" + filename + ".xlsx");
                excelapp.Quit();

            }
            String str1 = DateTime.Now.ToString("yy.MM.dd HH.mm.ss");
            if (String.IsNullOrWhiteSpace(richTextBox17.Text))
            {
                filename = "default" + str1;
                book.SaveAs(Filename: globalpath + @"\Unnamed\" + filename + ".xlsx");
            }
            
        }


        public void SelectTab(int index)
        {
            tabControl1.SelectTab(1);
        }

        String filename1;
        private Excel.Application excelapp1;
        //public Excel.Sheets excelSheets;
        public Excel.Worksheet sheet;

        private Excel.Workbook excelAppWorkbook;
        private Excel.Workbooks excelAppWorkbooks;

        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            excelapp1.Quit();
        }

        private void button4_Click(object sender, EventArgs e)//КНОПКА сохранить на вкладке ИММУНОЛОГИЯ
        {
            excelapp1 = new Excel.Application();//если не открыто, создаем новое
            excelAppWorkbooks = excelapp.Workbooks;
            excelAppWorkbook = excelAppWorkbooks.Open(globalpath + @"\ImunoTemplateFinal.xlsx");


            Excel.Workbook excelSheets = excelapp.ActiveWorkbook;
            sheet = (Excel.Worksheet)excelSheets.Worksheets[1];




            //sheet.Cells[5, 15] = richTextBox2.Text;
            filename1 = richTextBox5.Text;
            excelcells = sheet.get_Range("I12", "J12"); //Рэндж ячеек data recoltarii
            excelcells.Value = richTextBox31.Text;

            excelcells40 = sheet.get_Range("H97", "M97");//дата выдачи
            excelcells40.Value = richTextBox99.Text;

            excelcells1 = sheet.get_Range("N12", "O12"); //Рэндж ячеек data receptionarii
            excelcells1.Value = richTextBox28.Text;

            excelcells2 = sheet.get_Range("K15"); //Рэндж для имени и фамилии, а так же имени файла
            excelcells2.Value = richTextBox5.Text;

            excelcells3 = sheet.get_Range("O15"); //Рэндж возраста
            excelcells3.Value = richTextBox22.Text;

            excelcells4 = sheet.get_Range("K17"); //Рэндж номера идентификации
            excelcells4.Value = Convert.ToString(richTextBox27.Text);


            excelcells5 = sheet.get_Range("O17", "P17"); //Рендж номера страховки(полица де асигураре)
            excelcells5.Value = richTextBox21.Text;


            excelcells6 = sheet.get_Range("J19", "K19");// Рендж учреждения(институция)
            excelcells6.Value = richTextBox25.Text;

            excelcells7 = sheet.get_Range("O19", "P19");//рендж отделения(секция)
            excelcells7.Value = richTextBox20.Text;

            excelcells8 = sheet.get_Range("J21", "K21");//рендж участка
            excelcells8.Value = richTextBox24.Text;

            excelcells9 = sheet.get_Range("N21", "O21");//рендж номера медкарты
            excelcells9.Value = richTextBox19.Text;

            excelcells83 = sheet.get_Range("H65", "M65");//дата выдачи
            excelcells83.Value = richTextBox97.Text;


            excelcells10 = sheet.get_Range("N9");//номер анализа
            excelcells10.Value = richTextBox34.Text;

            if (String.IsNullOrWhiteSpace(comboBox1.Text)) //1 = первый элемент и так далее
            {
                var testRange1 = sheet.Range[sheet.Cells[29, 9], sheet.Cells[29, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[30, 9], sheet.Cells[30, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells11 = sheet.get_Range("I29", "K30");
            excelcells11.Value = comboBox1.Text;



            if (String.IsNullOrWhiteSpace(comboBox2.Text))//23 = Albumina
            {
                var testRange1 = sheet.Range[sheet.Cells[31, 9], sheet.Cells[31, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[32, 9], sheet.Cells[32, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells12 = sheet.get_Range("I31", "K32");
            excelcells12.Value = comboBox2.Text;



            if (String.IsNullOrWhiteSpace(comboBox3.Text))//26 = uree
            {
                var testRange1 = sheet.Range[sheet.Cells[33, 9], sheet.Cells[33, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[34, 9], sheet.Cells[34, 18]];
                testRange2.EntireRow.Hidden = true;
            }

            excelcells13 = sheet.get_Range("I33", "K34");
            excelcells13.Value = comboBox3.Text;

            if (String.IsNullOrWhiteSpace(comboBox4.Text))// 29 = Kreatinina
            {
                var testRange1 = sheet.Range[sheet.Cells[35, 9], sheet.Cells[35, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[36, 9], sheet.Cells[36, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells14 = sheet.get_Range("I35", "K36");
            excelcells14.Value = comboBox4.Text;

            if (String.IsNullOrWhiteSpace(comboBox5.Text))//32 = accid uric
            {
                var testRange1 = sheet.Range[sheet.Cells[37, 9], sheet.Cells[37, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[38, 9], sheet.Cells[38, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells15 = sheet.get_Range("I37", "K38");
            excelcells15.Value = comboBox5.Text;

            if (String.IsNullOrWhiteSpace(comboBox6.Text))//35 = bilirubiina totala
            {
                var testRange1 = sheet.Range[sheet.Cells[39, 9], sheet.Cells[39, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[40, 9], sheet.Cells[40, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells16 = sheet.get_Range("I39", "K40");
            excelcells16.Value = comboBox6.Text;

            if (String.IsNullOrWhiteSpace(comboBox7.Text))//38 = bilirubina conjugata
            {
                var testRange1 = sheet.Range[sheet.Cells[41, 9], sheet.Cells[41, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[42, 9], sheet.Cells[42, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells17 = sheet.get_Range("I41", "K42");
            excelcells17.Value = comboBox7.Text;


            if (String.IsNullOrWhiteSpace(comboBox8.Text)) // 44 = glucoza
            {
                var testRange1 = sheet.Range[sheet.Cells[43, 9], sheet.Cells[43, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[44, 9], sheet.Cells[44, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells18 = sheet.get_Range("I43", "K44");
            excelcells18.Value = comboBox9.Text;

            if (String.IsNullOrWhiteSpace(comboBox9.Text))//50 = ALAT
            {
                var testRange1 = sheet.Range[sheet.Cells[45, 9], sheet.Cells[45, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[46, 9], sheet.Cells[46, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells19 = sheet.get_Range("I45", "K46");
            excelcells19.Value = comboBox9.Text;

            if (String.IsNullOrWhiteSpace(comboBox10.Text))//53 = ASAT
            {
                var testRange1 = sheet.Range[sheet.Cells[47, 9], sheet.Cells[47, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[48, 9], sheet.Cells[48, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells20 = sheet.get_Range("I47", "K48");
            excelcells20.Value = comboBox10.Text;

            if (String.IsNullOrWhiteSpace(comboBox11.Text)) //56 = Amilaza
            {
                var testRange1 = sheet.Range[sheet.Cells[49, 9], sheet.Cells[49, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[50, 9], sheet.Cells[50, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells21 = sheet.get_Range("I49", "K50");
            excelcells21.Value = comboBox11.Text;

            if (String.IsNullOrWhiteSpace(comboBox12.Text))//59 = lipaza
            {
                var testRange1 = sheet.Range[sheet.Cells[51, 9], sheet.Cells[51, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[52, 9], sheet.Cells[52, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells22 = sheet.get_Range("I51", "K52");
            excelcells22.Value = comboBox12.Text;

            /* if (String.IsNullOrWhiteSpace(richTextBox62.Text))//62 = fosfataza alcolina
             {
                 var testRange1 = sheet.Range[sheet.Cells[59, 8], sheet.Cells[59, 16]];
                 testRange1.EntireRow.Hidden = true;
                 var testRange2 = sheet.Range[sheet.Cells[60, 8], sheet.Cells[60, 16]];
                 testRange2.EntireRow.Hidden = true;
             }
             excelcells23 = sheet.get_Range("L59", "M60");
             excelcells23.Value = richTextBox62.Text;
             */

            if (String.IsNullOrWhiteSpace(comboBox26.Text))//65 = Lactat dehidrogenaza
            {
                var testRange1 = sheet.Range[sheet.Cells[53, 9], sheet.Cells[53, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[54, 9], sheet.Cells[54, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells24 = sheet.get_Range("I53", "K54");
            excelcells24.Value = comboBox26.Text;


            if (String.IsNullOrWhiteSpace(comboBox28.Text))// glutamil trans peptidaza
            {
                var testRange1 = sheet.Range[sheet.Cells[55, 9], sheet.Cells[55, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[56, 9], sheet.Cells[56, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells25 = sheet.get_Range("I55", "K56");
            excelcells25.Value = comboBox28.Text;

            if (String.IsNullOrWhiteSpace(comboBox30.Text))//80 = colesterol total
            {
                var testRange1 = sheet.Range[sheet.Cells[57, 9], sheet.Cells[57, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[58, 9], sheet.Cells[58, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells26 = sheet.get_Range("I57", "K58");
            excelcells26.Value = comboBox30.Text;

            if (String.IsNullOrWhiteSpace(comboBox32.Text))// Trigliceride
            {
                var testRange1 = sheet.Range[sheet.Cells[59, 9], sheet.Cells[59, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[60, 9], sheet.Cells[60, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells27 = sheet.get_Range("I59", "K60");
            excelcells27.Value = comboBox32.Text;

            if (String.IsNullOrWhiteSpace(comboBox34.Text))//Colesterol densitatea inalta
            {
                var testRange1 = sheet.Range[sheet.Cells[61, 9], sheet.Cells[61, 18]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[62, 9], sheet.Cells[62, 18]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells28 = sheet.get_Range("I61", "K62");
            excelcells28.Value = comboBox34.Text;
            //till here заполняются все колонки названий


            excelcells29 = sheet.get_Range("L29", "M30"); //from here начинается заполнение колонки со значениями
            excelcells29.Value = richTextBox61.Text;


            excelcells30 = sheet.get_Range("L31", "M32");
            excelcells30.Value = richTextBox60.Text;


            excelcells31 = sheet.get_Range("L33", "M34");
            excelcells31.Value = richTextBox58.Text;



            excelcells32 = sheet.get_Range("L35", "M36");
            excelcells32.Value = richTextBox57.Text;

            excelcells33 = sheet.get_Range("L37", "M38");
            excelcells33.Value = richTextBox55.Text;


            excelcells34 = sheet.get_Range("L39", "M40");
            excelcells34.Value = richTextBox54.Text;



            excelcells35 = sheet.get_Range("L41", "M42");
            excelcells35.Value = richTextBox52.Text;


            excelcells36 = sheet.get_Range("L43", "M44");
            excelcells36.Value = richTextBox51.Text;


            excelcells37 = sheet.get_Range("L45", "M46");
            excelcells37.Value = richTextBox49.Text;


            excelcells38 = sheet.get_Range("L47", "M48");
            excelcells38.Value = richTextBox37.Text;


            excelcells39 = sheet.get_Range("L49", "M50");
            excelcells39.Value = richTextBox39.Text;


            excelcells40 = sheet.get_Range("L51", "M52");
            excelcells40.Value = richTextBox40.Text;

            excelcells41 = sheet.get_Range("L53", "M54");
            excelcells41.Value = richTextBox42.Text;

            excelcells42 = sheet.get_Range("L55", "M56");
            excelcells42.Value = richTextBox43.Text;

            excelcells43 = sheet.get_Range("L57", "M58");
            excelcells43.Value = richTextBox45.Text;

            excelcells44 = sheet.get_Range("L59", "M60");
            excelcells44.Value = richTextBox46.Text;

            excelcells45 = sheet.get_Range("L61", "M62");
            excelcells45.Value = richTextBox48.Text;

            //закончилось заполнение результатов, если заполнены все TextBox

            //Началось заполнение колонок нормы
            excelcells46 = sheet.get_Range("Q29", "R30");
            excelcells46.Value = comboBox13.Text;

            excelcells47 = sheet.get_Range("Q31", "R32");
            excelcells47.Value = comboBox14.Text;

            excelcells48 = sheet.get_Range("Q33", "R34");
            excelcells48.Value = comboBox15.Text;

            excelcells49 = sheet.get_Range("Q35", "R36");
            excelcells49.Value = comboBox16.Text;

            excelcells50 = sheet.get_Range("Q37", "R38");
            excelcells50.Value = comboBox17.Text;

            excelcells51 = sheet.get_Range("Q39", "R40");
            excelcells51.Value = comboBox18.Text;

            excelcells52 = sheet.get_Range("Q41", "R42");
            excelcells52.Value = comboBox19.Text;

            excelcells53 = sheet.get_Range("Q43", "R44");
            excelcells53.Value = comboBox20.Text;

            excelcells54 = sheet.get_Range("Q45", "R46");
            excelcells54.Value = comboBox21.Text;

            excelcells55 = sheet.get_Range("Q47", "R48");
            excelcells55.Value = comboBox22.Text;

            excelcells56 = sheet.get_Range("Q49", "R50");
            excelcells56.Value = comboBox23.Text;

            excelcells57 = sheet.get_Range("Q51", "R52");
            excelcells57.Value = comboBox24.Text;

            excelcells58 = sheet.get_Range("Q53", "R54");
            excelcells58.Value = comboBox25.Text;

            excelcells59 = sheet.get_Range("Q55", "R56");
            excelcells59.Value = comboBox27.Text;

            excelcells60 = sheet.get_Range("Q57", "R58");
            excelcells60.Value = comboBox29.Text;

            excelcells61 = sheet.get_Range("Q59", "R60");
            excelcells61.Value = comboBox31.Text;

            excelcells62 = sheet.get_Range("Q61", "R62");
            excelcells62.Value = comboBox33.Text;

            //Закончилось заполнение норм


            //началась проверка CheckBox-ов

            //61/13
            if (richTextBox61.Text != String.Empty && comboBox13.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox61.Text) > Convert.ToDouble(comboBox13.Text))
                {
                    excelcells63 = sheet.get_Range("O29", "P30");
                    excelcells63.Value = "POZITIV";
                }
            }


            if (richTextBox60.Text != String.Empty && comboBox14.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox60.Text) > Convert.ToDouble(comboBox14.Text))
                {
                    excelcells64 = sheet.get_Range("O31", "P32");
                    excelcells64.Value = "POZITIV";
                }
            }

            if (richTextBox58.Text != String.Empty && comboBox15.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox58.Text) > Convert.ToDouble(comboBox15.Text))
                {
                    excelcells65 = sheet.get_Range("O33", "P34");
                    excelcells65.Value = "POZITIV";
                }

            }

            if (richTextBox57.Text != String.Empty && comboBox16.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox57.Text) > Convert.ToDouble(comboBox16.Text))
                {
                    excelcells66 = sheet.get_Range("O35", "P36");
                    excelcells66.Value = "POZITIV";
                }
            }





            if (richTextBox55.Text != String.Empty && comboBox17.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox55.Text) > Convert.ToDouble(comboBox17.Text))
                {
                    excelcells67 = sheet.get_Range("O37", "P38");
                    excelcells67.Value = "POZITIV";
                }
            }


            if (richTextBox54.Text != String.Empty && comboBox18.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox54.Text) > Convert.ToDouble(comboBox18.Text))
                {
                    excelcells68 = sheet.get_Range("O39", "P40");
                    excelcells68.Value = "POZITIV";
                }
            }

            if (richTextBox52.Text != String.Empty && comboBox19.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox52.Text) > Convert.ToDouble(comboBox19.Text))
                {
                    excelcells69 = sheet.get_Range("O41", "P42");
                    excelcells69.Value = "POZITIV";
                }
            }



            if (richTextBox51.Text != String.Empty && comboBox20.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox51.Text) > Convert.ToDouble(comboBox20.Text))
                {
                    excelcells70 = sheet.get_Range("O43", "P44");
                    excelcells70.Value = "POZITIV";
                }
            }


            if (richTextBox49.Text != String.Empty && comboBox21.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox49.Text) > Convert.ToDouble(comboBox21.Text))
                {
                    excelcells71 = sheet.get_Range("O45", "P46");
                    excelcells71.Value = "POZITIV";
                }
            }


            if (richTextBox37.Text != String.Empty && comboBox22.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox37.Text) > Convert.ToDouble(comboBox22.Text))
                {
                    excelcells72 = sheet.get_Range("O47", "P48");
                    excelcells72.Value = "POZITIV";
                }
            }


            if (richTextBox39.Text != String.Empty && comboBox23.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox39.Text) > Convert.ToDouble(comboBox23.Text))
                {
                    excelcells73 = sheet.get_Range("O49", "P50");
                    excelcells73.Value = "POZITIV";
                }
            }


            if (richTextBox40.Text != String.Empty && comboBox24.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox40.Text) > Convert.ToDouble(comboBox24.Text))
                {
                    excelcells74 = sheet.get_Range("O51", "P52");
                    excelcells74.Value = "POZITIV";
                }
            }


            if (richTextBox42.Text != String.Empty && comboBox25.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox42.Text) > Convert.ToDouble(comboBox25.Text))
                {
                    excelcells75 = sheet.get_Range("O53", "P54");
                    excelcells75.Value = "POZITIV";
                }
            }



            if (richTextBox43.Text != String.Empty && comboBox27.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox43.Text) > Convert.ToDouble(comboBox27.Text))
                {
                    excelcells76 = sheet.get_Range("O55", "P56");
                    excelcells76.Value = "POZITIV";
                }
            }

            if (richTextBox45.Text != String.Empty && comboBox29.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox45.Text) > Convert.ToDouble(comboBox29.Text))
                {
                    excelcells77 = sheet.get_Range("O57", "P58");
                    excelcells77.Value = "POZITIV";
                }
            }


            if (richTextBox46.Text != String.Empty && comboBox31.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox46.Text) > Convert.ToDouble(comboBox31.Text))
                {
                    excelcells78 = sheet.get_Range("O59", "P60");
                    excelcells78.Value = "POZITIV";
                }
            }


            if (richTextBox48.Text != String.Empty && comboBox33.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox48.Text) > Convert.ToDouble(comboBox33.Text))
                {
                    excelcells79 = sheet.get_Range("O61", "P62");
                    excelcells79.Value = "POZITIV";
                }
            }

            // NEGGGGGGGGGAAAAAAAAAATIIIIIIIIIIIIVVVVVVVVVVVVVVVV
            if (richTextBox61.Text != String.Empty && comboBox13.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox61.Text) < Convert.ToDouble(comboBox13.Text))
                {
                    excelcells80 = sheet.get_Range("O29", "P30");
                    excelcells80.Value = "NEGATIV";
                }
            }

            if (richTextBox60.Text != String.Empty && comboBox14.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox60.Text) < Convert.ToDouble(comboBox14.Text))
                {
                    excelcells80 = sheet.get_Range("O31", "P32");
                    excelcells80.Value = "NEGATIV";
                }
            }
            if (richTextBox58.Text != String.Empty && comboBox15.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox58.Text) < Convert.ToDouble(comboBox15.Text))
                {
                    excelcells80 = sheet.get_Range("O33", "P34");
                    excelcells80.Value = "NEGATIV";
                }
            }


            if (richTextBox57.Text != String.Empty && comboBox16.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox57.Text) < Convert.ToDouble(comboBox16.Text))
                {
                    excelcells80 = sheet.get_Range("O35", "P36");
                    excelcells80.Value = "NEGATIV";
                }
            }


            if (richTextBox55.Text != String.Empty && comboBox17.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox55.Text) < Convert.ToDouble(comboBox17.Text))
                {
                    excelcells80 = sheet.get_Range("O37", "P38");
                    excelcells80.Value = "NEGATIV";
                }
            }


            if (richTextBox54.Text != String.Empty && comboBox18.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox54.Text) < Convert.ToDouble(comboBox18.Text))
                {
                    excelcells80 = sheet.get_Range("O39", "P40");
                    excelcells80.Value = "NEGATIV";
                }

            }


            if (richTextBox52.Text != String.Empty && comboBox19.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox52.Text) < Convert.ToDouble(comboBox19.Text))
                {
                    excelcells80 = sheet.get_Range("O41", "P42");
                    excelcells80.Value = "NEGATIV";
                }
            }

            if (richTextBox51.Text != String.Empty && comboBox20.Text != String.Empty)
            {

                if (Convert.ToDouble(richTextBox51.Text) < Convert.ToDouble(comboBox20.Text))
                {
                    excelcells80 = sheet.get_Range("O43", "P44");
                    excelcells80.Value = "NEGATIV";
                }
            }

            if (richTextBox49.Text != String.Empty && comboBox21.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox49.Text) < Convert.ToDouble(comboBox21.Text))
                {
                    excelcells80 = sheet.get_Range("O45", "P46");
                    excelcells80.Value = "NEGATIV";
                }
            }

            if (richTextBox37.Text != String.Empty && comboBox22.Text != String.Empty)
            {

                if (Convert.ToDouble(richTextBox37.Text) < Convert.ToDouble(comboBox22.Text))
                {
                    excelcells80 = sheet.get_Range("O47", "P48");
                    excelcells80.Value = "NEGATIV";
                }
            }


            if (richTextBox39.Text != String.Empty && comboBox23.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox39.Text) < Convert.ToDouble(comboBox23.Text))
                {
                    excelcells80 = sheet.get_Range("O49", "P50");
                    excelcells80.Value = "NEGATIV";
                }
            }

            if (richTextBox40.Text != String.Empty && comboBox24.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox40.Text) < Convert.ToDouble(comboBox24.Text))
                {
                    excelcells80 = sheet.get_Range("O51", "P52");
                    excelcells80.Value = "NEGATIV";
                }
            }


            if (richTextBox80.Text != String.Empty && comboBox25.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox42.Text) < Convert.ToDouble(comboBox25.Text))
                {
                    excelcells80 = sheet.get_Range("O53", "P54");
                    excelcells80.Value = "NEGATIV";
                }
            }


            if (richTextBox43.Text != String.Empty && comboBox27.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox43.Text) < Convert.ToDouble(comboBox27.Text))
                {
                    excelcells80 = sheet.get_Range("O55", "P56");
                    excelcells80.Value = "NEGATIV";
                }
            }


            if (richTextBox45.Text != String.Empty && comboBox29.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox45.Text) < Convert.ToDouble(comboBox29.Text))
                {
                    excelcells80 = sheet.get_Range("O57", "P58");
                    excelcells80.Value = "NEGATIV";
                }
            }


            if (richTextBox46.Text != String.Empty && comboBox31.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox46.Text) < Convert.ToDouble(comboBox31.Text))
                {
                    excelcells80 = sheet.get_Range("O59", "P60");
                    excelcells80.Value = "NEGATIV";
                }
            }


            if (richTextBox48.Text != String.Empty && comboBox33.Text != String.Empty)
            {
                if (Convert.ToDouble(richTextBox48.Text) < Convert.ToDouble(comboBox33.Text))
                {
                    excelcells80 = sheet.get_Range("O61", "P62");
                    excelcells80.Value = "NEGATIV";
                }
            }
            try
            {
                String str = DateTime.Now.ToString("yy.MM.dd HH.mm.ss");
                if (String.IsNullOrWhiteSpace(richTextBox5.Text)) { filename1 = "default" + str; }
                excelAppWorkbook.SaveAs(Filename: globalpath + @"\Imunologie\" + filename1 + ".xlsx");
            }

            catch { }
            excelAppWorkbook.Close();
            excelapp.Quit();
        }

        private void button5_Click(object sender, EventArgs e)//сохранить на ревмопробах
        {

            excelapp1 = new Excel.Application();//если не открыто, создаем новое
            excelAppWorkbooks = excelapp.Workbooks;
            excelAppWorkbook = excelAppWorkbooks.Open(globalpath + @"\ReumTemplate.xlsx");


            Excel.Workbook excelSheets = excelapp.ActiveWorkbook;
            sheet = (Excel.Worksheet)excelSheets.Worksheets[1];


            filename1 = richTextBox62.Text;
            excelcells = sheet.get_Range("I12", "J12"); //Рэндж ячеек data recoltarii
            excelcells.Value = richTextBox75.Text;



            excelcells1 = sheet.get_Range("N12", "O12"); //Рэндж ячеек data receptionarii
            excelcells1.Value = richTextBox73.Text;

            excelcells2 = sheet.get_Range("K15"); //Рэндж для имени и фамилии, а так же имени файла
            excelcells2.Value = richTextBox62.Text;

            excelcells3 = sheet.get_Range("O15"); //Рэндж возраста
            excelcells3.Value = richTextBox69.Text;

            excelcells4 = sheet.get_Range("K17"); //Рэндж номера идентификации
            excelcells4.Value = Convert.ToString(richTextBox72.Text);


            excelcells5 = sheet.get_Range("O17", "P17"); //Рендж номера страховки(полица де асигураре)
            excelcells5.Value = richTextBox67.Text;


            excelcells6 = sheet.get_Range("J19", "K19");// Рендж учреждения(институция)
            excelcells6.Value = richTextBox71.Text;

            excelcells7 = sheet.get_Range("O19", "P19");//рендж отделения(секция)
            excelcells7.Value = richTextBox66.Text;

            excelcells8 = sheet.get_Range("J21", "K21");//рендж участка
            excelcells8.Value = richTextBox70.Text;

            excelcells9 = sheet.get_Range("N21", "O21");//рендж номера медкарты
            excelcells9.Value = richTextBox64.Text;


            excelcells10 = sheet.get_Range("N9");//номер анализа
            excelcells10.Value = richTextBox78.Text;

            if (String.IsNullOrWhiteSpace(comboBox35.Text)) //1 = первый элемент и так далее
            {
                var testRange1 = sheet.Range[sheet.Cells[29, 8], sheet.Cells[29, 15]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[30, 8], sheet.Cells[30, 15]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells11 = sheet.get_Range("I29", "K30");
            excelcells11.Value = comboBox35.Text;



            if (String.IsNullOrWhiteSpace(comboBox36.Text))
            {
                var testRange1 = sheet.Range[sheet.Cells[31, 8], sheet.Cells[31, 15]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[32, 8], sheet.Cells[32, 15]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells12 = sheet.get_Range("I31", "K32");
            excelcells12.Value = comboBox36.Text;



            if (String.IsNullOrWhiteSpace(comboBox37.Text))
            {
                var testRange1 = sheet.Range[sheet.Cells[33, 8], sheet.Cells[33, 15]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[34, 8], sheet.Cells[34, 15]];
                testRange2.EntireRow.Hidden = true;
            }

            excelcells13 = sheet.get_Range("I33", "K34");
            excelcells13.Value = comboBox37.Text;

            if (String.IsNullOrWhiteSpace(comboBox38.Text))
            {
                var testRange1 = sheet.Range[sheet.Cells[35, 8], sheet.Cells[35, 15]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[36, 8], sheet.Cells[36, 15]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells14 = sheet.get_Range("I35", "K36");
            excelcells14.Value = comboBox38.Text;

            if (String.IsNullOrWhiteSpace(comboBox39.Text))
            {
                var testRange1 = sheet.Range[sheet.Cells[37, 8], sheet.Cells[37, 15]];
                testRange1.EntireRow.Hidden = true;
                var testRange2 = sheet.Range[sheet.Cells[38, 8], sheet.Cells[38, 15]];
                testRange2.EntireRow.Hidden = true;
            }
            excelcells15 = sheet.get_Range("I37", "K38");
            excelcells15.Value = comboBox39.Text;

            excelcells16 = sheet.get_Range("H43");//дата выдачи
            excelcells16.Value = richTextBox96.Text;



            excelcells29 = sheet.get_Range("K29", "L30"); //from here начинается заполнение колонки со значениями
            excelcells29.Value = richTextBox81.Text;


            excelcells30 = sheet.get_Range("K31", "L32");
            excelcells30.Value = richTextBox82.Text;


            excelcells31 = sheet.get_Range("K33", "L34");
            excelcells31.Value = richTextBox84.Text;



            excelcells32 = sheet.get_Range("K35", "L36");
            excelcells32.Value = richTextBox85.Text;


            excelcells33 = sheet.get_Range("K37", "L38");
            excelcells33.Value = richTextBox87.Text;
            // Величины

            //референс


            excelcells34 = sheet.get_Range("M29", "M30");
            excelcells34.Value = richTextBox88.Text;



            excelcells35 = sheet.get_Range("M31", "M32");
            excelcells35.Value = richTextBox90.Text;


            excelcells36 = sheet.get_Range("M33", "M34");
            excelcells36.Value = richTextBox91.Text;


            excelcells37 = sheet.get_Range("M35", "M36");
            excelcells37.Value = richTextBox93.Text;


            excelcells38 = sheet.get_Range("M37", "M38");
            excelcells38.Value = richTextBox94.Text;


            //интерпретация


            excelcells39 = sheet.get_Range("N29", "O30");
            excelcells39.Value = comboBox40.Text;


            excelcells40 = sheet.get_Range("N31", "O32");
            excelcells40.Value = comboBox41.Text;

            excelcells41 = sheet.get_Range("N33", "O34");
            excelcells41.Value = comboBox42.Text;

            excelcells42 = sheet.get_Range("N35", "O36");
            excelcells42.Value = comboBox43.Text;

            excelcells43 = sheet.get_Range("N37", "O38");
            excelcells43.Value = comboBox44.Text;

            try
            {
                String str = DateTime.Now.ToString("yy.MM.dd HH.mm.ss");
                if (String.IsNullOrWhiteSpace(richTextBox62.Text)) { filename1 = "default" + str; }
                excelAppWorkbook.SaveAs(Filename: globalpath + @"\Reumo.Probe\" + filename1 + ".xlsx");
            }

            catch { }
            excelAppWorkbook.Close();
            excelapp.Quit();











        }
    }
}
