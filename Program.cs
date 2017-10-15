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
namespace CLdb
{
    static class SelectedItemPath
    {
        public static String SelItemPath {
            set; get;
        }
        public static String FilePathExists
        {
            set; get;
        }
        
    }
    static class KillProc
    {
        public static void ExcelKillProcess()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
        }
    }
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
           Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

           
            

                String globalpath = @"C:\Users\" + Environment.UserName + @"\Documents\ClinicDB";


            if (Directory.Exists(globalpath) && Directory.Exists(globalpath + @"\Biochimie") && Directory.Exists(globalpath + @"\Imunologie") && Directory.Exists(globalpath + @"\Reumo.Probe"))
            {
                Application.Run(new Form1());
            }
                    //MessageBox.Show("Отсутствуют необходимые директории. Создаем...");          
            Application.Run(new CheckDirectoriesForm());
            
        }
    }

}

