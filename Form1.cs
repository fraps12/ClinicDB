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
    public partial class Form1 : Form
    {

        MySqlConnection conn;
        String globalpath = @"C:\Users\" + Environment.UserName + @"\Documents\ClinicDB";
        /*public void DriveTreeInit()
            {
                string[] drivesArray = Directory.GetLogicalDrives();

                treeView1.BeginUpdate();
                treeView1.Nodes.Clear();

                foreach (string s in drivesArray)
                {
                    TreeNode drive = new TreeNode(s, 0, 0);
                    treeView1.Nodes.Add(drive);

                    GetDirs(drive);
                }


                treeView1.EndUpdate();
            }
            public void GetDirs(TreeNode node)
            {
                DirectoryInfo[] diArray;

                node.Nodes.Clear();

                string globalpath = node.FullPath;
                DirectoryInfo di = new DirectoryInfo(globalpath);

                try
                {
                    diArray = di.GetDirectories();
                }
                catch
                {
                    return;
                }

                foreach (DirectoryInfo dirinfo in diArray)
                {
                    TreeNode dir = new TreeNode(dirinfo.Name, 0, 1);
                    node.Nodes.Add(dir);
                }
            }*/
            public Form1()
        {
            InitializeComponent();
            PopulateTreeView();
            this.treeView1.NodeMouseClick +=
            new TreeNodeMouseClickEventHandler(this.treeView1_NodeMouseClick);
            //CheckDirectory();
            //DriveTreeInit();


        }


        
       /* private void CheckDirectory()
        {
            if (Directory.Exists(globalpath) && Directory.Exists(globalpath + @"\Biochimie") && Directory.Exists(globalpath + @"\Imunologie") && Directory.Exists(globalpath + @"\Reumo.Probe"))
            {         
                return;
            }
            if(!Directory.Exists(globalpath) && !Directory.Exists(globalpath + @"\Biochimie") && !Directory.Exists(globalpath + @"\Imunologie") && !Directory.Exists(globalpath + @"\Reumo.Probe"))
            MessageBox.Show("Отсутствуют необходимые директории. Создаем...");
            
            DirectoryInfo di = Directory.CreateDirectory(globalpath);            
                DirectoryInfo biohim = Directory.CreateDirectory(globalpath + @"\Biochimie");              
                    DirectoryInfo imunolog = Directory.CreateDirectory(globalpath + @"\Imunologie");                                                    
                         DirectoryInfo reumprob = Directory.CreateDirectory(globalpath + @"\Reumo.Probe");
            
        }
     */
   




        private async void Form1_Load(object sender, EventArgs e)
        {
            /*string connStr = "server = localhost; user = root; database = mydb; port = 3306; password = jordi;";

            conn = new MySqlConnection(connStr);

            await conn.OpenAsync();            
            string sql = "SELECT * FROM biochimica";
            MySqlCommand cmd = new MySqlCommand(sql, conn);
            MySqlDataReader mysqlReader = null;
            */


            /*try
            {
                /*mysqlReader = cmd.ExecuteReader();

                while (await mysqlReader.ReadAsync())
                {
                    listView1.Items.Add(Convert.ToString(mysqlReader["Denumirea_inst"]) + "      " + Convert.ToString(mysqlReader["Numar_de_indentificare"]) + "       " + Convert.ToString(mysqlReader["Id_polita_de_asigurare"]) + "        " + Convert.ToString(mysqlReader["Sectia"]));
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            finally
            {
                if (mysqlReader != null)
                    mysqlReader.Close();
            }*/
        }



        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (conn != null && conn.State != ConnectionState.Closed)
                conn.Close();
            
            Application.Exit();
            
        }

        private void выходToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (conn != null && conn.State != ConnectionState.Closed)
                conn.Close();
           
            Application.Exit();
        }

        //Explorer
         private void PopulateTreeView()
         {
             TreeNode rootNode;

             DirectoryInfo info = new DirectoryInfo(globalpath);
             if (info.Exists)
             {
                 rootNode = new TreeNode(info.Name);
                 rootNode.Tag = info;
                 GetDirectories(info.GetDirectories(), rootNode);
                 treeView1.Nodes.Add(rootNode);
             }
         }

         private void GetDirectories(DirectoryInfo[] subDirs,
    TreeNode nodeToAddTo)
         {
             TreeNode aNode;
             DirectoryInfo[] subSubDirs;
             foreach (DirectoryInfo subDir in subDirs)
             {
                 aNode = new TreeNode(subDir.Name, 0, 0);
                 aNode.Tag = subDir;
                 aNode.ImageKey = "folder";
                 subSubDirs = subDir.GetDirectories();
                 if (subSubDirs.Length != 0)
                 {
                     GetDirectories(subSubDirs, aNode);
                 }
                 nodeToAddTo.Nodes.Add(aNode);
             }

         }
         //Explorer code ends

         //Click on folder 
         void treeView1_NodeMouseClick(object sender,
     TreeNodeMouseClickEventArgs e)
         {
             TreeNode newSelected = e.Node;
             listView1.Items.Clear();
             DirectoryInfo nodeDirInfo = (DirectoryInfo)newSelected.Tag;
             ListViewItem.ListViewSubItem[] subItems;
             ListViewItem item = null;

             foreach (DirectoryInfo dir in nodeDirInfo.GetDirectories())
             {
                 item = new ListViewItem(dir.Name, 0);
                 subItems = new ListViewItem.ListViewSubItem[]
                           {new ListViewItem.ListViewSubItem(item, "Directory"),
                    new ListViewItem.ListViewSubItem(item,
                 dir.LastAccessTime.ToShortDateString())};
                 item.SubItems.AddRange(subItems);
                 listView1.Items.Add(item);
             }
             foreach (FileInfo file in nodeDirInfo.GetFiles())
             {
                 item = new ListViewItem(file.Name, 1);
                 subItems = new ListViewItem.ListViewSubItem[]
                           { new ListViewItem.ListViewSubItem(item, "file"),
                    new ListViewItem.ListViewSubItem(item,
                 file.LastAccessTime.ToShortDateString())};

                 item.SubItems.AddRange(subItems);
                 listView1.Items.Add(item);
             }

             listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

         }



       
       

       /* private void treeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            treeView1.BeginUpdate();

            foreach (TreeNode node in e.Node.Nodes)
            {
                GetDirs(node);
            }

            treeView1.EndUpdate();
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode selectedNode = e.Node;
            globalpath = selectedNode.FullPath;

            DirectoryInfo di = new DirectoryInfo(globalpath);
            FileInfo[] fiArray;
            DirectoryInfo[] diArray;
           

            try
            {
                fiArray = di.GetFiles();
                diArray = di.GetDirectories();
            }
            catch
            {
                return;
            }

            listView1.Items.Clear();

            foreach (DirectoryInfo dirInfo in diArray)
            {
                ListViewItem lvi = new ListViewItem(dirInfo.Name);
                lvi.SubItems.Add("0");
                lvi.SubItems.Add(dirInfo.LastWriteTime.ToString());
                lvi.ImageIndex = 0;

                listView1.Items.Add(lvi);
            }


            foreach (FileInfo fileInfo in fiArray)
            {
                ListViewItem lvi = new ListViewItem(fileInfo.Name);
                lvi.Tag = fileInfo.FullName;
                lvi.Tag = fileInfo.FullName;
                lvi.SubItems.Add(fileInfo.Length.ToString());
                lvi.SubItems.Add(fileInfo.LastWriteTime.ToString());

                string filenameExtension =
                  Path.GetExtension(fileInfo.Name).ToLower();
                listView1.Items.Add(lvi);
            }


            

            

            


            
            void GetDataFunc(string name, DialogResult dr)
            {
                if (dr == DialogResult.OK)
                {
                    try
                    {
                        // пробуем переименовать

                        System.IO.Directory.Move(globalpath + "\\" + listView1.SelectedItems[0].Text, globalpath + "\\" + name);
                        listView1.SelectedItems[0].Text = name;
                        //MessageBox.Show(fullPath + "\\" + listView1.SelectedItems[0].Text);
                        //MessageBox.Show(fullPath + "\\" + name);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            





        }*/
        //функция поиска в дереве. Если ничего не найдено - возвращает null
        private TreeNode SearchNode(string SearchText, TreeNode StartNode)
        {
            TreeNode node = null;
            while (StartNode != null)
            {
                if (StartNode.Text.ToLower().Contains(SearchText.ToLower()))
                {
                    node = StartNode; //что-то нашли, выходим
                    break;
                };
                if (StartNode.Nodes.Count != 0) //у узла есть дочерние элементы
                {
                    node = SearchNode(SearchText, StartNode.Nodes[0]);//ищем рекурсивно в дочерних
                    if (node != null)
                    {
                        break;//что-то нашли
                    };
                };
                StartNode = StartNode.NextNode;
            };
            return node;//вернули результат поиска
        }

        //нажатие на клавишу поиска
        private void textBox1_TextChanged(object sender, EventArgs e)
       
        {
            string SearchText = this.textBox1.Text;
            if (SearchText == "")
            {
                return;
            };
            TreeNode SelectedNode = SearchNode(SearchText, treeView1.Nodes[0]);//пытаемся найти в поле Text
            if (SelectedNode != null)
            {
                //нашли, выделяем...
                this.treeView1.SelectedNode = SelectedNode;
                this.treeView1.SelectedNode.Expand();
                this.treeView1.Select();
            };
        }


        private void файлToolStripMenuItem_Click(object sender, EventArgs e)
        {



        }

        private void файлToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            // DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.


            // string file = openFileDialog1.FileName;


            // <-- Shows file size in debugging mode.

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }





        private void biochimiaToolStripMenuItem_Click(object sender, EventArgs e) //добавление файла в папку с биохимией
        {
            try
            {
                string newLocation = " ";

                string folderLocation = globalpath + @"\Biochimie";/////////////


                var OFD = new System.Windows.Forms.OpenFileDialog();
                if (OFD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string fileToCopy = OFD.FileName;
                    if (System.IO.File.Exists(fileToCopy))
                    {
                        var onlyFileName = System.IO.Path.GetFileName(OFD.FileName);
                        newLocation = folderLocation + "\\" + onlyFileName;

                        if (File.Exists(newLocation) == true)
                        {
                            MessageBox.Show("File allready exist in this directory, please rename Source file and try again!");
                        }
                        else
                        {
                            System.IO.File.Copy(fileToCopy, newLocation, true);
                            MessageBox.Show("File Copied");
                        }
                    }

                }

            }
            catch (IOException)
            {

            }
        }

        private void imunologiaToolStripMenuItem_Click(object sender, EventArgs e) //добавление файла в папку с имунологией
        {
            try
            {

                string newLocation = " ";
                string folderLocation = globalpath + @"\Imunologie";
                var OFD = new System.Windows.Forms.OpenFileDialog();

                if (OFD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string fileToCopy = OFD.FileName;
                    if (System.IO.File.Exists(fileToCopy))
                    {
                        var onlyFileName = System.IO.Path.GetFileName(OFD.FileName);
                        newLocation = folderLocation + "\\" + onlyFileName;

                        if (File.Exists(newLocation) == true)
                        {
                            MessageBox.Show("File allready exist in this directory, please rename Source file and try again!");

                        }
                        else
                        {
                            System.IO.File.Copy(fileToCopy, newLocation, true);
                            MessageBox.Show("File Copied");
                        }
                    }

                }
                

            }
            catch (IOException)
            {

            }


        }

        private void reumProbeToolStripMenuItem_Click(object sender, EventArgs e)//добавление в папку с реумопробами
        {
            try
            {
                string newLocation = " ";
                string folderLocation = globalpath + @"\Reum.Probe";
                var OFD = new System.Windows.Forms.OpenFileDialog();
                if (OFD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string fileToCopy = OFD.FileName;
                    if (System.IO.File.Exists(fileToCopy))
                    {
                        var onlyFileName = System.IO.Path.GetFileName(OFD.FileName);
                        newLocation = folderLocation + "\\" + onlyFileName;

                        if (File.Exists(newLocation) == true)
                        {
                            MessageBox.Show("File allready exist in this directory, please rename Source file and try again!");
                        }
                        else
                        {
                            System.IO.File.Copy(fileToCopy, newLocation, true);
                            MessageBox.Show("File Copied");
                        }
                    }

                }

            }
            catch (IOException)
            {

            }
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        
        
       

       

        

        
        private void button1_MouseClick(object sender, MouseEventArgs e)//изменить
        {

            Form3 form3 = new Form3();
            form3.Show();
            
        }

        private void button2_MouseClick(object sender, MouseEventArgs e)//удалить
        {
            if(listView1.SelectedItems.Count > 0)
            {
                String FileToDelete = listView1.SelectedItems[0].Text;
                String path = globalpath;
                String dir0 = path + @"\Biochimie";
                String dir1 = path + @"\Imunologie";
                String dir2 = path + @"\Reumo.Probe";

                if (File.Exists(dir0 + "\\" + FileToDelete) == true)
                {
                    String fullPath = dir0 + "\\" + FileToDelete;
                    try
                    {
                        File.Delete(fullPath);
                        foreach (ListViewItem item in listView1.Items)
                            if (System.IO.File.Exists(item.ToString()))
                                listView1.Items.Remove(item);
                    }
                    catch (DirectoryNotFoundException dirNotFound)
                    {
                        MessageBox.Show(dirNotFound.Message);
                    }
                    catch (UnauthorizedAccessException)
                    {
                        MessageBox.Show("UnAuthorizedAccessException: Unable to access file. ");
                    }
                    
                    listView1.SelectedItems.Clear();

                }
                if (File.Exists(dir1 + "\\" + FileToDelete) == true)
                {
                    String fullPath = dir1 + "\\" + FileToDelete;
                    try
                    {
                        File.Delete(fullPath);
                        foreach (ListViewItem item in listView1.Items)
                            if (System.IO.File.Exists(item.ToString()))
                                listView1.Items.Remove(item);
                    }
                    catch (DirectoryNotFoundException dirNotFound)
                    {
                        MessageBox.Show(dirNotFound.Message);
                    }
                    catch (UnauthorizedAccessException)
                    {
                        MessageBox.Show("UnAuthorizedAccessException: Unable to access file. ");
                    }
                    
                    listView1.SelectedItems.Clear();
                }
                if (File.Exists(dir2 + "\\" + FileToDelete) == true)
                {

                    String fullPath = dir2 + "\\" + FileToDelete;
                    
                    try
                    {
                        File.Delete(fullPath);
                        foreach (ListViewItem item in listView1.Items)
                            if (System.IO.File.Exists(item.ToString()))
                                listView1.Items.Remove(item);
                    }
                    catch (DirectoryNotFoundException dirNotFound)
                    {
                        MessageBox.Show(dirNotFound.Message);
                    }
                    catch (UnauthorizedAccessException)
                    {
                        MessageBox.Show("UnAuthorizedAccessException: Unable to access file. ");
                    }
                    
                    listView1.SelectedItems.Clear();
                }

                if (listView1.SelectedIndices.Count <= 0)
                {
                    return;
                }
            }
        }

       

        private void button3_MouseClick(object sender, MouseEventArgs e)//создать
        {
            Form2 form2 = new Form2();
            form2.Show();
        }

        
        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                String text = listView1.SelectedItems[0].Text;
                String path = globalpath;
                String[] directory = Directory.GetDirectories(path);
                String dir0 = path + @"\Biochimie";
                String dir1 = path + @"\Imunologie";
                String dir2 = path + @"\Reumo.Probe";

                if (File.Exists(dir0 + "\\" + text) == true)
                {
                    String fullPath = dir0 + "\\" + text;
                    Process.Start(fullPath);
                    listView1.SelectedItems.Clear();

                }
                if (File.Exists(dir1 + "\\" + text) == true)
                {
                    String fullPath = dir1 + "\\" + text;
                    Process.Start(fullPath);
                    listView1.SelectedItems.Clear();
                }
                if (File.Exists(dir2 + "\\" + text) == true)
                {

                    String fullPath = dir2 + "\\" + text;
                    Process.Start(fullPath);
                    listView1.SelectedItems.Clear();
                }

                if (listView1.SelectedIndices.Count <= 0)
                {
                    return;
                }
            }
        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {

                SelectedItemPath.SelItemPath = listView1.SelectedItems[0].Text;
                String path = globalpath;
                String[] directory = Directory.GetDirectories(path);
                String dir0 = path + @"\Biochimie";
                String dir1 = path + @"\Imunologie";
                String dir2 = path + @"\Reumo.Probe";

                try
                {
                    if (File.Exists(dir0 + "\\" + SelectedItemPath.SelItemPath) == true)
                    {
                        SelectedItemPath.FilePathExists = dir0 + "\\" + SelectedItemPath.SelItemPath;
                        //listView1.SelectedItems.Clear();

                    }
                    if (File.Exists(dir1 + "\\" + SelectedItemPath.SelItemPath) == true)
                    {
                        SelectedItemPath.FilePathExists = dir1 + "\\" + SelectedItemPath.SelItemPath;
                        //listView1.SelectedItems.Clear();
                    }
                    if (File.Exists(dir2 + "\\" + SelectedItemPath.SelItemPath) == true)
                    {

                        SelectedItemPath.FilePathExists = dir2 + "\\" + SelectedItemPath.SelItemPath;
                        //listView1.SelectedItems.Clear();
                    }
                }
                catch
                {
                    if (File.Exists(dir0 + "\\" + SelectedItemPath.SelItemPath) == false)
                    {
                        SelectedItemPath.FilePathExists = dir0 + "\\" + "default.xlsx";
                        //listView1.SelectedItems.Clear();

                    }
                    if (File.Exists(dir1 + "\\" + SelectedItemPath.SelItemPath) == false)
                    {
                        SelectedItemPath.FilePathExists = dir1 + "\\" + "default.xlsx";
                        //listView1.SelectedItems.Clear();
                    }
                    if (File.Exists(dir2 + "\\" + SelectedItemPath.SelItemPath) == false)
                    {

                        SelectedItemPath.FilePathExists = dir2 + "\\" + "default.xlsx";
                        //listView1.SelectedItems.Clear();
                    }
                }
                if (listView1.SelectedIndices.Count <= 0)
                {
                    return;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            KillProc.ExcelKillProcess();
        }
    }
}