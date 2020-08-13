using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ReadAndWrite;
using System.Data;
namespace combinexls
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //直接处理
            mm();
        }


        private void mm()
        {
            DataTable dt = new DataTable();
            string fil = "excel(*.xlsx)|*.xlsx|excel(*.xls)|*.xls";
            List<string> files = XlsRAW.Readloads(fil);
            if(files.Count>0)
            {
                try
                {
                    foreach(string fl in files)
                    {
                        DataTable tb = XlsRAW.readxls(fl);
                        //判断是否已创建列
                        if(dt.Columns.Count==0)
                        {
                            dt = tb.Copy();
                            //for(int i=0;i<tb.Columns.Count;i++)
                            //{
                            //    dt.Columns.Add(tb.Columns[i].ColumnName);
                            //}                    
                        }
                        else
                        {
                            //添加新行
                            for(int i=0;i<tb.Rows.Count;i++)
                            {                        
                                //dt.Rows.Add(tb.Rows[i]);                    
                                dt.ImportRow(tb.Rows[i]);
                            }
                        }

                    }
                    string dir=files[0];
                    string writename=dir.Replace(dir.Substring(dir.LastIndexOf('\\')+1),"合并结果.xlsx");
                    XlsRAW.Writexls(writename, dt);
                    this.Close();
                }
                catch
                {
                    MessageBox.Show("结构不一样");
                }
            }
            else
            {
                MessageBox.Show("sth wrong");
            }            

        }
    }
}
