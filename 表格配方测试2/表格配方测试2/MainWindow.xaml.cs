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

using System.Windows.Media.Animation;
using System.Windows.Threading;
using System.IO;
using System.Diagnostics;
using System.Security.Cryptography;
using unvell.ReoGrid.DataFormat;
namespace 表格配方测试2
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public class STEPItem
        {
            public string C1 { get; set; }
            public string C2 { get; set; }
            public string C3 { get; set; }
            public string C4 { get; set; }
            public string C5 { get; set; }
            public string C6 { get; set; }
            public string C7 { get; set; }
            public string C8 { get; set; }
            public string C9 { get; set; }
            public string C10 { get; set; }
            public string C11 { get; set; }
            public string C12 { get; set; }
        }
        public string PCindex;
        public int User = 0;
        private DispatcherTimer timer;
        
        public MainWindow()
        {
            InitializeComponent();
            ExcelInit(PC1reogrid);
            ExcelInit(PC2reogrid);
            DataFormatterManager.Instance.DataFormatters.Add(CellDataFormatFlag.Custom, new MyDataFormatter1());
            PC1reogrid.CurrentWorksheet.SetRangeDataFormat("D3:I22", CellDataFormatFlag.Custom);
            PC2reogrid.CurrentWorksheet.SetRangeDataFormat("D3:I22", CellDataFormatFlag.Custom);

            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(6);
            timer.Tick += seedListView;
            //timer.Start();
            //runstop.IsChecked = true;

            //获取欲启动进程名
            string strProcessName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
            //检查进程是否已经启动，已经启动则显示报错信息退出程序。
            if (System.Diagnostics.Process.GetProcessesByName(strProcessName).Length > 1)
            {
                MessageBox.Show("多个程序不能同时运行！", "系统错误");
                try
                {
                    System.Diagnostics.Process.GetCurrentProcess().Kill();
                }
                catch { }
                return;
            }

            if (!Directory.Exists("D:\\配方程序\\PC1")) { Directory.CreateDirectory("D:\\配方程序\\PC1"); }
            if (!Directory.Exists("D:\\配方程序\\PC2")) { Directory.CreateDirectory("D:\\配方程序\\PC2"); }
        }
        private void ExcelInit(unvell.ReoGrid.ReoGridControl Worksheet)
        {
            var worksheet = Worksheet.CurrentWorksheet;
            worksheet.Rows = 22;
            worksheet.Columns = 12;
            worksheet.DisableSettings(unvell.ReoGrid.WorksheetSettings.Edit_DragSelectionToMoveCells);
            worksheet.DisableSettings(unvell.ReoGrid.WorksheetSettings.View_ShowColumnHeader);
            worksheet.DisableSettings(unvell.ReoGrid.WorksheetSettings.View_ShowRowHeader);//工作表初始化 行 列

            worksheet.MergeRange("A1:B1");
            worksheet.MergeRange("C1:E1");
            worksheet.Ranges["A1:E1"].Style.TextColor = unvell.ReoGrid.Graphics.SolidColor.Blue;//合并单元格

            worksheet["A1"] = "配方名称";
            worksheet["A2:L2"] = new object[] { "NO", "Step Name", "Step Time(s)", "H2(slm)", "SiH4(slm)", "B2H6(slm)", "PH3(slm)", "CO2(slm)", "Gap(mm)", "Pressure(Pa)", "RF_Power(W)", "Recipe NO" };
            worksheet.Ranges["A2:L2"].IsReadonly = true;
            worksheet["A3:A22"] = new object[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 };
            worksheet.Ranges["A3:A22"].Style.TextColor = unvell.ReoGrid.Graphics.SolidColor.Black;
            worksheet.Ranges["A2:L2"].IsReadonly = true;

            worksheet.Ranges["A1:L22"].Style.HorizontalAlign = unvell.ReoGrid.ReoGridHorAlign.Center;
            worksheet.Ranges["A1:L1"].Style.BackColor = unvell.ReoGrid.Graphics.SolidColor.LightSteelBlue;
            worksheet.Ranges["A2:L2"].Style.BackColor = unvell.ReoGrid.Graphics.SolidColor.LightSkyBlue;
            worksheet.Ranges["A2:L2"].Style.TextColor = unvell.ReoGrid.Graphics.SolidColor.Black;
            worksheet.Ranges["D3:H22"].Style.TextColor = unvell.ReoGrid.Graphics.SolidColor.Black;
            worksheet.Ranges["B3:C22"].Style.BackColor = unvell.ReoGrid.Graphics.SolidColor.LightSteelBlue;
            worksheet.Ranges["D3:D22"].Style.BackColor = unvell.ReoGrid.Graphics.SolidColor.FromArgb(100);
            worksheet.Ranges["E3:E22"].Style.BackColor = unvell.ReoGrid.Graphics.SolidColor.FromArgb(10);
            worksheet.Ranges["F3:F22"].Style.BackColor = unvell.ReoGrid.Graphics.SolidColor.FromArgb(5);
            worksheet.Ranges["G3:G22"].Style.BackColor = unvell.ReoGrid.Graphics.SolidColor.FromArgb(5);
            worksheet.Ranges["H3:H22"].Style.BackColor = unvell.ReoGrid.Graphics.SolidColor.FromArgb(1);
            worksheet.Ranges["I3:L22"].Style.BackColor = unvell.ReoGrid.Graphics.SolidColor.SkyBlue;//颜色布局设置
            worksheet.Ranges["I3:I22"].Style.Underline = true;

            worksheet.SetRangeBorders("A1:L22", unvell.ReoGrid.BorderPositions.InsideAll, new unvell.ReoGrid.RangeBorderStyle { Color = unvell.ReoGrid.Graphics.SolidColor.White, Style = unvell.ReoGrid.BorderLineStyle.Dashed });
            worksheet.SetColumnsWidth(0, 12, 80);//列宽度设置
            worksheet.SetRangeDataFormat("J3:L22", CellDataFormatFlag.Number, new NumberDataFormatter.NumberFormatArgs()
            {
                // decimal digit places, e.g. 0.1234
                DecimalPlaces = 0,
                //  // negative number style, e.g. -123 -> (123) 
                //   NegativeStyle = NumberDataFormatter.NumberNegativeStyle.RedBrackets,
            });
            worksheet.SetRangeDataFormat("A1:C1", CellDataFormatFlag.Text);
            worksheet.SetRangeDataFormat("B3:B22", CellDataFormatFlag.Text);
            //, new NumberDataFormatter.NumberFormatArgs()
        }

        private void seedListView(object sender, EventArgs e)
        {
            //Define
            var data = new[]
            {
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12],
                new string[12]
            };
            //Read
            runingrepview.Items.Clear();
            Stopwatch ST = new Stopwatch();
            ST.Start();
            if (pc1.IsSelected)
            {
                PCindex = "PC1";
            }
            else
            {
                PCindex = "PC2";
            }
            int I;
            CCHMIRUNTIME.HMIRuntime HMIRT = new CCHMIRUNTIME.HMIRuntime();
            try
            {
                for (I = 1; I < 21; I++)
                {
                    data[I - 1][0] = I.ToString();
                    data[I - 1][1] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_StepName"].Read();
                    data[I - 1][2] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_StepTime"].Read().ToString();
                    data[I - 1][3] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_H2"].Read().ToString();
                    data[I - 1][4] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_SiH4"].Read().ToString();
                    data[I - 1][5] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_B2H6"].Read().ToString();
                    data[I - 1][6] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_PH3/H2"].Read().ToString();
                    data[I - 1][7] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_CO2"].Read().ToString();
                    data[I - 1][8] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_Gap"].Read().ToString();
                    data[I - 1][9] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_Pressure"].Read().ToString();
                    data[I - 1][10] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_RF_power"].Read().ToString();
                    data[I - 1][11] = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_Recipe_No"].Read().ToString();
                }
                runningrepname.Text = HMIRT.Tags[PCindex + "_STEP工艺参数_STEP配方名"].Read();
                readtime.IsChecked = false;
            }
            catch (Exception)
            {
                timer.Stop();
                runstop.IsChecked = false;
                MessageBox.Show("读取失败，请检查WINCC是否运行");
                return;
            }
            //Add
            
            foreach (string[] version in data)
            {
                runingrepview.Items.Add(new STEPItem { C1 = version[0] , C2 = version[1] , C3 = version[2], C4 = version[3], C5 = version[4], C6 = version[5], C7 = version[6], C8 = version[7], C9 = version[8], C10 = version[9], C11 = version[10], C12 = version[11] });
            }
            ST.Stop();
            readtime.Content = "读取时间：" + ST.Elapsed.TotalMilliseconds.ToString()+"ms";
            ST.Reset();
            DoubleAnimation animation = new DoubleAnimation(
                PGbar.Value = 0,                       // From
                PGbar.Value = 100,                  // To
                new Duration(TimeSpan.FromSeconds(6)))      // Duration    间隔是 10s
            {
                AccelerationRatio = 0,       // 设置加速
                DecelerationRatio = 0,       // 设置减速
            };
            PGbar.BeginAnimation(ProgressBar.ValueProperty,animation);
        }//读取正在运行配方

        private void writeAS(unvell.ReoGrid.ReoGridControl Worksheet)
        {
            int I;
            CCHMIRUNTIME.HMIRuntime HMIRT = new CCHMIRUNTIME.HMIRuntime();
            Stopwatch ST = new Stopwatch();
            ST.Start();
            if (pc1.IsSelected)
            {
                PCindex = "PC1";
            }
            else
            {
                PCindex = "PC2";
            }
            try
            {
                for (I = 1; I < 21; I++)
                {
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_StepName"].Write(Worksheet.CurrentWorksheet[I + 1, 1]);
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_StepTime"].Write(Worksheet.CurrentWorksheet[I + 1, 2]);
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_H2"].Write(Worksheet.CurrentWorksheet[I + 1, 3]);
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_SiH4"].Write(Worksheet.CurrentWorksheet[I + 1, 4]);
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_B2H6"].Write(Worksheet.CurrentWorksheet[I + 1, 5]);
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_PH3/H2"].Write(Worksheet.CurrentWorksheet[I + 1, 6]);
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_CO2"].Write(Worksheet.CurrentWorksheet[I + 1, 7]);
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_Gap"].Write(Worksheet.CurrentWorksheet[I + 1, 8]);
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_Pressure"].Write(Worksheet.CurrentWorksheet[I + 1, 9]);
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_RF_power"].Write(Worksheet.CurrentWorksheet[I + 1, 10]);
                    HMIRT.Tags[PCindex + "_STEP工艺参数_STEP设定[" + "" + I + "" + "]_Recipe_No"].Write(Worksheet.CurrentWorksheet[I + 1, 11]);
                }
                HMIRT.Tags[PCindex + "_STEP工艺参数_STEP配方名"].Write(Worksheet.CurrentWorksheet["C1"]);
            }
            catch (Exception)
            {
                MessageBox.Show("写入失败");
                return;
            }
            ST.Stop();
            writetime.Content = "写入时间：" + ST.Elapsed.TotalMilliseconds.ToString()+"ms";
            ST.Reset();
        }//写入配方

        class MyDataFormatter1 : IDataFormatter
        {
            public static Int32 ParseRGB(Color color)
            {
                return (Int32)(((uint)color.B << 16) | (ushort)(((ushort)color.G << 8) | color.R));
            }
            public string FormatCell(unvell.ReoGrid.Cell cell)
            {
                if (cell.Style.Underline)
                {
                    double gap = cell.GetData<double>();
                    if (gap == 0)
                    { return ""; }
                    else if (gap < 18)
                    { cell.Data = 18; return "≥18"; }
                    else if (gap > 40)
                    { cell.Data = 40; return "≤40"; }
                    else
                    { return((int)gap).ToString(); }
                }
                else
                {
                    int TEST = cell.Style.BackColor.ToArgb();//ParseRGB(cell.Style.BackColor) cell.Style.BackColor.B
                    double val = cell.GetData<double>();
                    if (val < 0)
                    { cell.Data = 0; return ">0"; }
                    else if (val > TEST)
                    { cell.Data = TEST; return "≤" + "" + TEST + ""; }
                    else return val.ToString("######.00");//string.Format("{0}", (-val).ToString("###,###,##0.00"))
                }
            }
            public bool PerformTestFormat()
            {
                return true;
            }
        }

        private void openfilebutton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            //过滤文件类型
            openFileDialog.Filter = "配方(*.rep)|*.rep";
            //单选
            openFileDialog.Multiselect = false;
            unvell.ReoGrid.ReoGridControl worksheet;
            if (pc1.IsSelected)
            {
                openFileDialog.InitialDirectory = "D:\\配方程序\\PC1";
                worksheet = PC1reogrid;
            }
            else
            {
                openFileDialog.InitialDirectory = "D:\\配方程序\\PC2";
                worksheet = PC2reogrid;
            }
            if (log.Password == "123456")
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    String filePath = openFileDialog.FileName;
                    using (FileStream fs = new FileStream(filePath, FileMode.Open))
                    {
                        StreamReader rd = new StreamReader(fs);//读取文件中的数据
                        try
                        {
                            for (int I = 0; I < 20; I++)  //读入数据并赋予数组
                            {
                                string line = rd.ReadLine();
                                string[] data = line.Split(' ');
                                for (int Y = 0; Y < 11; Y++)
                                {
                                    worksheet.CurrentWorksheet[I + 2, Y + 1] = AuthcodeHelper.Decode(data[Y]);
                                }
                            }
                            worksheet.CurrentWorksheet["C1"] = rd.ReadLine();
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("文件读取失败");
                            return;
                        }
                        rd.Close();
                        fs.Close();
                    }
                }
            }
        }

        public class AuthcodeHelper
        {
            const string KEY_64 = "HuidTeac";//注意了，是8个字符
            const string IV_64 = "HuidTeac";

            public static string Encode(string data)
            {
                byte[] byKey = System.Text.ASCIIEncoding.ASCII.GetBytes(KEY_64);
                byte[] byIV = System.Text.ASCIIEncoding.ASCII.GetBytes(IV_64);

                DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
                int i = cryptoProvider.KeySize;
                MemoryStream ms = new MemoryStream();
                CryptoStream cst = new CryptoStream(ms, cryptoProvider.CreateEncryptor(byKey, byIV), CryptoStreamMode.Write);

                StreamWriter sw = new StreamWriter(cst);
                sw.Write(data);
                sw.Flush();
                cst.FlushFinalBlock();
                sw.Flush();
                return Convert.ToBase64String(ms.GetBuffer(), 0, (int)ms.Length);

            }

            public static string Decode(string data)
            {
                byte[] byKey = System.Text.ASCIIEncoding.ASCII.GetBytes(KEY_64);
                byte[] byIV = System.Text.ASCIIEncoding.ASCII.GetBytes(IV_64);

                byte[] byEnc;
                try
                {
                    byEnc = Convert.FromBase64String(data);
                }
                catch
                {
                    return null;
                }

                DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
                MemoryStream ms = new MemoryStream(byEnc);
                CryptoStream cst = new CryptoStream(ms, cryptoProvider.CreateDecryptor(byKey, byIV), CryptoStreamMode.Read);
                StreamReader sr = new StreamReader(cst);
                return sr.ReadToEnd();
            }

        }//https://blog.csdn.net/yutiedun/article/details/105547508

        private void runstop_Click(object sender, RoutedEventArgs e)
        {
            if (runstop.IsChecked == false)
            {
                //PGbar.IsIndeterminate = false;
                timer.Stop();
                PGbar.Value = 0;
            }
            else
            {
                //PGbar.IsIndeterminate = true;
                timer.Start();
                DoubleAnimation animation = new DoubleAnimation(
                PGbar.Value = 0,                       // From
                PGbar.Value = 100,                  // To
                new Duration(TimeSpan.FromSeconds(6)))      // Duration    间隔是 10s
                {
                    AccelerationRatio = 0,       // 设置加速
                    DecelerationRatio = 0,       // 设置减速
                };
                PGbar.BeginAnimation(ProgressBar.ValueProperty, animation);
            }
                
        }

        private void LoadSTEP_Click(object sender, RoutedEventArgs e)
        {
            unvell.ReoGrid.ReoGridControl worksheet;
            if (pc1.IsSelected)
            {
                worksheet = PC1reogrid;
            }
            else
            {
                worksheet = PC2reogrid;
            }
            try
            {
                int I;
                for (I = 1; I < 21; I++)
                {
                    var version = runingrepview.Items[I - 1] as STEPItem;
                    worksheet.CurrentWorksheet[I + 1, 1] = version.C2;
                    worksheet.CurrentWorksheet[I + 1, 2] = version.C3;
                    worksheet.CurrentWorksheet[I + 1, 3] = version.C4;
                    worksheet.CurrentWorksheet[I + 1, 4] = version.C5;
                    worksheet.CurrentWorksheet[I + 1, 5] = version.C6;
                    worksheet.CurrentWorksheet[I + 1, 6] = version.C7;
                    worksheet.CurrentWorksheet[I + 1, 7] = version.C8;
                    worksheet.CurrentWorksheet[I + 1, 8] = version.C9;
                    worksheet.CurrentWorksheet[I + 1, 9] = version.C10;
                    worksheet.CurrentWorksheet[I + 1, 10] = version.C11;
                    worksheet.CurrentWorksheet[I + 1, 11] = version.C12;
                }
                worksheet.CurrentWorksheet["C1"] = runningrepname.Text;
            }
            catch (Exception)
            {
                MessageBox.Show("上传失败，请读取在线配方");
                return;
            }
        }

        private void WriteSTEP_Click(object sender, RoutedEventArgs e)
        {
            unvell.ReoGrid.ReoGridControl worksheet;
            if (pc1.IsSelected)
            {
                worksheet = PC1reogrid;
            }
            else
            {
                worksheet = PC2reogrid;
            }
            worksheet.CurrentWorksheet.EndEdit(unvell.ReoGrid.EndEditReason.NormalFinish);//先取消选中单元格
            writeAS(worksheet);
        }

        private void savefilebutton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "配方(*.rep)|*.rep";
            unvell.ReoGrid.ReoGridControl worksheet;
            if (pc1.IsSelected)
            {
                saveFileDialog.InitialDirectory = "D:\\配方程序\\PC1";
                worksheet = PC1reogrid;
            }
            else
            {
                saveFileDialog.InitialDirectory = "D:\\配方程序\\PC2";
                worksheet = PC2reogrid;
            }
            worksheet.CurrentWorksheet.EndEdit(unvell.ReoGrid.EndEditReason.NormalFinish);//先取消选中单元格
            if (worksheet.CurrentWorksheet["C1"] != null)
            {
                saveFileDialog.FileName = worksheet.CurrentWorksheet["C1"].ToString();

                if (saveFileDialog.ShowDialog() == true)
                {
                    String filePath = saveFileDialog.FileName;
                    try
                    {
                        FileStream fs = new FileStream(filePath, FileMode.Create);
                        StreamWriter sw = new StreamWriter(fs);
                        int I, Y;
                        for (I = 0; I < 20; I++)
                        {
                            for (Y = 0; Y < 11; Y++)
                            {
                                sw.Write(AuthcodeHelper.Encode(worksheet.CurrentWorksheet.GetCellText(I + 2, Y + 1)) + " ");
                            }
                            sw.WriteLine();
                        }
                        sw.Write(worksheet.CurrentWorksheet["C1"]);
                        //清空缓冲区
                        sw.Flush();
                        //关闭流
                        sw.Close();
                        fs.Close();
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("文件保存失败");
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("配方名不能为空");
            }
        }

        private void Load_Click(object sender, RoutedEventArgs e)
        {
            if (log.Password == "123456")
            {
                User = 1;
                LoadSTEP.IsEnabled = true;
                WriteSTEP.IsEnabled = true;
                openfilebutton.IsEnabled = true;
                savefilebutton.IsEnabled = true;
                runstop.IsEnabled = true;
                Load.Background = Brushes.LightGreen;
            }
        }
    }
}
