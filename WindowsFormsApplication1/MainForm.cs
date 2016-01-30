using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        /*
注意：使用之前需要引用COM：Microsoft Office 11.0 Object Library 如果引用列表中没有，需要自行添加 C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE

调用方法：

MengXianhui.Utility.ExcelReport.InsertPictureToExcel ipt = new MengXianhui.Utility.ExcelReport.InsertPictureToExcel();   
ipt.Open();   
ipt.InsertPicture("B2", @"C:\Excellogo.gif");   
ipt.InsertPicture("B8", @"C:\Excellogo.gif",120,80);   
ipt.SaveFile(@"C:\ExcelTest.xls");   
ipt.Dispose();   
*//// <summary>   
  /// 功能：实现Excel应用程序的打开   
  /// </summary>   
  /// <param name="TemplateFilePath">模板文件物理路径</param>   
        public void Open(string TemplateFilePath)
        {
            //打开对象   
            // cadApp = new Autodesk.AutoCAD.Interop.AcadApplication();

            if (m_objExcel == null)
            {
                m_objExcel = new Excel.Application();
            }
            
            if (this.chkShowTab.Checked)
            {
                m_objExcel.Visible = true;
                //cadApp.Visible = true;
            }
            else
            {
                m_objExcel.Visible = false;
                //cadApp.Visible = false;
            }
            
            m_objExcel.DisplayAlerts = false;
            //if (m_objExcel.Version != "11.0")
            //{
            //    MessageBox.Show("您的 Excel 版本不是 11.0 （Office 2003），操作可能会出现问题。");
            //    m_objExcel.Quit();
            //    return;
            //}
            m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
            if (TemplateFilePath.Equals(String.Empty))
            {
                m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));
            }
            else
            {
                m_objBook = m_objBooks.Open(TemplateFilePath, m_objOpt, m_objOpt, m_objOpt, m_objOpt,
                  m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
            }
            m_objSheets = (Excel.Sheets)m_objBook.Worksheets;
            //m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item(1));
            int k = 0;
            bool bOKA = false;
           
            for (k = 1; k <= m_objSheets.Count; k++)
            {
                string strName = "退耕地小班外业调查表";
                //string strNameB = "小班农户统计表";


                if (m_objSheets.Item[k].Name == strName)
                {
                    m_objSheet = (Excel._Worksheet)(m_objSheets.Item[k]);//打开特定的工作表
                    m_objSheet.Select();
                    m_objExcel.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(m_objExcel_WorkbookBeforeClose);

                    bOKA = true;
                }
                //if (m_objSheets.Item[k].Name == strNameB)
                //{
                //    m_objSheetB = (Excel._Worksheet)(m_objSheets.Item[k]);//打开特定的工作表
                //    m_objExcel.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(m_objExcel_WorkbookBeforeClose);
                    
                //}
            }

            if (!bOKA)
                MessageBox.Show("没找到 退耕地小班外业调查表！");
            //if (!bOKB)
            //    MessageBox.Show("没找到 小班农户统计表！");


        }
        private void m_objExcel_WorkbookBeforeClose(Excel.Workbook m_objBooks, ref bool _Cancel)
        {
           // MessageBox.Show("保存完毕！打开任务管理器清理CAD");
            
        }
        /// <summary>   
        /// 将图片插入到指定的单元格位置。   
        /// 注意：图片必须是绝对物理路径   
        /// </summary>   
        /// <param name="RangeName">单元格名称，例如：B4</param>   
        /// <param name="PicturePath">要插入图片的绝对路径。</param>   
        public void InsertPicture(string RangeName, string PicturePath)
        {
            m_objRange = m_objSheet.get_Range(RangeName, m_objOpt);
            m_objRange.Select();
            Excel.Pictures pics = (Excel.Pictures)m_objSheet.Pictures(m_objOpt);
            pics.Insert(PicturePath, m_objOpt);
        }
        //Excel.Range rgA = null;
        public void InsertData(string strRa,string strRb,string strData)
        {
            Excel.Range rgA = null;
            m_objSheet.Select();

            if (strRa == strRb)
            {
                rgA = m_objSheet.get_Range(strRa, m_objOpt);
            }
            else
            {
                rgA = m_objSheet.get_Range(strRa, strRb);
            }
           if(rgA.MergeArea!=null)
            {

                //清楚原有数据

                rgA.UnMerge();
            }

            m_objSheet.Select();
            if (strRa == strRb)
            {
                rgA = m_objSheet.get_Range(strRa, m_objOpt);
            }
            else
            {
                rgA = m_objSheet.get_Range(strRa, strRb);
            }
            try
            {
                rgA.Select();
                rgA.Merge(0);
                rgA.Select();
                rgA.Value2 = strData;
                rgA.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                rgA.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("添加常规数据出错"+ex.Message.ToString());
            }
            
            
            
            // m_objRange = null;
        }
        /// <summary>   
        /// 将图片插入到指定的单元格位置，并设置图片的宽度和高度。   
        /// 注意：图片必须是绝对物理路径   
        /// </summary>   
        /// <param name="RangeName">单元格名称，例如：B4</param>   
        /// <param name="PicturePath">要插入图片的绝对路径。</param>   
        /// <param name="PictuteWidth">插入后，图片在Excel中显示的宽度。</param>   
        /// <param name="PictureHeight">插入后，图片在Excel中显示的高度。</param>   
        public void InsertPicture(string RangeName,string rgNameB, string PicturePath, double PictuteWidth, double  PictureHeight)
        {
            m_objSheet.Select();
            if (RangeName == rgNameB)
            {
                m_objRange = m_objSheet.get_Range(RangeName, m_objOpt);
            }
            else
            {
                m_objRange = m_objSheet.get_Range(RangeName, rgNameB);
            }
            try
            {
                m_objRange.UnMerge();
            }
            catch (Exception ex)
            {
                MessageBox.Show("insert pic"+ex.Message.ToString());
            }
           

            m_objRange.Select();
            m_objRange.Merge();
            //计算图片高度，宽度
            double desHeight=0, desWidth=0;
            if (RangeName == rgNameB)
            {
                m_objRange = m_objSheet.get_Range(RangeName, m_objOpt);
            }
            else
            {
                m_objRange = m_objSheet.get_Range(RangeName, rgNameB);
            }


            if (m_objRange.MergeArea != null)
            {
                desWidth = m_objRange.MergeArea.Width;
                desHeight = m_objRange.MergeArea.Height;
            }
            else
            {
                desWidth = desHeight = m_objRange.RowHeight * 4-2;
            }
           
            m_objRange.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            m_objRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            this.txtHeight.Text = desHeight.ToString();
            this.txtWidth.Text = desWidth.ToString();
            float PicLeft, PicTop;
            if (m_objRange.MergeArea != null)
            {
                PicLeft = (float)(m_objRange.Left + (m_objRange.MergeArea.Width - desHeight) / 2);
            }
            else
            {
                PicLeft = (float)m_objRange.Left+(float)m_objRange.Width / 2-(float)desHeight/2;
            }

            
            PicTop = (float)m_objRange.Top+(float)m_objRange.Height/2-(float)desHeight/2;
            
           
            //参数含义：   
            //图片路径   
            //是否链接到文件    m_objRange.RowHeight
            //图片插入时是否随文档一起保存   
            //图片在文档中的坐标位置（单位：points）   
            //图片显示的宽度和高度（单位：points）   
            //参数详细信息参见：http://msdn2.microsoft.com/zh-cn/library/aa221765(office.11).aspx   
            m_objSheet.Shapes.AddPicture(PicturePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, PicLeft, PicTop, (float)desHeight, (float)desHeight);
            
        }
        //Autodesk.AutoCAD.Interop.AcadApplication cadApp = null;
        //Autodesk.AutoCAD.Interop.AcadDocument cadDoc = null;
        /// <summary>   
        /// 将图片插入到指定的单元格位置，并设置图片的宽度和高度。   
        /// 注意：图片必须是绝对物理路径   
        /// </summary>   
        /// <param name="RangeName">单元格名称，例如：B4</param>   
        /// <param name="PicturePath">要插入图片的绝对路径。</param>   
        /// <param name="PictuteWidth">插入后，图片在Excel中显示的宽度。</param>   
        /// <param name="PictureHeight">插入后，图片在Excel中显示的高度。</param>   
        public void InsertOleCAD(string RangeName, string rgNameB, string PicturePath, double PictuteWidth, double PictureHeight)
        {
            m_objSheet.Select();
            if (RangeName == rgNameB)
            {
                m_objRange = m_objSheet.get_Range(RangeName, m_objOpt);
            }
            else
            {
                m_objRange = m_objSheet.get_Range(RangeName, rgNameB);
            }
            try
            {
                m_objRange.UnMerge();
            }
            catch (Exception ex)
            {
                MessageBox.Show("拆分单元格出错"+ex.Message.ToString());
            }


            m_objRange.Select();
            m_objRange.Merge();
            //计算图片高度，宽度
            double desHeight = 0, desWidth = 0;
            if (RangeName == rgNameB)
            {
                m_objRange = m_objSheet.get_Range(RangeName, m_objOpt);
            }
            else
            {
                m_objRange = m_objSheet.get_Range(RangeName, rgNameB);
            }
            m_objRange.Select();
           

            if (m_objRange.MergeArea != null)
            {
                desWidth = m_objRange.MergeArea.Width;
                desHeight = m_objRange.MergeArea.Height;
            }
            else
            {
                desWidth = desHeight = m_objRange.RowHeight * 4 ;
            }

            m_objRange.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            m_objRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            this.txtHeight.Text = desHeight.ToString();
            this.txtWidth.Text = desWidth.ToString();
            float PicLeft, PicTop;

            // PicLeft = (float)m_objRange.Left+ (float)(m_objRange.Width- m_objRange.Height)/2;
            PicLeft = (float)m_objRange.Left;

            PicTop = (float)m_objRange.Top;


            //参数含义：   
            //图片路径   
            //是否链接到文件    m_objRange.RowHeight
            //图片插入时是否随文档一起保存   
            //图片在文档中的坐标位置（单位：points）   
            //图片显示的宽度和高度（单位：points）   
            //参数详细信息参见：http://msdn2.microsoft.com/zh-cn/library/aa221765(office.11).aspx   

            //ActiveSheet.OLEObjects.Add(Filename:="D:\CADS\1_张贵财_12_40.619.dwg", Link:= _False, DisplayAsIcon:= False).Select
            //m_objSheet.PasteSpecial(Excel.XlPasteType.xlPasteComments, PicturePath, false, m_objOpt, m_objOpt, m_objOpt, m_objOpt);

            try
            {
                //cadDoc = cadApp.Documents.Open(PicturePath);
                //cadDoc.Activate();
                Excel.Shape sp = m_objSheet.Shapes.AddOLEObject("AutoCAD.Application.16", PicturePath, false,false, m_objOpt, m_objOpt, m_objOpt, PicLeft, PicTop, (float)m_objRange.Width, (float)m_objRange.Height);
                sp.Select();
                sp.Width = (float)m_objRange.Width;
                sp.Height =(float)m_objRange.Height;
                sp.ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStyleNotAPreset;
                
                //m_objSheet.OLEObjects.AddOwnedForm()
                // sp.VerticalFlip
                sp.Apply();
                
                sp = null;




                // cadDoc.Close(false);


                //Process pr = Process.GetProcessById(sp.Creator);
                //pr.Close();
                SaveFile(this.txtSaveAs.Text.ToString());
                m_objBook = null;
               // m_objExcel.Quit();
                //m_objExcel = null;
                Open(this.txtSaveAs.Text.ToString());
            }
            catch(Exception ex)
            {

                SaveFile(this.txtSaveAs.Text.ToString());
                // m_objBook.Application.Quit();

                m_objExcel.Quit();
                m_objExcel = null;
                GC.Collect();

                // m_objExcel = null;

                //保存并关闭表格
                Open(this.txtSaveAs.Text.ToString());
                //重新打开表格
                MessageBox.Show("插入CDA出错"+ex.Message.ToString());


               // this.listBox2.Items.Add(this.listBox1.Items[this.listBox1.SelectedIndex].ToString());
                
            }


   
            // m_objSheet.Shapes.AddOLEObject("AutoCAD.Application.16", PicturePath, m_objOpt, false, m_objOpt, m_objOpt, m_objOpt, PicLeft, PicTop,10, 10);
           


        }
        public void CopyData(string strRgA,string strRgB,string strRgC,string strRgD)
        {
            m_objSheet.Select();
            Excel.Range rg1 = m_objSheet.get_Range(strRgA,strRgB);
            m_objSheetB.Select();
            Excel.Range rgtemp = null;
            Excel.Range rg2 = m_objSheetB.get_Range(strRgC, strRgD);
            rg2.Select();
            rgtemp.Value2 = rg2.Value2;

            m_objSheet.Select();
            rg1.Select();
            rg1.Value2 = rgtemp.Value2;
            //打开excel 打开工作簿 

        }
        

        /// <summary>   
        /// 将Excel文件保存到指定的目录，目录必须事先存在，文件名称不一定要存在。   
        /// </summary>   
        /// <param name="OutputFilePath">要保存成的文件的全路径。</param>   
        public void SaveFile(string OutputFilePath)
        {
            try
            {
              


                m_objBook.SaveAs(OutputFilePath, m_objOpt, m_objOpt,
                 m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlNoChange,
                 m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                m_objBook.Close();
                m_objBook = null;

                

            }
            catch
            {
                MessageBox.Show("保存失败，可能没有关闭表");
            }
             
            
        }
        private Excel.Application m_objExcel = null;
        private Excel.Workbooks m_objBooks = null;
        private Excel._Workbook m_objBook = null;
        private Excel.Sheets m_objSheets = null;
        private Excel._Worksheet m_objSheet = null;
        private Excel._Worksheet m_objSheetB = null;
        private Excel.Range m_objRange = null;
        private object m_objOpt = System.Reflection.Missing.Value;
        private void button1_Click(object sender, EventArgs e)
        {
            int fileCout = 0;
            DirectoryInfo folder = new DirectoryInfo(this.txtPicPath.Text.ToString());
            Open(this.txtXLSPath.Text.ToString());
            //另存表格
           


            if (this.chkTuban.Checked == true)//序号和图斑号一致，不拆分文件名
            {

            }
            else//序号和图斑号不一致，读取序号
            {
                
                this.listBox1.Items.Clear();
                this.labMsg.Text = "";
                
                this.progressBar1.Maximum = Searchfile(this.txtPicPath.Text.ToString());
                this.progressBar1.Minimum = 0;
                
                float progress = 0;
                if (this.chkIsPic.Checked)
                {
                    foreach (FileInfo file in folder.GetFiles("*.jpg"))
                    {

                        //try
                        //{
                        string xuhao = "";
                        string qs = "";
                        double mj = 0;
                        string strMj = "";

                        xuhao = file.Name.ToString().Split('_')[0];
                        qs = file.Name.ToString().Split('_')[1];
                        strMj = file.Name.ToString().Split('_')[3];
                        mj = Math.Round(double.Parse(strMj) / 666.67, 2);
                        strMj = mj.ToString();
                        int toNum = 0;
                        toNum = (Int32.Parse(xuhao) - 1) * 4 + 13;
                        int toNumB = 0;
                        toNumB = toNum + 3;
                        this.listBox1.Items.Add(file.Name);
                        this.listBox1.SelectedIndex = this.listBox1.Items.Count - 1;
                        //InsertPicture("U" + rowNum.ToString(), file.FullName.ToString());

                        InsertData("A" + toNum.ToString(), "A" + toNumB.ToString(), this.txtCun.Text.ToString());
                        InsertData("B" + toNum.ToString(), "B" + toNumB.ToString(), this.txtCunzuming.Text.ToString());
                        InsertData("C" + toNum.ToString(), "C" + toNumB.ToString(), "");

                        InsertData("D" + toNum.ToString(), "D" + toNumB.ToString(), xuhao);

                        InsertData("E" + toNum.ToString(), "E" + toNumB.ToString(), qs);

                        InsertData("F" + toNum.ToString(), "F" + toNumB.ToString(), "");
                        InsertData("G" + toNum.ToString(), "G" + toNumB.ToString(), "");
                        InsertData("H" + toNum.ToString(), "H" + toNumB.ToString(), "");
                        InsertData("I" + toNum.ToString(), "I" + toNumB.ToString(), "");
                        InsertData("J" + toNum.ToString(), "J" + toNumB.ToString(), "");
                        InsertData("K" + toNum.ToString(), "K" + toNumB.ToString(), "");
                        InsertData("L" + toNum.ToString(), "L" + toNumB.ToString(), "");
                        InsertData("M" + toNum.ToString(), "M" + toNumB.ToString(), "");
                        InsertData("N" + toNum.ToString(), "N" + toNumB.ToString(), "");
                        InsertData("O" + toNum.ToString(), "O" + toNumB.ToString(), "");
                        InsertData("P" + toNum.ToString(), "P" + toNumB.ToString(), "");
                        InsertData("Q" + toNum.ToString(), "Q" + toNumB.ToString(), "");



                        InsertData("R" + toNum.ToString(), "R" + toNumB.ToString(), strMj);
                        InsertData("S" + toNum.ToString(), "S" + toNumB.ToString(), "");
                        InsertData("T" + toNum.ToString(), "T" + toNumB.ToString(), "");
                        InsertPicture("U" + toNum.ToString(), "W" + toNumB.ToString(), file.FullName.ToString(), double.Parse(this.txtWidth.Text.ToString()), double.Parse(this.txtHeight.Text.ToString()));

                        InsertData("X" + toNum.ToString(), "X" + toNumB.ToString(), "");

                        InsertData("Y" + toNum.ToString(), "Y" + toNum.ToString(), "东");
                        InsertData("Y" + (toNum + 1).ToString(), "Y" + (toNum + 1).ToString(), "南");
                        InsertData("Y" + (toNum + 2).ToString(), "Y" + (toNum + 2).ToString(), "西");
                        InsertData("Y" + (toNum + 3).ToString(), "Y" + (toNum + 3).ToString(), "北");
                        //rowNum += 4;
                        fileCout++;
                        //Console.WriteLine(file.FullName);
                        //}
                        //catch(Exception ex)
                        //{
                        //    this.labMsg.Text = ex.Message.ToString();
                        //}
                        this.progressBar1.Increment(1);

                    }
                }
                else
                {
                    foreach (FileInfo file in folder.GetFiles("*.dwg"))
                    {
                        progress++;
                        this.labMsg.Text = progress.ToString();
                        if ((int)(progress / 1000) == progress / 1000)
                        {
                            SaveFile(this.txtSaveAs.Text.ToString());
                            // m_objBook.Application.Quit();

                            m_objExcel.Quit();

                            m_objExcel = null;

                            //保存并关闭表格
                            Open(this.txtSaveAs.Text.ToString());
                            //重新打开表格
                            MessageBox.Show("手动关闭CAD");

                        }
                        
                            //try
                            //{
                            string xuhao = "";
                            string qs = "";
                            double mj = 0;
                            string strMj = "";

                            xuhao = file.Name.ToString().Split('_')[0];
                            qs = file.Name.ToString().Split('_')[1];
                            strMj = file.Name.ToString().Split('_')[3];
                            if (strMj.IndexOf("dwg") >= 1)
                            { mj = Math.Round(double.Parse(strMj.Substring(0,strMj.IndexOf(".dwg")+1)) / 666.67, 2); }
                            else
                            { mj = Math.Round(double.Parse(strMj) / 666.67, 2); }

                            strMj = mj.ToString();
                            int toNum = 0;
                            toNum = (Int32.Parse(xuhao) - 1) * 4 + 13;
                            int toNumB = 0;
                            toNumB = toNum + 3;
                            this.listBox1.Items.Add(file.Name);
                            this.listBox1.SelectedIndex = this.listBox1.Items.Count - 1;
                            //InsertPicture("U" + rowNum.ToString(), file.FullName.ToString());

                            InsertData("A" + toNum.ToString(), "A" + toNumB.ToString(), this.txtCun.Text.ToString());
                            InsertData("B" + toNum.ToString(), "B" + toNumB.ToString(), this.txtCunzuming.Text.ToString());
                            InsertData("C" + toNum.ToString(), "C" + toNumB.ToString(), "");

                            InsertData("D" + toNum.ToString(), "D" + toNumB.ToString(), xuhao);

                            InsertData("E" + toNum.ToString(), "E" + toNumB.ToString(), qs);

                            InsertData("F" + toNum.ToString(), "F" + toNumB.ToString(), "");
                            InsertData("G" + toNum.ToString(), "G" + toNumB.ToString(), "");
                            InsertData("H" + toNum.ToString(), "H" + toNumB.ToString(), "");
                            InsertData("I" + toNum.ToString(), "I" + toNumB.ToString(), "");
                            InsertData("J" + toNum.ToString(), "J" + toNumB.ToString(), "");
                            InsertData("K" + toNum.ToString(), "K" + toNumB.ToString(), "");
                            InsertData("L" + toNum.ToString(), "L" + toNumB.ToString(), "");
                            InsertData("M" + toNum.ToString(), "M" + toNumB.ToString(), "");
                            InsertData("N" + toNum.ToString(), "N" + toNumB.ToString(), "");
                            InsertData("O" + toNum.ToString(), "O" + toNumB.ToString(), "");
                            InsertData("P" + toNum.ToString(), "P" + toNumB.ToString(), "");
                            InsertData("Q" + toNum.ToString(), "Q" + toNumB.ToString(), "");



                            InsertData("R" + toNum.ToString(), "R" + toNumB.ToString(), strMj);
                            InsertData("S" + toNum.ToString(), "S" + toNumB.ToString(), "");
                            InsertData("T" + toNum.ToString(), "T" + toNumB.ToString(), "");
                            InsertOleCAD("U" + toNum.ToString(), "W" + toNumB.ToString(), file.FullName.ToString(), double.Parse(this.txtWidth.Text.ToString()), double.Parse(this.txtHeight.Text.ToString()));

                            InsertData("X" + toNum.ToString(), "X" + toNumB.ToString(), "");

                            InsertData("Y" + toNum.ToString(), "Y" + toNum.ToString(), "东");
                            InsertData("Y" + (toNum + 1).ToString(), "Y" + (toNum + 1).ToString(), "南");
                            InsertData("Y" + (toNum + 2).ToString(), "Y" + (toNum + 2).ToString(), "西");
                            InsertData("Y" + (toNum + 3).ToString(), "Y" + (toNum + 3).ToString(), "北");
                            //rowNum += 4;
                            fileCout++;
                            //Console.WriteLine(file.FullName);
                            //}
                            //catch(Exception ex)
                            //{
                            //    this.labMsg.Text = ex.Message.ToString();
                            //}
                            

                        

                        this.progressBar1.Increment(1);

                    }
                }
                
                if (fileCout > 0)
                {
                    this.labMsg.Text = "共添加图片" + fileCout.ToString();
                }
                else
                {
                    this.labMsg.Text = "处理失败！";
                }
                
                SaveFile(this.txtSaveAs.Text);
            }
           
        }
        public int Searchfile(string Directory)
        {
            int countfile = 0;//文件数量
            int countdir = 0;//文件夹数量
            DirectoryInfo dir = new DirectoryInfo(Directory);
            FileSystemInfo[] fi = dir.GetFileSystemInfos();//获取文件夹下的文件

            foreach (FileSystemInfo f in fi)
            {
                if (f is DirectoryInfo) //判断是否为文件夹
                {
                    countdir += 1;
                    Searchfile(f.FullName); //递归调用

                }
                else
                {
                    countfile += 1;
                }
            }
           
            return countfile;


        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnXlsBrows_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == this.openFileDialog1.ShowDialog())
            {
                this.txtXLSPath.Text = this.openFileDialog1.FileName.ToString();
                FileInfo fi = new FileInfo(this.txtXLSPath.Text.ToString());

                this.txtSaveAs.Text = fi.DirectoryName+"\\" + fi.Name.Split('.')[0].ToString() + DateTime.Now.DayOfYear.ToString()+ fi.Extension.ToString();

                //读取村组名
                try
                {
                    this.txtCun.Text = fi.Name.Substring(0, fi.Name.IndexOf("村") + 1).ToString();
                    this.txtCunzuming.Text = fi.Name.Substring(fi.Name.IndexOf("村"), fi.Name.IndexOf("组") - fi.Name.IndexOf("村") + 1).ToString();
                }
                catch
                {
                    this.txtCun.Text = "沿河村";
                    this.txtCunzuming.Text = "三河沟组";
                }

            }
        }

        private void btnPicBrows_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == this.folderBrowserDialog1.ShowDialog())
            {
                this.txtPicPath.Text = this.folderBrowserDialog1.SelectedPath.ToString();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (this.chkIsPic.Checked)
            {
                this.label2.Text = "图片路径";

            }
            else
            {
                this.label2.Text = "CAD路径";
            }
        }

        private void chkIsPic_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkIsPic.Checked)
            {
                this.label2.Text = "图片路径";

            }
            else
            {
                this.label2.Text = "CAD路径";
            }
        }
    }
}
