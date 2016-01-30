using System;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication1
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public static void OpenExcelDocs2(string filename, double[] content)
        {

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application(); //引用Excel对象
            Microsoft.Office.Interop.Excel.Workbook book = excel.Workbooks.Open(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                , Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);   //引用Excel工作簿
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets.get_Item(1); ;  //引用Excel工作页面
            excel.Visible = false;

            sheet.Cells[24, 3] = content[1];
            sheet.Cells[25, 3] = content[0];

            book.Save();
            book.Close(Type.Missing, Type.Missing, Type.Missing);
            excel.Quit();  //应用程序推出，但是进程还在运行

            IntPtr t = new IntPtr(excel.Hwnd);          //杀死进程的好方法，很有效
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();

            //sheet = null;
            //book = null;
            //excel = null;   //不能杀死进程

            //System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);  //可以释放对象，但是不能杀死进程
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);


        }
        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            m_objBook.Close(false);
           
            m_objExcel.Quit();
          
            try
            {
                IntPtr t = new IntPtr(m_objExcel.Hwnd);
                int k = 0;
                GetWindowThreadProcessId(t, out k);
                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
                p.Kill();
            }
            catch
            { }


            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnRun = new System.Windows.Forms.Button();
            this.txtXLSPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPicPath = new System.Windows.Forms.TextBox();
            this.btnXlsBrows = new System.Windows.Forms.Button();
            this.btnPicBrows = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.chkTuban = new System.Windows.Forms.CheckBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.txtSaveAs = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.labMsg = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtHeight = new System.Windows.Forms.TextBox();
            this.txtWidth = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.txtCun = new System.Windows.Forms.TextBox();
            this.txtCunzuming = new System.Windows.Forms.TextBox();
            this.chkShowTab = new System.Windows.Forms.CheckBox();
            this.chkIsPic = new System.Windows.Forms.CheckBox();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(369, 402);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 23);
            this.btnRun.TabIndex = 0;
            this.btnRun.Text = "执行";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtXLSPath
            // 
            this.txtXLSPath.Location = new System.Drawing.Point(107, 37);
            this.txtXLSPath.Name = "txtXLSPath";
            this.txtXLSPath.Size = new System.Drawing.Size(302, 21);
            this.txtXLSPath.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "Excel 路径";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "图片路径";
            // 
            // txtPicPath
            // 
            this.txtPicPath.Location = new System.Drawing.Point(107, 68);
            this.txtPicPath.Name = "txtPicPath";
            this.txtPicPath.Size = new System.Drawing.Size(302, 21);
            this.txtPicPath.TabIndex = 4;
            // 
            // btnXlsBrows
            // 
            this.btnXlsBrows.Location = new System.Drawing.Point(430, 37);
            this.btnXlsBrows.Name = "btnXlsBrows";
            this.btnXlsBrows.Size = new System.Drawing.Size(75, 23);
            this.btnXlsBrows.TabIndex = 5;
            this.btnXlsBrows.Text = "浏览";
            this.btnXlsBrows.UseVisualStyleBackColor = true;
            this.btnXlsBrows.Click += new System.EventHandler(this.btnXlsBrows_Click);
            // 
            // btnPicBrows
            // 
            this.btnPicBrows.Location = new System.Drawing.Point(430, 66);
            this.btnPicBrows.Name = "btnPicBrows";
            this.btnPicBrows.Size = new System.Drawing.Size(75, 23);
            this.btnPicBrows.TabIndex = 6;
            this.btnPicBrows.Text = "浏览";
            this.btnPicBrows.UseVisualStyleBackColor = true;
            this.btnPicBrows.Click += new System.EventHandler(this.btnPicBrows_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "Excel";
            // 
            // chkTuban
            // 
            this.chkTuban.AutoSize = true;
            this.chkTuban.Location = new System.Drawing.Point(107, 128);
            this.chkTuban.Name = "chkTuban";
            this.chkTuban.Size = new System.Drawing.Size(132, 16);
            this.chkTuban.TabIndex = 7;
            this.chkTuban.Text = "图斑序号和图号一致";
            this.chkTuban.UseVisualStyleBackColor = true;
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 12;
            this.listBox1.Location = new System.Drawing.Point(22, 211);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(129, 148);
            this.listBox1.TabIndex = 8;
            // 
            // txtSaveAs
            // 
            this.txtSaveAs.Location = new System.Drawing.Point(107, 101);
            this.txtSaveAs.Name = "txtSaveAs";
            this.txtSaveAs.Size = new System.Drawing.Size(302, 21);
            this.txtSaveAs.TabIndex = 9;
            this.txtSaveAs.Text = "d:\\\\aa.xls";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(23, 110);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 10;
            this.label3.Text = "另存为";
            // 
            // labMsg
            // 
            this.labMsg.AutoSize = true;
            this.labMsg.Location = new System.Drawing.Point(23, 397);
            this.labMsg.Name = "labMsg";
            this.labMsg.Size = new System.Drawing.Size(125, 12);
            this.labMsg.TabIndex = 11;
            this.labMsg.Text = "Mapinfo导出图片或CAD";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtHeight);
            this.groupBox1.Controls.Add(this.txtWidth);
            this.groupBox1.Font = new System.Drawing.Font("宋体", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox1.Location = new System.Drawing.Point(39, 150);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(175, 55);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "插入图片大小";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 40);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(14, 9);
            this.label5.TabIndex = 3;
            this.label5.Text = "高";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 13);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(14, 9);
            this.label4.TabIndex = 2;
            this.label4.Text = "宽";
            // 
            // txtHeight
            // 
            this.txtHeight.Location = new System.Drawing.Point(28, 37);
            this.txtHeight.Name = "txtHeight";
            this.txtHeight.Size = new System.Drawing.Size(49, 18);
            this.txtHeight.TabIndex = 1;
            this.txtHeight.Text = "100";
            // 
            // txtWidth
            // 
            this.txtWidth.Location = new System.Drawing.Point(28, 13);
            this.txtWidth.Name = "txtWidth";
            this.txtWidth.Size = new System.Drawing.Size(49, 18);
            this.txtWidth.TabIndex = 0;
            this.txtWidth.Text = "100";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(22, 371);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(483, 23);
            this.progressBar1.TabIndex = 13;
            // 
            // txtCun
            // 
            this.txtCun.Location = new System.Drawing.Point(249, 158);
            this.txtCun.Name = "txtCun";
            this.txtCun.Size = new System.Drawing.Size(100, 21);
            this.txtCun.TabIndex = 14;
            this.txtCun.Text = "沿河村";
            // 
            // txtCunzuming
            // 
            this.txtCunzuming.Location = new System.Drawing.Point(249, 183);
            this.txtCunzuming.Name = "txtCunzuming";
            this.txtCunzuming.Size = new System.Drawing.Size(100, 21);
            this.txtCunzuming.TabIndex = 15;
            this.txtCunzuming.Text = "三河沟组";
            // 
            // chkShowTab
            // 
            this.chkShowTab.AutoSize = true;
            this.chkShowTab.Checked = true;
            this.chkShowTab.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkShowTab.Location = new System.Drawing.Point(271, 128);
            this.chkShowTab.Name = "chkShowTab";
            this.chkShowTab.Size = new System.Drawing.Size(60, 16);
            this.chkShowTab.TabIndex = 16;
            this.chkShowTab.Text = "显示表";
            this.chkShowTab.UseVisualStyleBackColor = true;
            // 
            // chkIsPic
            // 
            this.chkIsPic.AutoSize = true;
            this.chkIsPic.Location = new System.Drawing.Point(369, 128);
            this.chkIsPic.Name = "chkIsPic";
            this.chkIsPic.Size = new System.Drawing.Size(114, 16);
            this.chkIsPic.TabIndex = 17;
            this.chkIsPic.Text = "插入图片否为cad";
            this.chkIsPic.UseVisualStyleBackColor = true;
            this.chkIsPic.CheckedChanged += new System.EventHandler(this.chkIsPic_CheckedChanged);
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.ItemHeight = 12;
            this.listBox2.Location = new System.Drawing.Point(249, 211);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(168, 148);
            this.listBox2.TabIndex = 18;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(517, 437);
            this.Controls.Add(this.listBox2);
            this.Controls.Add(this.chkIsPic);
            this.Controls.Add(this.chkShowTab);
            this.Controls.Add(this.txtCunzuming);
            this.Controls.Add(this.txtCun);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.labMsg);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtSaveAs);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.chkTuban);
            this.Controls.Add(this.btnPicBrows);
            this.Controls.Add(this.btnXlsBrows);
            this.Controls.Add(this.txtPicPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtXLSPath);
            this.Controls.Add(this.btnRun);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.TextBox txtXLSPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPicPath;
        private System.Windows.Forms.Button btnXlsBrows;
        private System.Windows.Forms.Button btnPicBrows;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.CheckBox chkTuban;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.TextBox txtSaveAs;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label labMsg;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtHeight;
        private System.Windows.Forms.TextBox txtWidth;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox txtCun;
        private System.Windows.Forms.TextBox txtCunzuming;
        private System.Windows.Forms.CheckBox chkShowTab;
        private System.Windows.Forms.CheckBox chkIsPic;
        private System.Windows.Forms.ListBox listBox2;
    }
}

