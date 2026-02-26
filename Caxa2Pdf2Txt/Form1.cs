
using Interop.CaxaCappInfo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Caxa2Pdf2Txt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string basePath = @"C:\Users\admin\Desktop\caxa202\";

            Interop.CaxaCappInfo.CxCappCnvrtTool convrt = null;

            try
            {
                convrt = new Interop.CaxaCappInfo.CxCappCnvrtTool();
                convrt.Init();

                foreach (var inputPath in Directory.EnumerateFiles(basePath, "*.cxp", SearchOption.AllDirectories))
                {
                    try
                    {
                        string outputPath = Path.ChangeExtension(inputPath, ".pdf");

                        convrt.OpenCxpFile(inputPath, "");
                        convrt.SaveAsPDF(outputPath);
                        convrt.CloseFile();
                    }
                    catch (Exception ex)
                    {
                        // 单个文件失败不影响整体
                        Console.WriteLine($"转换失败: {inputPath}");
                        Console.WriteLine(ex.Message);
                    }
                }

                MessageBox.Show("全部转换完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化失败: " + ex.Message);
            }
            finally
            {
                if (convrt != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(convrt);
                    convrt = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string basePath = @"C:\Users\admin\Desktop\caxa202\";

            Interop.CaxaCappInfo.CAPPInfo cappinfo = null;

            try
            {
                cappinfo = new Interop.CaxaCappInfo.CAPPInfo();

                var files = Directory.GetFiles(basePath, "*.cxp", SearchOption.AllDirectories);

                foreach (var inputPath in files)
                {
                    string outputPath = inputPath + ".txt";

                    cappinfo.GetCappInfoToTxt(inputPath, outputPath);
                }

                MessageBox.Show("全部转换完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误：" + ex.Message);
            }
            finally
            {
                if (cappinfo != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(cappinfo);
                    cappinfo = null;
                    GC.Collect();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var convrt = new Interop.CaxaCappInfo.CxCappCnvrtTool();
            convrt.Init();
            convrt.OpenCxpFile("C:\\Users\\admin\\Desktop\\caxa202\\test.cxp", "");
            convrt.SaveAsPDF("C:\\Users\\admin\\Desktop\\caxa202\\test.cxp.pdf");
            convrt.CloseFile();

        }

        private void button4_Click(object sender, EventArgs e)
        {
           Interop.CaxaCappInfo.ICAPPInfo cappInfo = new Interop.CaxaCappInfo.CAPPInfo();
           Interop.CaxaCappInfo.ICAPPXmlInfo xmlinfo = (Interop.CaxaCappInfo.ICAPPXmlInfo)cappInfo;
            string bsFilePath = @"C:\Users\admin\Desktop\caxa202\test.cxp";     // cxp文件路径
            string bsPassword = "";                // 密码
            string bstCardname = "工序目录";          // 卡片名称
            string bstColInfo = "编制日期^0226&校对日期^0227"; // 列名^内容
            int m_iRowNum = 0;                     // 行号
            xmlinfo.OpenFile(bsFilePath, bsPassword);
            xmlinfo.WriteTxtInfoToCard(bstCardname,
                bstColInfo,
                m_iRowNum);
            xmlinfo.CloseFile();


        }

        private void button5_Click(object sender, EventArgs e)
        {
            Interop.CaxaCappInfo.ICAPPInfo cappInfo = new Interop.CaxaCappInfo.CAPPInfo();
            Interop.CaxaCappInfo.ICAPPXmlInfo xmlinfo = (Interop.CaxaCappInfo.ICAPPXmlInfo)cappInfo;
            string bsFilePath = @"C:\Users\admin\Desktop\caxa202\test.cxp";     // cxp文件路径
            string imageFilePath = @"C:\Users\admin\Desktop\caxa202\张三.png";     // 签字图片路径
            string bsPassword = "";                    // 密码
            string bstCardname = "工序目录";          // 卡片名称
            string bstColInfo = "编制";              // 列名^内容
            int m_iRowNum = 0;
            byte alignMode = 1;                    //对齐方式，取值012345，分别对应左，中，右，垂直上，垂BOOL  
            byte bFillInMode = 1;// 行号
            xmlinfo.OpenFile(bsFilePath, bsPassword);
            xmlinfo.WriteImgToCard(imageFilePath,bstCardname,
                bstColInfo,
                m_iRowNum, alignMode, bFillInMode);
            xmlinfo.CloseFile();

        }
    }
}
