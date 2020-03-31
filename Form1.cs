using Slb.Ocean.Core;
using Slb.Ocean.Petrel.DomainObject.Well;
using Slb.Ocean.Petrel.UI;
using Slb.Ocean.Petrel.Workflow;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static OceanReadingData.read;
using PetrelLogger = Slb.Ocean.Petrel.PetrelLogger;
using Excel = Microsoft.Office.Interop.Excel;
using Slb.Ocean.Petrel;
using Slb.Ocean.Petrel.DomainObject.Seismic;

namespace OceanReadingData
{

    public partial class Form1 : Form
    {
        DropTarget blueArrow;
        public Form1()
        {
            InitializeComponent();
            this.blueArrow = new Slb.Ocean.Petrel.UI.DropTarget();
            blueArrow.DragDrop += new System.Windows.Forms.DragEventHandler(dropTarget1_Load);

        }

        /* private void dropTarget1_Load(object sender, DragEventArgs e)
         {
             BoreholeCollection g = e.Data.GetData(typeof(BoreholeCollection)) as BoreholeCollection;
             if (g == null) return;
             INameInfoFactory nameFact;
             nameFact = CoreSystem.GetService<INameInfoFactory>(g);
             NameInfo nameInfo = nameFact.GetNameInfo(g);
             presentationBox1.Text = nameInfo.Name;
             IImageInfoFactory imageFact;
             imageFact = CoreSystem.GetService<IImageInfoFactory>(g);
             ImageInfo imageInfo = imageFact.GetImageInfo(g);
             presentationBox1.Image = imageInfo.TypeImage;
             // use Tag later to get object from PresentationBox
             presentationBox1.Tag = g;
         }
         */

        Arguments arguments;
        WorkflowRuntimeContext context;
        private void button1_Click(object sender, EventArgs e)
        {

            //progressBar1.Visible = true;
            #region
            StreamReader rd = File.OpenText(@"d:\ak.txt");
            string restOfStream = rd.ReadToEnd();
            rd.Close();
            //输出DataTable中保存的数组                                            
            //foreach (DataRow r in tb.Rows)                                     
            //Console.WriteLine("assa{0}",restOfStream);          
            string[] arr1 = restOfStream.Split('(', ')', ',', '\n', '[', ']');
            int a;                                                            
            double bb;                                                       
            bb = arr1.Length / 6;                                            
            a = (int)Math.Floor(bb);                                         
            int N = a;


            #region
            Cursor c = System.Windows.Forms.Cursors.Cross;
            IProgress p = PetrelLogger.NewProgress(0, 25 + 3 * N,
            ProgressType.Cancelable, c);

            int max = 25 + 3 * N;
            int step = 1;
            //progressBar1.Maximum = 25+3*N;//设置最大长度值           
                    //progressBar1.Value = 0;//设置当前值                      
                    //progressBar1.Step = 1;//设置没次增长多少                   
            p.ProgressStatus += step * 10;
            //progressBar1.Value += progressBar1.Step*10;//让进度条增加一次                    10
            //SeismicProject p = presentationBox1.Tag as SeismicProject;
            BoreholeCollection bc = presentationBox1.Tag as BoreholeCollection;
            #endregion

            #region
            //FileStream fs = new FileStream("D:\\ak.txt", FileMode.Create);

            //excel open


            /*//brfore
            // 利用SaveFileDialog，让用户指定文件的路径名
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "文本文件|*.xlsx";
            Gloable.excel_path = saveDlg.FileName;
            if (saveDlg.ShowDialog() == DialogResult.OK)
            { }
            // 创建文件，将textBox1中的内容保存到文件中
            // saveDlg.FileName 是用户指定的文件路径
            FileStream fs_excel = File.Open(Gloable.excel_path,
                    FileMode.Create,
                    FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs_excel);

                // 保存中所有内容

                //关闭文件
                sw.Flush();
                sw.Close();
                fs_excel.Close();
                */

            // 将路径改为txt的路径
            /* string[] split_path = saveDlg.FileName.Split('.');
             string txt_name = split_path[0] + "_for_to_excel" + ".txt";
             FileStream fs_txt = new FileStream(txt_name, FileMode.Create);
             fs_txt.Close();
             Gloable.txt_path = txt_name;
            /* StreamReader rd = File.OpenText(txt_name);
             string restOfStream = rd.ReadToEnd();
             rd.Close();*/
            //输出DataTable中保存的数组


            //excel 路径创建文件

            #endregion
            //fs.Close();



            //excel 路径创建文件
            try
            {
                string fileTest1 = Gloable.excel_path;
                textBox1.Text = fileTest1;
                string fileTest = Gloable.excel_path;
                if (File.Exists(fileTest))
                {
                    //File.Delete(fileTest);
                }
                else
                {
                    Microsoft.Office.Interop.Excel.Application oApp;
                    Microsoft.Office.Interop.Excel.Worksheet oSheet;
                    Microsoft.Office.Interop.Excel.Workbook oBook;

                    oApp = new Microsoft.Office.Interop.Excel.Application();
                    oBook = oApp.Workbooks.Add();
                    oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oBook.Worksheets.get_Item(1);
                    //oSheet.Cells[1, 1] = "some value";
                    oBook.SaveAs(fileTest);
                    oBook.Close();
                    oApp.Quit();
                    //app.Quit();

                    //释放掉多余的excel进程
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp);
                    oApp = null;
                }
            }
            finally { }
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(Gloable.excel_path);
            Microsoft.Office.Interop.Excel.Worksheet ws = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            p.ProgressStatus += step * 10;
            //progressBar1.Value += progressBar1.Step*10;//让进度条增加一次                               20

            BoreholePropertyCollection bpc = bc.BoreholePropertyCollection;

            //示例
            #region
            /*foreach (Borehole bh in bc)
            {

                IPropertyAccess propAccess = bh.PropertyAccess;
                BoreholePropertyType OffsetType = WellKnownBoreholePropertyTypes.Offset;
                BoreholeProperty Offset =    bpc.GetWellKnownProperty(OffsetType);
                BoreholeProperty wellheadX = bpc.GetWellKnownProperty(WellKnownBoreholePropertyTypes.WellHeadX);
                BoreholeProperty wellheadY = bpc.GetWellKnownProperty(WellKnownBoreholePropertyTypes.WellHeadY);
       //         BoreholeProperty boreholename = bpc.GetWellKnownProperty(WellKnownBoreholePropertyTypes.BoreholeName);
                double[] value = { 100, 33, 4, 3, 5, 6, 6, 6, 6, 4, 4, 3, 3, 2, 1, 1, 4, 12, 412, 4, 214, 3, 23 };
                PetrelLogger.InfoOutputWindow("new print changed");
                //num_bh is the row of the data
                using (ITransaction tr = DataManager.NewTransaction())
                {
                    tr.Lock(bh);
                    bh.PropertyAccess.SetPropertyValue(Offset, value[num_bh]);
                    tr.Commit();
                }
                num_bh++;
            }*/
            #endregion

            //read property
            #region 
            Slb.Ocean.Petrel.PetrelLogger.InfoOutputWindow("properties name");
            foreach (BoreholeProperty bpp in bpc.Properties)
            {
                Slb.Ocean.Petrel.PetrelLogger.InfoOutputWindow(bpp.Name.ToString());
                //progressBar1.Value += progressBar1.Step;//让进度条增加一次                               20+N
            }
            Slb.Ocean.Petrel.PetrelLogger.InfoOutputWindow("properties PropertyType");
            foreach (BoreholeProperty bpp in bpc.Properties)
            {
                PetrelLogger.InfoOutputWindow(bpp.ToString());
                //progressBar1.Value += progressBar1.Step;//让进度条增加一次                               20+2N
            }
            PetrelLogger.InfoOutputWindow("properties iswirte");
            foreach (BoreholeProperty bpp in bpc.Properties)
            {
                PetrelLogger.InfoOutputWindow(bpp.IsWritable.ToString());
                //progressBar1.Value += progressBar1.Step;//让进度条增加一次                               20+3N
            }
            PetrelLogger.InfoOutputWindow("properties Template");
            foreach (BoreholeProperty bpp in bpc.Properties)
            {
                PetrelLogger.InfoOutputWindow(bpp.Template.ToString());
               // progressBar1.Value += progressBar1.Step;//让进度条增加一次                               20+4N
            }
            PetrelLogger.InfoOutputWindow("properties DATATYPE");
            foreach (BoreholeProperty bpp in bpc.Properties)
            {
                Slb.Ocean.Petrel.PetrelLogger.InfoOutputWindow(bpp.DataType.ToString());
                //progressBar1.Value += progressBar1.Step;//让进度条增加一次                               20+5N
            }
            #endregion



           



            //写入excel

            // 新建一个DataTable
            //DataTable tb = new DataTable();
            // 添加一列用于存放读入的浮点数
            // DataColumn c = tb.Columns.Add("Value", typeof(double));

            // 打开文件准备读取数据   
            // 打开文件准备读取数据   
            
            try
            {
                for (int i = 1; i <= arr1.Length / 6; i++)
                {
                    ws.Cells[i + 2, 2] = arr1[6 * (i - 1)];
                    int b = 6 * (i - 1) + 1;
                    ws.Cells[i + 2, 1] = Convert.ToString(i);
                    ws.Cells[i + 2, 3] = arr1[6 * (i - 1) + 1];
                    ws.Cells[i + 2, 4] = arr1[6 * (i - 1) + 2];
                    ws.Cells[i + 2, 5] = arr1[6 * (i - 1) + 3];
                    ws.Cells[i + 2, 6] = arr1[6 * (i - 1) + 4];
                    //ws.Cells[i + 2, 7] = arr1[6 * (i - 1) + 5];
                    p.ProgressStatus += step ;
                    //progressBar1.Value += progressBar1.Step;//让进度条增加一次                               20+N

                }

            }
            finally { }

            try
            {
                ws.get_Range("A1", "H1").Merge(ws.get_Range("A1", "F1").MergeCells);
                ws.Cells[1, 1] = "WELL MANAGEMENT";
                Excel.Range excelRange1 = ws.get_Range("A1", "H1");
                excelRange1.Font.ColorIndex = 9;
                excelRange1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelRange1.Font.Size = 18;
                p.ProgressStatus += step * 5;
                //progressBar1.Value += progressBar1.Step*5;//让进度条增加一次                               25+N
            }
            finally { }


            try
            {
                ws.Cells[2, 1] = "Number";
                ws.Cells[2, 2] = "ID";
                ws.Cells[2, 3] = "X";
                ws.Cells[2, 4] = "Y";
                ws.Cells[2, 5] = "Well datum name";
                ws.Cells[2, 6] = "Well datum value";
                Excel.Range excelRange2 = ws.get_Range("A2", "H2");
                excelRange2.Font.Italic = true;
                excelRange2.Font.Bold = true;
                excelRange2.Font.Size = 11;
                excelRange2.ColumnWidth = 20;
                excelRange2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            finally { }

            try
            {
                for (int i = 1; i <= arr1.Length; i++)
                {
                    Excel.Range excelRange11 = ws.get_Range("A2", "A2");
                    excelRange11.Cells.Interior.Color = System.Drawing.Color.FromArgb(255, 0, 0).ToArgb();
                    Excel.Range excelRange22 = ws.get_Range("B2", "B2");
                    excelRange22.Cells.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 0).ToArgb();
                    Excel.Range excelRange33 = ws.get_Range("C2", "C2");
                    excelRange33.Cells.Interior.Color = System.Drawing.Color.FromArgb(255, 0, 255).ToArgb();
                    Excel.Range excelRange44 = ws.get_Range("D2", "D2");
                    excelRange44.Cells.Interior.Color = System.Drawing.Color.FromArgb(0, 255, 153).ToArgb();
                    Excel.Range excelRange55 = ws.get_Range("E2", "E2");
                    excelRange55.Cells.Interior.Color = System.Drawing.Color.FromArgb(0, 150, 255).ToArgb();
                    Excel.Range excelRange66 = ws.get_Range("F2", "F2");
                    excelRange66.Cells.Interior.Color = System.Drawing.Color.FromArgb(0, 204, 153).ToArgb();

                    Excel.Range excelRangeall = ws.get_Range("A2", "F100");
                    excelRangeall.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    p.ProgressStatus += step * 2;
                    //progressBar1.Value += progressBar1.Step * 2;//让进度条增加一次                               25+ N+2N
                    //PetrelLogger.InfoOutputWindow(progressBar1.Value.ToString());
                    if (p.ProgressStatus >= 3 * N + 22)
                    {
                        p.ProgressStatus = 3 * N + 23;
                    }
                   

                }
            }
            finally { }
            p.ProgressStatus = max;
            //progressBar1.Value = progressBar1.Maximum;
            p.Dispose();
            MessageBox.Show("The file has been successfully saved to " + Gloable.excel_path);
            PetrelLogger.InfoOutputWindow("The file has been successfully saved to " + Gloable.excel_path);
            
            //progressBar1.Value = 0;//设置当前值
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            excel = null;
            GC.Collect();//垃圾回收
                         //清空缓冲区、关闭流

            // 提示用户：文件保存的位置和文件名

            //MessageBox.Show("文件已成功保存到" + "D:\\openexcel.xlsx");
            //PetrelLogger.InfoOutputWindow("文件已成功保存到" + "D:\\openexcel.xlsx");
            



            //注释代码
            #region
            /*
            // Create borehole collections
            BoreholeCollection boreholeCollection = read.GetOrCreateBoreholeCollection();
            BoreholeCollection someBoreholes = read.GetOrCreateBoreholeCollection(boreholeCollection, "Some boreholes");
            BoreholeCollection someOtherBoreholes = MyHelperClass.CreateBoreholeCollection(boreholeCollection, "Some other boreholes");

            // Create boreholes
            Borehole aaa = MyHelperClass.CreateMainBorehole(someBoreholes, "AAA");
            Borehole bbb = MyHelperClass.CreateLateralBorehole(boreholeCollection, "BBB", aaa);
            Borehole ccc = MyHelperClass.CreateLateralBorehole(someOtherBoreholes, "CCC", bbb);
            Borehole ddd = MyHelperClass.CreateLateralBorehole(someBoreholes, "DDD", aaa);
            Borehole eee = MyHelperClass.CreateMainBorehole(someOtherBoreholes, "EEE");

            // Get the lateral/tie-in relations
            bool isMainBorehole = aaa.IsMainBorehole; // true
            int aaaLateralCount = aaa.LateralBoreholeCount; // 2
            IEnumerable aaaLaterals = aaa.LateralBoreholes; // { 'BBB', 'DDD' }
            isMainBorehole = ddd.IsMainBorehole; // false
            Borehole dddTieInBorehole = ddd.TieInBorehole; // 'AAA'

            Borehole eeeTieInBorehole = eee.TieInBorehole; // Borehole.NullObject
            int eeeLateralCount = eee.LateralBoreholeCount; // 0

            // Set well head, working reference level and reference levels
            MyHelperClass.SetWellHead(aaa, new Point2(100.0, 200.0));
            IList<ReferenceLevel> referenceLevels = new List<ReferenceLevel>
            {
                new ReferenceLevel("KB", 25.0, "Kelly bushing"),
                new ReferenceLevel("RF", 10.0, "Rig floor")
            };
            MyHelperClass.SetReferenceLevels(aaa, referenceLevels);
            MyHelperClass.SetWorkingReferenceLevel(aaa, referenceLevels[1]);

            // All the laterals share these properties
            Point2 dddWellHead = ddd.WellHead; // (100.0, 200.0)
            int bbbReferenceLevelCount = bbb.ReferenceLevelsCount; // 2
            IEnumerable<ReferenceLevel> cccReferenceLevels = ccc.ReferenceLevels; // { ("KB", 25.0, "Kelly bushing"), ("RF", 10.0, "Rig floor") }
            ReferenceLevel cccReferenceLevel = ccc.WorkingReferenceLevel; // ("RF", 10.0, "Rig floor")

            // Trying to set these properties on a lateral throws exception
            MyHelperClass.SetWellHead(ddd, new Point2(100.0, 10.0)); // InvalidOperationException
            MyHelperClass.SetWorkingReferenceLevel(ccc, referenceLevels[0]); // InvalidOperationException
            MyHelperClass.SetReferenceLevels(bbb, referenceLevels); // InvalidOperationException
            */
            #endregion

            #endregion
        }

        private void dropTarget1_Load(object sender, EventArgs e)
        {
             //automatically
            Slb.Ocean.Petrel.DomainObject.Project proj = PetrelProject.PrimaryProject;
            // get the root of all domain objects 
            WellRoot wr = WellRoot.Get(proj);
            BoreholeCollection bc = wr.BoreholeCollection;
            // SeismicRoot root = SeismicRoot.Get(proj);
            //SeismicProject g = root.SeismicProject;
            if (bc == null) return;
            INameInfoFactory nameFact;
            nameFact = CoreSystem.GetService<INameInfoFactory>(bc);
            NameInfo nameInfo = nameFact.GetNameInfo(bc);
            presentationBox1.Text = nameInfo.Name;
            IImageInfoFactory imageFact;
            imageFact = CoreSystem.GetService<IImageInfoFactory>(bc);
            ImageInfo imageInfo = imageFact.GetImageInfo(bc);
            presentationBox1.Image = imageInfo.TypeImage;
            // use Tag later to get object from PresentationBox
            presentationBox1.Tag = bc;
            
        }

        private void button2_Click(object sender, EventArgs e)
        {

            

            BoreholeCollection bc = presentationBox1.Tag as BoreholeCollection;
            //SeismicProject p = presentationBox1.Tag as SeismicProject;
            //SeismicProject p = presentationBox1.Tag as SeismicProject;
            //BoreholeCollection bc = presentationBox1.Tag as BoreholeCollection;
            int count = 0;
            foreach(Borehole bbh in bc)
            {
                count++;
                
            }
            int N = count;
            Gloablee.NN = N;


            #region
            //progressBar1.Maximum = 10+6*N;//设置最大长度值  10+10N
            //progressBar1.Value = 0;//设置当前值
            //progressBar1.Step = 1;//设置没次增长多少
                                  //progressBar1.Value += progressBar1.Step * 10;//让进度条增加一次         



            Cursor c = System.Windows.Forms.Cursors.Cross;
            IProgress p = PetrelLogger.NewProgress(0, 10 + 6 * N,
            ProgressType.Cancelable, c);
            int step = 1;
            int max = 10 + 6 * N;
      
                //p.ProgressStatus
                if (p.IsCanceled)
                {
                    // clean up any resources or data
                    //break;
                }
               
            
            


            #endregion



            IList<ReferenceLevel> referenceLevels = new List<ReferenceLevel>
                {
                    new ReferenceLevel("KB", 25.0, "Kelly bushing"),
                    new ReferenceLevel("RF", 10.0, "Rig floor")
                };
            p.ProgressStatus += 10;
            //progressBar1.Value += 10;//让进度条增加一次                               10
            #region

            //FileStream fs = new FileStream("D:\\ak.txt", FileMode.Create);

            //excel open


            /*//brfore
            // 利用SaveFileDialog，让用户指定文件的路径名
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "文本文件|*.xlsx";
            Gloable.excel_path = saveDlg.FileName;
            if (saveDlg.ShowDialog() == DialogResult.OK)
            { }
            // 创建文件，将textBox1中的内容保存到文件中
            // saveDlg.FileName 是用户指定的文件路径
            FileStream fs_excel = File.Open(Gloable.excel_path,
                    FileMode.Create,
                    FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs_excel);

                // 保存中所有内容

                //关闭文件
                sw.Flush();
                sw.Close();
                fs_excel.Close();
                */

            // 将路径改为txt的路径
            /* string[] split_path = saveDlg.FileName.Split('.');
             string txt_name = split_path[0] + "_for_to_excel" + ".txt";
             FileStream fs_txt = new FileStream(txt_name, FileMode.Create);
             fs_txt.Close();
             Gloable.txt_path = txt_name;
            /* StreamReader rd = File.OpenText(txt_name);
             string restOfStream = rd.ReadToEnd();
             rd.Close();*/
            //输出DataTable中保存的数组
            #endregion
            // PetrelLogger.InfoOutputWindow(bc.);
            #region
            //excel 路径创建文件
            /*
            try
            {
                //string fileTest = "D:\\openexcel.xlsx";
                string fileTest = Gloable.excel_path;
                if (File.Exists(fileTest))
                {
                    File.Delete(fileTest);
                }
                Microsoft.Office.Interop.Excel.Application oApp;
                Excel.Worksheet oSheet;
                Excel.Workbook oBook;

                oApp = new Excel.Application();
                oBook = oApp.Workbooks.Add();
                oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);
                //oSheet.Cells[1, 1] = "some value";
                oBook.SaveAs(fileTest);
                oBook.Close();
                oApp.Quit();
            }
            finally { }
            */
            #endregion
            //fs.Close();

            //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(Gloable.excel_path); //"\\openexcel.xlsx"
            //Microsoft.Office.Interop.Excel.Worksheet ws = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            BoreholePropertyCollection bpc = bc.BoreholePropertyCollection;
           
            //示例
            #region
            /*foreach (Borehole bh in bc)
            {

                IPropertyAccess propAccess = bh.PropertyAccess;
                BoreholePropertyType OffsetType = WellKnownBoreholePropertyTypes.Offset;
                BoreholeProperty Offset =    bpc.GetWellKnownProperty(OffsetType);
                BoreholeProperty wellheadX = bpc.GetWellKnownProperty(WellKnownBoreholePropertyTypes.WellHeadX);
                BoreholeProperty wellheadY = bpc.GetWellKnownProperty(WellKnownBoreholePropertyTypes.WellHeadY);
       //         BoreholeProperty boreholename = bpc.GetWellKnownProperty(WellKnownBoreholePropertyTypes.BoreholeName);
                double[] value = { 100, 33, 4, 3, 5, 6, 6, 6, 6, 4, 4, 3, 3, 2, 1, 1, 4, 12, 412, 4, 214, 3, 23 };
                PetrelLogger.InfoOutputWindow("new print changed");
                //num_bh is the row of the data
                using (ITransaction tr = DataManager.NewTransaction())
                {
                    tr.Lock(bh);
                    bh.PropertyAccess.SetPropertyValue(Offset, value[num_bh]);
                    tr.Commit();
                }
                num_bh++;
            }*/
            #endregion

            //read property
            #region 
            Slb.Ocean.Petrel.PetrelLogger.InfoOutputWindow("properties name");
            foreach (BoreholeProperty bpp in bpc.Properties)
            {
                Slb.Ocean.Petrel.PetrelLogger.InfoOutputWindow(bpp.Name.ToString());
                //progressBar1.Value += progressBar1.Step ;//让进度条增加一次                               10+N
            }
            Slb.Ocean.Petrel.PetrelLogger.InfoOutputWindow("properties PropertyType");
            foreach (BoreholeProperty bpp in bpc.Properties)
            {
                PetrelLogger.InfoOutputWindow(bpp.ToString());
                //progressBar1.Value += progressBar1.Step;//让进度条增加一次                               10+2N
            }
            PetrelLogger.InfoOutputWindow("properties iswirte");
            foreach (BoreholeProperty bpp in bpc.Properties)
            {
                PetrelLogger.InfoOutputWindow(bpp.IsWritable.ToString());
                //progressBar1.Value += progressBar1.Step;//让进度条增加一次                               10+3N
            }
            PetrelLogger.InfoOutputWindow("properties Template");
            foreach (BoreholeProperty bpp in bpc.Properties)
            {
                PetrelLogger.InfoOutputWindow(bpp.Template.ToString());
               // progressBar1.Value += progressBar1.Step;//让进度条增加一次                               10+4N
            }
            PetrelLogger.InfoOutputWindow("properties DATATYPE");
            foreach (BoreholeProperty bpp in bpc.Properties)
            {
                Slb.Ocean.Petrel.PetrelLogger.InfoOutputWindow(bpp.DataType.ToString());
               // progressBar1.Value += progressBar1.Step;//让进度条增加一次                               10+5N
            }
            #endregion



            //写入txt

            FileStream fs = new FileStream("D:\\ak.txt", FileMode.Create);

            //reading the data
            // BoreholeProperty bp = bp;

            foreach (Borehole bh in bc)
            {
                PetrelLogger.InfoOutputWindow(bh.Name);
                IEnumerable<ReferenceLevel> bh_refer = bh.ReferenceLevels;// read the KB and RF

                ReferenceLevel bh_RF = bh.WorkingReferenceLevel; // ("RF", 10.0, "Rig floor")
                                                                 // ReferenceLevel bh_RF = bh.refer // ("RF", 10.0, "Rig floor")

                #region

                PetrelLogger.InfoOutputWindow(bh.Name);
                byte[] data2 = System.Text.Encoding.Default.GetBytes(bh.Name.ToString());
                byte[] data3 = System.Text.Encoding.Default.GetBytes(bh.WellHead.ToString());
                //progressBar1.Value += progressBar1.Step*2;//让进度条增加一次                               10+2N
                //borehole bh = sample in 
                /*  using (ITransaction t = DataManager.NewTransaction())
                  {
                      t.Lock(bh);
                      bh.Name = "123";
                     //bh.Extensions.Add(xyzColl);
                      t.Commit();
                  }
                  */

                #endregion

               // progressBar1.Value += progressBar1.Step;//让进度条增加一次                          
                //bh_refer is the 
                PetrelLogger.InfoOutputWindow("print each_bh_reder in bh_refer");

                //print before changing
                PetrelLogger.InfoOutputWindow("be data start to change");
                foreach (ReferenceLevel ea_bh_refer in bh_refer)
                {
                    PetrelLogger.InfoOutputWindow(ea_bh_refer.ToString());
                    
                }
                //p.ProgressStatus += step;
                //progressBar1.Value += progressBar1.Step;//让进度条增加一次                               10+3N
                PetrelLogger.InfoOutputWindow("be data has changed");
                //修改示例
                #region
                /*   using (ITransaction t = DataManager.NewTransaction())
                   {
                       t.Lock(bh_refer);
                       // each_bh_refer = referenceLevels;
                       bh_refer = referenceLevels;
                       //bh_refer.Per
                       // bh_refer = "12312";
                       //bh.Name = "123";
                       //bh.Extensions.Add(xyzColl);

                       t.Commit();
                   }*/
                #endregion
                PetrelLogger.InfoOutputWindow("data start to change");
                foreach (ReferenceLevel ea_bh_refer in bh_refer)
                {
                    PetrelLogger.InfoOutputWindow(ea_bh_refer.ToString());
                    
                }
                p.ProgressStatus += step;
                //progressBar1.Value += progressBar1.Step;//让进度条增加一次                               10+4N
                PetrelLogger.InfoOutputWindow("data has changed");




                foreach (ReferenceLevel each_bh_refer in bh_refer)
                {
                    PetrelLogger.InfoOutputWindow(each_bh_refer.ToString());
                    //read.SetReferenceLevels(bh,referenceLevels);
                    //修改 试一试

                    //获得字节数组
                    byte[] data1 = System.Text.Encoding.Default.GetBytes(each_bh_refer.ToString() + '\n');
                    // PetrelLogger.InfoOutputWindow("output  data1{0}   data2{1}   data3{2}", data1, data2, data3);
                    PetrelLogger.InfoOutputWindow("set the data");

                    //开始写入
                    //data2 name
                    fs.Write(data2, 0, data2.Length);
                    //data3 point
                    fs.Write(data3, 0, data3.Length);
                    //data1 kelly
                    fs.Write(data1, 0, data1.Length);
                    PetrelLogger.InfoOutputWindow("finish writing4");
                    //
   

                    #region
                    //修改petrol
                    /*using (ITransaction t = DataManager.NewTransaction())
                    {
                        t.Lock(bh);
                        //each_bh_refer.Name = "12312";
                        bh.Name = "123";

                        //bh.Extensions.Add(xyzColl);
                        t.Commit();
                    }*/

                    //  IEnumerable<ReferenceLevel> bh_refer = bh.ReferenceLevels;// read the KB and RF
                    //   ReferenceLevel bh_RF = bh.WorkingReferenceLevel; // ("RF", 10.0, "Rig floor")
                    /*
                           IList<ReferenceLevel> referenceLevels = new List<ReferenceLevel>
                           {
                               new ReferenceLevel("KB", 25.0, "Kelly bushing"),
                               new ReferenceLevel("RF", 10.0, "Rig floor")
                   };
                   */

                    /* using (ITransaction t = DataManager.NewTransaction())
                     {
                         t.Lock(bh_refer);
                         bh_refer = referenceLevels;
                         // bh_refer = "12312";
                         //bh.Name = "123";
                         //bh.Extensions.Add(xyzColl);

                         t.Commit();

                         PetrelLogger.InfoOutputWindow("data start to change");
                             foreach(ReferenceLevel ea_bh_refer in bh_refer)
                         {
                             PetrelLogger.InfoOutputWindow(ea_bh_refer.ToString());
                         }
                         PetrelLogger.InfoOutputWindow("data has changed");
                     }*/
                    #endregion
                   
                }
                p.ProgressStatus += step * 2;
                //progressBar1.Value += progressBar1.Step * 2;//让进度条增加一次                               10+6N
                if(p.ProgressStatus >= 8+6*N)
                {
                    p.ProgressStatus = 8 + 6 * N;
                }

            }
            p.ProgressStatus = 10 + 6 * N;
            fs.Flush();
            fs.Close();
            p.Dispose();
            //progressBar1.Value = progressBar1.Maximum;
            MessageBox.Show("Data check finished");
            //progressBar1.Value = 0;
            


        }

        private void presentationBox1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            /*
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "文本文件|*.xlsx";
            Gloable.excel_path = saveDlg.FileName;
            if (saveDlg.ShowDialog() == DialogResult.OK)
            { }
            textBox1.Text = Gloable.excel_path;
            */
            //choose the fold dirction
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "Please choose the path ";
            string foldPath = "";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                foldPath = dialog.SelectedPath;
                //MessageBox.Show("已选择文件夹:" + foldPath, "选择文件夹提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            textBox1.Text = foldPath;
            string tem = textBox2.Text;
            Gloable.excel_path = foldPath + "\\"+tem+".xlsx";
            textBox1.Text = Gloable.excel_path;

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {

            //open Excel
            #region
            PetrelLogger.InfoOutputWindow("save the data");
            //save the data
            Gloable.if_change = 1;
            //open excel 
            PetrelLogger.InfoOutputWindow(string.Format("{0} clicked", @"WellData"));
            if (MessageBox.Show("EXCEL will open", "Sure", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.UseShellExecute = true;
                p.StartInfo.FileName = Gloable.excel_path;//"D:\\openexcel.xlsx";
                p.Start();
            }
            for(int i = 1;i<101;i++)
            {
                //progressBar1.Value += progressBar1.Step * 10;//让进度条增加一次                               10
                System.Threading.Thread.Sleep(10);//暂停1秒
               
            }
         

            #endregion
        }
        static class Gloablee
        {
            public static int NN = 0;
          
        }

        private void button4_Click(object sender, EventArgs e)
        {



            //recieve data
            //open the excel      
            //TODO: Add command execution logic here
            int N = Gloablee.NN;
            #region
            //progressBar1.Maximum = 10+4*N;//设置最大长度值  10+10N
            //progressBar1.Value = 0;//设置当前值
            //progressBar1.Step = 1;//设置没次增长多少
            //progressBar1.Visible = true;

            int max = 10 + 4 * N;
            int step = 1;
            Cursor c = System.Windows.Forms.Cursors.Cross;
            IProgress p = PetrelLogger.NewProgress(0, 10 + 4 * N,
            ProgressType.Cancelable, c);
            
                //p.ProgressStatus = i;
                if (p.IsCanceled)
                {
                    //clean up any resources or data
                    //break;
                }
                // operation step ...
            
            

            #endregion

            #region
            if (Gloable.if_change == 1)
            {
                Gloable.if_change = 0;
            }
                //read the data
                PetrelLogger.InfoOutputWindow("read the data");

                //find the txt and read
                // 打开文件准备读取数据   
                PetrelLogger.InfoOutputWindow("read text");
                StreamReader rd = File.OpenText("D:\\ak_xls_txt.txt");
                PetrelLogger.InfoOutputWindow("success read");

                string restOfStream = rd.ReadToEnd();
                rd.Close();
                //输出DataTable中保存的数组
                PetrelLogger.InfoOutputWindow(restOfStream.ToString());
                //  restOfStream = restOfStream.Replace(""," ");
                string[] arr1 = restOfStream.Split(' ', '\n');
            //progressBar1.Value += progressBar1.Step * 10;//让进度条增加一次                               10
            int count = 0;
                for (int i = 0; i < arr1.Length; i++)
                {
                    if (arr1[i] == "")
                    {
                    }
                    else
                    {
                        arr1[count] = arr1[i];
                        count++;
                        PetrelLogger.InfoOutputWindow(arr1[i]);
                        //Console.WriteLine(arr1[i]);
                    

                }
                }


                //开始回写数据
                WellRoot root = WellRoot.Get(PetrelProject.PrimaryProject);
                PetrelLogger.InfoOutputWindow("BoreholeCollection bc");
                BoreholeCollection bc = root.BoreholeCollection;
                PetrelLogger.InfoOutputWindow("BoreholeCollection pbc");
                BoreholePropertyCollection bpc = bc.BoreholePropertyCollection;
                //IPropertyAccess propAccess = bh.PropertyAccess;
                PetrelLogger.InfoOutputWindow("BoreholePropertyType OffsetType");
                BoreholePropertyType OffsetType = WellKnownBoreholePropertyTypes.Offset;
                PetrelLogger.InfoOutputWindow("BoreholePropertyType Offset");
                BoreholeProperty Offset = bpc.GetWellKnownProperty(OffsetType);
                PetrelLogger.InfoOutputWindow("BoreholePropertyType wellX");
                BoreholeProperty wellheadX = bpc.GetWellKnownProperty(WellKnownBoreholePropertyTypes.WellHeadX);
                BoreholeProperty wellheadY = bpc.GetWellKnownProperty(WellKnownBoreholePropertyTypes.WellHeadY);
                //BoreholeProperty boreholename = bpc.GetWellKnownProperty(WellKnownBoreholePropertyTypes.BoreholeName);
                //double[] value = { 100, 33, 4, 3, 5, 6, 6, 6, 6, 4, 4, 3, 3, 2, 1, 1, 4, 12, 412, 4, 214, 3, 23 };
                PetrelLogger.InfoOutputWindow("new print changed");



                int num_bh = 0;
                PetrelLogger.InfoOutputWindow("changed in WellData");
                double dou = Convert.ToDouble(arr1[num_bh + 1]);
                foreach (Borehole bh in bc)
                {


                    //num_bh is the row of the data
                    using (ITransaction tr = DataManager.NewTransaction())
                    {
                        tr.Lock(bh);
                        //  bh.PropertyAccess.SetPropertyValue(boreholename, arr1[num_bh * 5 + 0]);  //j is the %5
                        PetrelLogger.InfoOutputWindow("the name");
                        PetrelLogger.InfoOutputWindow(arr1[num_bh * 5 + 0].ToString());
                        bh.Name = arr1[num_bh * 5 + 0].ToString();
                        PetrelLogger.InfoOutputWindow(num_bh.ToString());
                        dou = Convert.ToDouble(arr1[num_bh * 5 + 1]);
                        bh.PropertyAccess.SetPropertyValue(wellheadX, dou);  //j is the %5
                        //double.Parse('23');
                        Convert.ToDouble(214);
                        PetrelLogger.InfoOutputWindow(dou.ToString());
                        dou = Convert.ToDouble(arr1[num_bh * 5 + 2]);
                        bh.PropertyAccess.SetPropertyValue(wellheadY, dou);  //j is the %5
                                                                             //bh.PropertyAccess.SetPropertyValue(Offset, arr1[num_bh * 5 + 3]);  //j is the %5
                        dou = Convert.ToDouble(arr1[num_bh * 5 + 4]);
                        bh.PropertyAccess.SetPropertyValue(Offset, dou);  //j is the %5
                        tr.Commit();
                    p.ProgressStatus += step;
                    //progressBar1.Value += progressBar1.Step;//让进度条增加一次                               10+N
                    PetrelLogger.InfoOutputWindow(arr1[num_bh * 5 + 1]);//x
                        PetrelLogger.InfoOutputWindow(arr1[num_bh * 5 + 2]);//y
                        PetrelLogger.InfoOutputWindow(arr1[num_bh * 5 + 3]);//KB
                        PetrelLogger.InfoOutputWindow(arr1[num_bh * 5 + 4]);//35
                    }
                p.ProgressStatus += step * 3;
                //progressBar1.Value += progressBar1.Step*3;//让进度条增加一次                               10+4N
                if(p.ProgressStatus == 10+4*N)
                {
                    p.ProgressStatus--;
                    p.ProgressStatus--;
                    p.ProgressStatus--;
                }
                

                                                        /*  using (ITransaction t = DataManager.NewTransaction())
                                                          {
                                                              t.Lock(bh);
                                                              bh.Name = "123";
                                                              //bh.Extensions.Add(xyzColl);
                                                              t.Commit();
                                                          }*/
                num_bh++;
                }
            p.ProgressStatus = max;
            //progressBar1.Value = progressBar1.Maximum;
            p.Dispose();
            PetrelLogger.InfoOutputWindow("changed ok ");
            //(MessageBox.Show("Data changed successfully", "Sure", MessageBoxButtons.YesNo);
            MessageBox.Show("Data changed successfully");
            #region
            /* int column_tem = 0;

                 for(int i = 0; i < count / 5; i++)
                 {
                 //arr1[0]->arr1[count-1] 

                 //遍历所有的arr1的元素
                 //判断是行数
                 int row_tem = i / 5;
                 //判断是列数
                 column_tem = i % 5;
                 //判断第几列
                 #region
                 if (column_tem == 0)
                 {
                     //更改输入name 
                 }
                 else if (column_tem == 1)
                 {
                     //更改输入X
                 }
                 else if (column_tem == 2)
                 {
                     //更改输入Y

                 }
                 else if (column_tem == 3)
                 {
                     //更改输入KB
                 }
                 else if (column_tem == 4)
                 {
                     //更改输入KB value

                 }

                 #endregion
                 }*/
            //修改数据代码
            //.arguments = arguments;
            //this.context = context;

            //arr1[0]->arr1[count-1] 
            #endregion
            #endregion
            //progressBar1.Visible = false;
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // already 1 image in the imageList1 instance
            imageList1.Images.Add(Slb.Ocean.Petrel.UI.PetrelImages.WellBlue);
            imageList1.Images.Add(Slb.Ocean.Petrel.UI.PetrelImages.WellBlue);
            tabPage1.ImageIndex = 1;
            button4.BackgroundImage = PetrelImages.Apply;
            button4.BackgroundImageLayout = ImageLayout.None;
            button3.BackgroundImage = PetrelImages.Editor;
            button3.BackgroundImageLayout = ImageLayout.None;
            button1.BackgroundImage = PetrelImages.Comments;
            button1.BackgroundImageLayout = ImageLayout.None;
            button2.BackgroundImage = PetrelImages.BulletGreen;
            button2.BackgroundImageLayout = ImageLayout.None;






            // second thing in the ImageList -> Ball
            //tabPage2.ImageIndex = 2;
            // third thing in the ImageList -> Info
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Cursor c = System.Windows.Forms.Cursors.Cross;
            IProgress p = PetrelLogger.NewProgress(0, 100,
            ProgressType.Cancelable, c);
            for (int i = 0; i <= 100; i += 10)
            {
                p.ProgressStatus = i;
                if (p.IsCanceled)
                {
                    // clean up any resources or data
                    break;
                }
                // operation step ...
            }
            p.Dispose();
            /*
            Cursor c = System.Windows.Forms.Cursors.Cross;
            IProgress p = PetrelLogger.NewProgress(0, 100,
            ProgressType.Cancelable, c);
           
                p.ProgressStatus = i;
                if (p.IsCanceled)
                {
                }
               
            p.Dispose();
            */
        }

        private void dropTarget1_Click(object sender, EventArgs e)
        {
            Slb.Ocean.Petrel.DomainObject.Project proj = PetrelProject.PrimaryProject;
            // get the root of all domain objects 
            WellRoot wr = WellRoot.Get(proj);
            BoreholeCollection bc = wr.BoreholeCollection;
            // SeismicRoot root = SeismicRoot.Get(proj);
            //SeismicProject g = root.SeismicProject;
            if (bc == null) return;
            INameInfoFactory nameFact;
            nameFact = CoreSystem.GetService<INameInfoFactory>(bc);
            NameInfo nameInfo = nameFact.GetNameInfo(bc);
            presentationBox1.Text = nameInfo.Name;
            IImageInfoFactory imageFact;
            imageFact = CoreSystem.GetService<IImageInfoFactory>(bc);
            ImageInfo imageInfo = imageFact.GetImageInfo(bc);
            presentationBox1.Image = imageInfo.TypeImage;
            // use Tag later to get object from PresentationBox
            presentationBox1.Tag = bc;
        }

        private void toolTipPanel1_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
