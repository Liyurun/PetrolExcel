using System;

using Slb.Ocean.Core;
using Slb.Ocean.Petrel;
using Slb.Ocean.Petrel.UI;
using Slb.Ocean.Petrel.Workflow;
using Slb.Ocean.Petrel.DomainObject.Well;
using Slb.Ocean.Petrel.DomainObject;
using System.Collections.Generic;
using Slb.Ocean.Units;
using Slb.Ocean.Geometry;
using System.Collections;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Slb.Ocean.Petrel.DomainObject.Basics;
using Slb.Ocean.Petrel.Simulation.EclipseKeywords;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApplication1;

namespace WindowsFormsApplication1
{
    static class Gloable
    {
        public static int if_change = 0;
        public static string txt_path = "txt_path";
        public static string excel_path = "excel_path";
    }
}

namespace OceanReadingData
{
    /// <summary>
    /// This class contains all the methods and subclasses of the read.
    /// Worksteps are displayed in the workflow editor.
    /// </summary>
    class read : Workstep<read.Arguments>, IExecutorSource, IAppearance, IDescriptionSource
    {
        #region Overridden Workstep methods

        /// <summary>
        /// Creates an empty Argument instance
        /// </summary>
        /// <returns>New Argument instance.</returns>

        protected override read.Arguments CreateArgumentPackageCore(IDataSourceManager dataSourceManager)
        {
            return new Arguments(dataSourceManager);
        }
        /// <summary>
        /// Copies the Arguments instance.
        /// </summary>
        /// <param name="fromArgumentPackage">the source Arguments instance</param>
        /// <param name="toArgumentPackage">the target Arguments instance</param>
        protected override void CopyArgumentPackageCore(Arguments fromArgumentPackage, Arguments toArgumentPackage)
        {
            DescribedArgumentsHelper.Copy(fromArgumentPackage, toArgumentPackage);
        }

        /// <summary>
        /// Gets the unique identifier for this Workstep.
        /// </summary>
        protected override string UniqueIdCore
        {
            get
            {
                return "509d32f4-e781-44c4-ba87-98623d13e5fc";
            }
        }
        #endregion

        #region IExecutorSource Members and Executor class

        /// <summary>
        /// Creates the Executor instance for this workstep. This class will do the work of the Workstep.
        /// </summary>
        /// <param name="argumentPackage">the argumentpackage to pass to the Executor</param>
        /// <param name="workflowRuntimeContext">the context to pass to the Executor</param>
        /// <returns>The Executor instance.</returns>
        public Slb.Ocean.Petrel.Workflow.Executor GetExecutor(object argumentPackage, WorkflowRuntimeContext workflowRuntimeContext)
        {
            return new Executor(argumentPackage as Arguments, workflowRuntimeContext);
        }

        public class Executor : Slb.Ocean.Petrel.Workflow.Executor
        {
            Arguments arguments;
            WorkflowRuntimeContext context;
           
            public Executor(Arguments arguments, WorkflowRuntimeContext context)
            {
                this.arguments = arguments;
                this.context = context;
                                           
                BoreholeCollection bc = this.arguments.Bore_input;

                IList<ReferenceLevel> referenceLevels = new List<ReferenceLevel>
                {
                    new ReferenceLevel("KB", 25.0, "Kelly bushing"),
                    new ReferenceLevel("RF", 10.0, "Rig floor")
                };

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
               

                #region
                //excel 路径创建文件
                try
                {
                    string fileTest = "D:\\openexcel.xlsx";
                    if (File.Exists(fileTest))
                    {
                        File.Delete(fileTest);
                    }
                    Excel.Application oApp;
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
                #endregion
                //fs.Close();

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("D:\\openexcel.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet ws = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

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
                PetrelLogger.InfoOutputWindow("properties name");
                foreach (BoreholeProperty bpp in bpc.Properties)
                {
                    PetrelLogger.InfoOutputWindow(bpp.Name.ToString());
                }
                PetrelLogger.InfoOutputWindow("properties PropertyType");
                foreach (BoreholeProperty bpp in bpc.Properties)
                {
                    PetrelLogger.InfoOutputWindow(bpp.ToString());
                }
                PetrelLogger.InfoOutputWindow("properties iswirte");
                foreach (BoreholeProperty bpp in bpc.Properties)
                {
                    PetrelLogger.InfoOutputWindow(bpp.IsWritable.ToString());
                }
                PetrelLogger.InfoOutputWindow("properties Template");
                foreach (BoreholeProperty bpp in bpc.Properties)
                {
                    PetrelLogger.InfoOutputWindow(bpp.Template.ToString());
                }
                PetrelLogger.InfoOutputWindow("properties DATATYPE");
                foreach (BoreholeProperty bpp in bpc.Properties)
                {
                    PetrelLogger.InfoOutputWindow(bpp.DataType.ToString());
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


                    //bh_refer is the 
                    PetrelLogger.InfoOutputWindow("print each_bh_reder in bh_refer");

                    //print before changing
                    PetrelLogger.InfoOutputWindow("be data start to change");
                    foreach (ReferenceLevel ea_bh_refer in bh_refer)
                    {
                        PetrelLogger.InfoOutputWindow(ea_bh_refer.ToString());
                    }
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
                }
                           fs.Flush();
                           fs.Close();
                           //写入excel

                           // 新建一个DataTable
                           //DataTable tb = new DataTable();
                           // 添加一列用于存放读入的浮点数
                           // DataColumn c = tb.Columns.Add("Value", typeof(double));

                           // 打开文件准备读取数据   
                           // 打开文件准备读取数据   
                           StreamReader rd = File.OpenText(@"d:\ak.txt");
                           string restOfStream = rd.ReadToEnd();
                           rd.Close();
                           //输出DataTable中保存的数组
                           //foreach (DataRow r in tb.Rows)
                           //Console.WriteLine("assa{0}",restOfStream);
                           string[] arr1 = restOfStream.Split('(', ')', ',', '\n', '[', ']');
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

                               }
                           }
                           finally { }

                           sheet.Close(true, Type.Missing, Type.Missing);
                           excel.Quit();
                           excel = null;
                           GC.Collect();//垃圾回收
                                        //清空缓冲区、关闭流

                // 提示用户：文件保存的位置和文件名
                MessageBox.Show("文件已成功保存到" + "D:\\openexcel.xlsx");
                PetrelLogger.InfoOutputWindow("文件已成功保存到" + "D:\\openexcel.xlsx");







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


            }


            public override void ExecuteSimple()
            {
                // TODO: Implement the workstep logic here.
                //get current primariy project
                /*Project proj = PetrelProject.PrimaryProject;
                // get the root of all domain objects 
                WellRoot root = WellRoot.Get(proj);
                BoreholeCollection wells = root.BoreholeCollection;
                WellLog poro = WellLog.NullObject;
                Template template = PetrelProject.WellKnownTemplates.PetrophysicalGroup.Porosity;
                // ITemplates template = PetrelProject.WellKnownTemplates.PetrophysicalGroup.Porosity;
                IUnitMeasurement um = template.UnitMeasurement;
                foreach(BoreholeCollection bhc in wells.BoreholeCollections)
                {
                    foreach (Borehole bh in bhc)
                    {
                        foreach(WellLog l in bh.Logs.WellLogs)
                        {
                      
                                poro = l;
                            
                        }
                        if (!poro.IsGood) break;
                    }
                    if (!poro.IsGood) break;
                }
                if (poro == WellLog.NullObject) return;
                PetrelLogger.InfoOutputWindow("Found log " + poro.Name);
                */
            }
        }

        private static void SetReferenceLevels(Borehole bbb, IList<ReferenceLevel> referenceLevels)
        {
            throw new NotImplementedException();
        }

        private static BoreholeCollection GetOrCreateBoreholeCollection(BoreholeCollection boreholeCollection, string v)
        {
            throw new NotImplementedException();
        }

        private static BoreholeCollection GetOrCreateBoreholeCollection()
        {
            throw new NotImplementedException();
        }

        #endregion

        /// <summary>
        /// ArgumentPackage class for read.
        /// Each public property is an argument in the package.  The name, type and
        /// input/output role are taken from the property and modified by any
        /// attributes applied.
        /// </summary>
        public class Arguments : DescribedArgumentsByReflection
        {
            public Arguments()
                : this(DataManager.DataSourceManager)
            {                
            }

            public Arguments(IDataSourceManager dataSourceManager)
            {
            }

            private Slb.Ocean.Petrel.DomainObject.Well.BoreholeCollection bore_input;
            private Slb.Ocean.Petrel.DomainObject.Well.BoreholeCollection bore_output;

            public Slb.Ocean.Petrel.DomainObject.Well.BoreholeCollection Bore_input
            {
                internal get { return this.bore_input; }
                set { this.bore_input = value; }
            }

            public Slb.Ocean.Petrel.DomainObject.Well.BoreholeCollection Bore_output
            {
                get { return this.bore_output; }
                internal set { this.bore_output = value; }
            }


        }
    
        #region IAppearance Members
        public event EventHandler<TextChangedEventArgs> TextChanged;
        protected void RaiseTextChanged()
        {
            if (this.TextChanged != null)
                this.TextChanged(this, new TextChangedEventArgs(this));
        }

        public string Text
        {
            get { return Description.Name; }
            private set 
            {
                // TODO: implement set
                this.RaiseTextChanged();
            }
        }

        public event EventHandler<ImageChangedEventArgs> ImageChanged;
        protected void RaiseImageChanged()
        {
            if (this.ImageChanged != null)
                this.ImageChanged(this, new ImageChangedEventArgs(this));
        }

        public System.Drawing.Bitmap Image
        {
            get { return PetrelImages.Modules; }
            private set 
            {
                // TODO: implement set
                this.RaiseImageChanged();
            }
        }
        #endregion

        #region IDescriptionSource Members

        /// <summary>
        /// Gets the description of the read
        /// </summary>
        public IDescription Description
        {
            get { return readDescription.Instance; }
        }

        /// <summary>
        /// This singleton class contains the description of the read.
        /// Contains Name, Shorter description and detailed description.
        /// </summary>
        public class readDescription : IDescription
        {
            /// <summary>
            /// Contains the singleton instance.
            /// </summary>
            private  static readDescription instance = new readDescription();
            /// <summary>
            /// Gets the singleton instance of this Description class
            /// </summary>
            public static readDescription Instance
            {
                get { return instance; }
            }

            #region IDescription Members

            /// <summary>
            /// Gets the name of read
            /// </summary>
            public string Name
            {
                get { return "read"; }
            }
            /// <summary>
            /// Gets the short description of read
            /// </summary>
            public string ShortDescription
            {
                get { return ""; }
            }
            /// <summary>
            /// Gets the detailed description of read
            /// </summary>
            public string Description
            {
                get { return ""; }
            }

            #endregion
        }
        #endregion


    }
}