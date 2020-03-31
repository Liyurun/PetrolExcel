using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Slb.Ocean.Petrel.Commands;
using Slb.Ocean.Petrel;
using System.Windows.Forms;
using WindowsFormsApplication1;
using System.IO;
using Slb.Ocean.Petrel.DomainObject.Well;
using Slb.Ocean.Petrel.DomainObject.Basics;
using static OceanReadingData.read;
using Slb.Ocean.Petrel.Workflow;
using Slb.Ocean.Core;

namespace OceanReadingData
{
    static class Gloable
    {
        public static int if_change = 0;
        public static string txt_path = "txt_path";
        public static string excel_path = "excel_path";
    }
    class openexcel : SimpleCommandHandler
    {
        
        public static string ID = "OceanReadingData.NewCommand";

        #region SimpleCommandHandler Members

        public override bool CanExecute(Slb.Ocean.Petrel.Contexts.Context context)
        { 
            return true;
        }

        public override void Execute(Slb.Ocean.Petrel.Contexts.Context context)
        {

            //open the excel      
            //TODO: Add command execution logic here


            if (Gloable.if_change == 1)
            {
                Gloable.if_change = 0;
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
                PetrelLogger.InfoOutputWindow("changed in openexcel");
                double dou = Convert.ToDouble(arr1[num_bh+1]);
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

                        PetrelLogger.InfoOutputWindow(arr1[num_bh * 5 + 1]);//x
                        PetrelLogger.InfoOutputWindow(arr1[num_bh * 5 + 2]);//y
                        PetrelLogger.InfoOutputWindow(arr1[num_bh * 5 + 3]);//KB
                        PetrelLogger.InfoOutputWindow(arr1[num_bh * 5 + 4]);//35
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
                PetrelLogger.InfoOutputWindow("changed ok ");
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

            }
            else
            {
                PetrelLogger.InfoOutputWindow("save the data");
                //save the data
                Gloable.if_change = 1;
                //open excel 
                PetrelLogger.InfoOutputWindow(string.Format("{0} clicked", @"openEXCEL"));
                if (MessageBox.Show("EXCEL will open", "Sure", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    System.Diagnostics.Process p = new System.Diagnostics.Process();
                    p.StartInfo.UseShellExecute = true;
                    p.StartInfo.FileName = "D:\\openexcel.xlsx";
                    p.Start();
                }
            }
            
        }

        #endregion
    }
}
