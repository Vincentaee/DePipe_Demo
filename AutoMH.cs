using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.Creation;
using ExReaderConsole;
using Autodesk.Revit.DB.Architecture;

namespace Auto_ManHole
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class Handler_AutoMH : IExternalEventHandler
    {
        public string file_path
        {
            get;
            set;
        }
        public IList<double> xyz_shift
        {
            get;
            set;
        }
    
        public void Execute(UIApplication app)
        {
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ExReader mhx = new ExReader();
            mhx.SetData(file_path, 1);
            mhx.PassMH();
            mhx.CloseEx();

            Transaction trans = new Transaction(doc);
            trans.Start("交易開始");
            
            FilteredElementCollector collector1 = new FilteredElementCollector(doc);
            ICollection<Element> collection = collector1.WhereElementIsNotElementType().ToElements();
            ICollection<FamilySymbol> fsym = null;
            fsym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            //若沒有事先建立過的族群將會無法建置
            
            //初始偏移XYZ
            double xshift = xyz_shift[0];
            double yshift = xyz_shift[1];
            double zshift = xyz_shift[2];

            //訂定土體範圍            
            List<XYZ> topxyz = new List<XYZ>();
            //迴圈放置人孔和手孔
            foreach (List<string> rows in mhx.MHdata)
            {
                //位置
                XYZ pos = new XYZ(double.Parse(rows[5]) - xshift, double.Parse(rows[6]) - yshift, double.Parse(rows[7]) - zshift);
                //敷地範圍給定
                topxyz.Add(pos * 1000 / 304.8);
               
                //判斷類型: 方人孔 & 圓人孔 & 方手孔 & 圓手孔

                if (rows[3] == "人孔")
                {
                    if (rows[4] == "方形")
                    {
                        foreach (FamilySymbol symbq in fsym)
                        {
                            if (symbq.Name == "方人孔")
                            {
                                //建立實體寫入人孔參數
                                symbq.Activate();
                                FamilyInstance instance = doc.Create.NewFamilyInstance(pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                                instance.LookupParameter("類別碼").Set(rows[0]);
                                instance.LookupParameter("編號").Set(rows[1]);
                                instance.LookupParameter("類型").Set(rows[2]);
                                instance.LookupParameter("型式").Set(rows[3]);
                                instance.LookupParameter("方位角").SetValueString((Double.Parse(rows[8])).ToString());
                                instance.LookupParameter("箱左長").SetValueString((Double.Parse(rows[21]) * 1000).ToString());
                                instance.LookupParameter("箱下寬").SetValueString((Double.Parse(rows[23]) * 1000).ToString());
                                instance.LookupParameter("箱右長").SetValueString((Double.Parse(rows[22]) * 1000).ToString());
                                instance.LookupParameter("箱上寬").SetValueString((Double.Parse(rows[24]) * 1000).ToString());
                                try { instance.LookupParameter("頸部深度").SetValueString((Double.Parse(rows[18]) * 1000).ToString()); } catch { TaskDialog.Show("message", "方人孔:頸部深度"); }
                                instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString());
                                instance.LookupParameter("附註").Set(rows[27]);

                                try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "方人孔"); };
                                instance.LookupParameter("圓蓋直徑").SetValueString((Double.Parse(rows[12]) * 1000).ToString());
                                instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString());
                            }
                        }

                    }
                    else
                    {
                        foreach (FamilySymbol symbq in fsym)
                        {
                            if (symbq.Name == "圓人孔")
                            {
                                //建立實體寫入人孔參數
                                symbq.Activate();
                                FamilyInstance instance = doc.Create.NewFamilyInstance(pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                                instance.LookupParameter("類別碼").Set(rows[0]);
                                instance.LookupParameter("編號").Set(rows[1]);
                                try { instance.LookupParameter("類型").Set(rows[2]); } catch { TaskDialog.Show("message", "類型"); }
                                try { instance.LookupParameter("型式").Set(rows[3]); } catch { TaskDialog.Show("message", "型式"); }

                                try { instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體深度"); }
                                try { instance.LookupParameter("附註").Set(rows[27]); } catch { TaskDialog.Show("message", "附註"); }

                                try { instance.LookupParameter("圓蓋直徑").SetValueString((Double.Parse(rows[12]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓蓋直徑"); }
                                try { instance.LookupParameter("頸部深度").SetValueString((Double.Parse(rows[18]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓人孔:頸部深度"); }
                                try { instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString()); } catch { TaskDialog.Show("message", "壁厚"); }
                                try { instance.LookupParameter("箱體直徑").SetValueString((Double.Parse(rows[20]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體直徑"); }

                                try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "圓人孔"); };
                            }
                        }


                    }

                }

                else if (rows[3] == "陰井")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "陰井")
                        {
                            //建立實體寫入人孔參數
                            symbq.Activate();
                            FamilyInstance instance = doc.Create.NewFamilyInstance(pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            instance.LookupParameter("類型").Set(rows[2]);
                            instance.LookupParameter("型式").Set(rows[3]);
                            instance.LookupParameter("方位角").SetValueString((Double.Parse(rows[8])).ToString());
                            instance.LookupParameter("長度").SetValueString((Double.Parse(rows[13]) * 1000).ToString());
                            instance.LookupParameter("寬度").SetValueString((Double.Parse(rows[14]) * 1000).ToString());
                            instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString());
                            instance.LookupParameter("附註").Set(rows[27]);

                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "陰井"); };
                            instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString());
                        }
                    }
                }
                else if (rows[3] == "電力電桿")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "電力電桿")
                        {
                            //建立實體寫入人孔參數
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 9);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            try { instance.LookupParameter("類型").Set(rows[2]); } catch { TaskDialog.Show("message", "類型"); }
                            try { instance.LookupParameter("型式").Set(rows[3]); } catch { TaskDialog.Show("message", "型式"); }

                            try { instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體深度"); }
                            try { instance.LookupParameter("附註").Set(rows[27]); } catch { TaskDialog.Show("message", "附註"); }

                            try { instance.LookupParameter("圓蓋直徑").SetValueString((Double.Parse(rows[12]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓蓋直徑"); }
                            try { instance.LookupParameter("頸部深度").SetValueString((Double.Parse(rows[18]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓人孔:頸部深度"); }
                            try { instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString()); } catch { TaskDialog.Show("message", "壁厚"); }
                            try { instance.LookupParameter("箱體直徑").SetValueString((Double.Parse(rows[20]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體直徑"); }

                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "圓人孔"); };
                        }
                    }
                }
                else if (rows[3] == "電塔")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "電塔")
                        {
                            //建立實體寫入人孔參數
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 25);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            try { instance.LookupParameter("類型").Set(rows[2]); } catch { TaskDialog.Show("message", "類型"); }
                            try { instance.LookupParameter("型式").Set(rows[3]); } catch { TaskDialog.Show("message", "型式"); }

                            try { instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體深度"); }
                            try { instance.LookupParameter("附註").Set(rows[27]); } catch { TaskDialog.Show("message", "附註"); }

                            try { instance.LookupParameter("圓蓋直徑").SetValueString((Double.Parse(rows[12]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓蓋直徑"); }
                            try { instance.LookupParameter("頸部深度").SetValueString((Double.Parse(rows[18]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓人孔:頸部深度"); }
                            try { instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString()); } catch { TaskDialog.Show("message", "壁厚"); }
                            try { instance.LookupParameter("箱體直徑").SetValueString((Double.Parse(rows[20]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體直徑"); }

                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "圓人孔"); };
                        }
                    }
                }
                else if (rows[3] == "電信電桿")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "電信電桿")
                        {
                            //建立實體寫入人孔參數
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 5);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            try { instance.LookupParameter("類型").Set(rows[2]); } catch { TaskDialog.Show("message", "類型"); }
                            try { instance.LookupParameter("型式").Set(rows[3]); } catch { TaskDialog.Show("message", "型式"); }

                            try { instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體深度"); }
                            try { instance.LookupParameter("附註").Set(rows[27]); } catch { TaskDialog.Show("message", "附註"); }

                            try { instance.LookupParameter("圓蓋直徑").SetValueString((Double.Parse(rows[12]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓蓋直徑"); }
                            try { instance.LookupParameter("頸部深度").SetValueString((Double.Parse(rows[18]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓人孔:頸部深度"); }
                            try { instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString()); } catch { TaskDialog.Show("message", "壁厚"); }
                            try { instance.LookupParameter("箱體直徑").SetValueString((Double.Parse(rows[20]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體直徑"); }

                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "圓人孔"); };
                        }
                    }
                }
                else if (rows[0] == "8020190" && rows[2] == "電力")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "變電箱")
                        {
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 1.2);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            instance.LookupParameter("類型").Set(rows[2]);
                            instance.LookupParameter("型式").Set(rows[3]);
                            instance.LookupParameter("方位角").SetValueString((Double.Parse(rows[8])).ToString());
                            instance.LookupParameter("長度").SetValueString((Double.Parse(rows[13]) * 1000).ToString());
                            instance.LookupParameter("寬度").SetValueString((Double.Parse(rows[14]) * 1000).ToString());
                            instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString());
                            instance.LookupParameter("附註").Set(rows[27]);

                            instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString());
                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "方手孔"); };
                        }
                    }
                }
                else if (rows[0] == "8010190" && rows[2] == "電信")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "電信箱")
                        {
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 1.4);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            instance.LookupParameter("類型").Set(rows[2]);
                            instance.LookupParameter("型式").Set(rows[3]);
                            instance.LookupParameter("方位角").SetValueString((Double.Parse(rows[8])).ToString());
                            instance.LookupParameter("長度").SetValueString((Double.Parse(rows[13]) * 1000).ToString());
                            instance.LookupParameter("寬度").SetValueString((Double.Parse(rows[14]) * 1000).ToString());
                            instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString());
                            instance.LookupParameter("附註").Set(rows[27]);

                            instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString());
                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "方手孔"); };
                        }
                    }
                }
                else if (rows[3] == "地上消防栓")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "消防栓")
                        {
                            //建立實體寫入人孔參數
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 0.65);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            try { instance.LookupParameter("類型").Set(rows[2]); } catch { TaskDialog.Show("message", "類型"); }
                            try { instance.LookupParameter("型式").Set(rows[3]); } catch { TaskDialog.Show("message", "型式"); }

                            try { instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體深度"); }
                            try { instance.LookupParameter("附註").Set(rows[27]); } catch { TaskDialog.Show("message", "附註"); }

                            try { instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString()); } catch { TaskDialog.Show("message", "壁厚"); }

                            try { instance.LookupParameter("圓蓋直徑").SetValueString((Double.Parse(rows[12]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓蓋直徑"); }
                            try { instance.LookupParameter("頸部深度").SetValueString((Double.Parse(rows[18]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓人孔:頸部深度"); }
                            try { instance.LookupParameter("箱體直徑").SetValueString((Double.Parse(rows[20]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體直徑"); }

                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "圓人孔"); };
                        }
                    }
                }
                else if (rows[3] == "路燈電桿")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "路燈電桿")
                        {
                            //建立實體寫入人孔參數
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 5.3);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            try { instance.LookupParameter("類型").Set(rows[2]); } catch { TaskDialog.Show("message", "類型"); }
                            try { instance.LookupParameter("型式").Set(rows[3]); } catch { TaskDialog.Show("message", "型式"); }

                            try { instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體深度"); }
                            try { instance.LookupParameter("附註").Set(rows[27]); } catch { TaskDialog.Show("message", "附註"); }

                            try { instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString()); } catch { TaskDialog.Show("message", "壁厚"); }

                            try { instance.LookupParameter("圓蓋直徑").SetValueString((Double.Parse(rows[12]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓蓋直徑"); }
                            try { instance.LookupParameter("頸部深度").SetValueString((Double.Parse(rows[18]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓人孔:頸部深度"); }
                            try { instance.LookupParameter("箱體直徑").SetValueString((Double.Parse(rows[20]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體直徑"); }

                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "圓人孔"); };
                        }
                    }
                }
                else if (rows[3] == "號誌電桿")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "號誌電桿")
                        {
                            //建立實體寫入人孔參數
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 2.3);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            try { instance.LookupParameter("類型").Set(rows[2]); } catch { TaskDialog.Show("message", "類型"); }
                            try { instance.LookupParameter("型式").Set(rows[3]); } catch { TaskDialog.Show("message", "型式"); }

                            try { instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體深度"); }
                            try { instance.LookupParameter("附註").Set(rows[27]); } catch { TaskDialog.Show("message", "附註"); }

                            try { instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString()); } catch { TaskDialog.Show("message", "壁厚"); }

                            try { instance.LookupParameter("圓蓋直徑").SetValueString((Double.Parse(rows[12]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓蓋直徑"); }
                            try { instance.LookupParameter("頸部深度").SetValueString((Double.Parse(rows[18]) * 1000).ToString()); } catch { TaskDialog.Show("message", "圓人孔:頸部深度"); }
                            try { instance.LookupParameter("箱體直徑").SetValueString((Double.Parse(rows[20]) * 1000).ToString()); } catch { TaskDialog.Show("message", "箱體直徑"); }

                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "圓人孔"); };
                        }
                    }
                }
                else if (rows[3] == "地上設備" && rows[2] == "電信")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "號誌設備")
                        {
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 1.08);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            instance.LookupParameter("類型").Set(rows[2]);
                            instance.LookupParameter("型式").Set(rows[3]);
                            instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString());
                            instance.LookupParameter("方位角").SetValueString((Double.Parse(rows[8])).ToString());
                            instance.LookupParameter("長度").SetValueString((Double.Parse(rows[13]) * 1000).ToString());
                            instance.LookupParameter("寬度").SetValueString((Double.Parse(rows[14]) * 1000).ToString());
                            instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString());
                            instance.LookupParameter("附註").Set(rows[27]);

                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "方手孔"); };
                        }
                    }
                }
                else if (rows[3] == "路燈開關")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "電力開關")
                        {
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 0.2);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);
                            
                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            instance.LookupParameter("類型").Set(rows[2]);
                            instance.LookupParameter("型式").Set(rows[3]);
                            instance.LookupParameter("方位角").SetValueString((Double.Parse(rows[8])).ToString());
                            instance.LookupParameter("長度").SetValueString((Double.Parse(rows[13]) * 1000).ToString());
                            instance.LookupParameter("寬度").SetValueString((Double.Parse(rows[14]) * 1000).ToString());
                            instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString());
                            instance.LookupParameter("附註").Set(rows[27]);

                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "方手孔"); };
                        }
                    }
                }
                else if (rows[3] == "號誌開關")
                {
                    foreach (FamilySymbol symbq in fsym)
                    {
                        if (symbq.Name == "號誌開關")
                        {
                            //建立實體寫入人孔參數
                            symbq.Activate();
                            XYZ new_pos = new XYZ(pos.X, pos.Y, pos.Z + 0.2);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(new_pos * 1000 / 304.8, symbq, StructuralType.NonStructural);
                            
                            instance.LookupParameter("類別碼").Set(rows[0]);
                            instance.LookupParameter("編號").Set(rows[1]);
                            instance.LookupParameter("類型").Set(rows[2]);
                            instance.LookupParameter("型式").Set(rows[3]);
                            instance.LookupParameter("方位角").SetValueString((Double.Parse(rows[8])).ToString());
                            instance.LookupParameter("長度").SetValueString((Double.Parse(rows[13]) * 1000).ToString());
                            instance.LookupParameter("寬度").SetValueString((Double.Parse(rows[14]) * 1000).ToString());
                            instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString());
                            instance.LookupParameter("附註").Set(rows[27]);

                            instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString());
                            try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "方手孔"); };
                        }
                    }
                }
                else
                {
                    if (rows[4] == "方形")
                    {
                        foreach (FamilySymbol symbq in fsym)
                        {
                            if (symbq.Name == "方手孔")
                            {
                                //建立實體寫入人孔參數
                                symbq.Activate();
                                FamilyInstance instance = doc.Create.NewFamilyInstance(pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                                instance.LookupParameter("類別碼").Set(rows[0]);
                                instance.LookupParameter("編號").Set(rows[1]);
                                instance.LookupParameter("類型").Set(rows[2]);
                                instance.LookupParameter("型式").Set(rows[3]);
                                instance.LookupParameter("方位角").SetValueString((Double.Parse(rows[8])).ToString());
                                instance.LookupParameter("長度").SetValueString((Double.Parse(rows[13]) * 1000).ToString());
                                instance.LookupParameter("寬度").SetValueString((Double.Parse(rows[14]) * 1000).ToString());
                                instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString());
                                instance.LookupParameter("附註").Set(rows[27]);

                                instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString());
                                try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "方手孔"); };
                            }
                        }


                    }
                    else
                    {
                        foreach (FamilySymbol symbq in fsym)
                        {
                            if (symbq.Name == "圓手孔")
                            {
                                //建立實體寫入人孔參數
                                symbq.Activate();
                                FamilyInstance instance = doc.Create.NewFamilyInstance(pos * 1000 / 304.8, symbq, StructuralType.NonStructural);

                                instance.LookupParameter("類別碼").Set(rows[0]);
                                instance.LookupParameter("編號").Set(rows[1]);
                                instance.LookupParameter("類型").Set(rows[2]);
                                instance.LookupParameter("型式").Set(rows[3]);
                                instance.LookupParameter("箱體深度").SetValueString((Double.Parse(rows[25]) * 1000).ToString());
                                instance.LookupParameter("箱體直徑").SetValueString((Double.Parse(rows[20]) * 1000).ToString());
                                instance.LookupParameter("圓蓋直徑").SetValueString((Double.Parse(rows[12]) * 1000).ToString());
                                instance.LookupParameter("附註").Set(rows[27]);

                                instance.LookupParameter("壁厚").SetValueString((Double.Parse(rows[19]) * 1000).ToString());
                                try { instance.LookupParameter("形狀").Set(rows[4]); } catch { TaskDialog.Show("message", "圓手孔"); };
                            }
                        }
                    }
                }
            }
            /*
            if(topxyz.Count >= 3)
            {
                TopographySurface.Create(doc, topxyz);
            }
            */

            TaskDialog.Show("Done", "人手孔建置完畢");
            trans.Commit();
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}



