using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.IO;
using ExReaderConsole;
using Autodesk.Revit.DB.Structure;
using System.Runtime.InteropServices;

namespace PipeV2
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class Handler_autoPipe : IExternalEventHandler
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
            try
            {
                Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
                UIDocument uidoc = new UIDocument(document);
                Document doc = uidoc.Document;

                Application app_ = doc.Application;
                UIApplication uiapp = new UIApplication(app_);

                var filename = doc.PathName;
                //excel
                ExReader dan = new ExReader();
                ExReader mhx = new ExReader();
                //dan.SetData(file_path, 1);
                //dan.PassMH();
                //dan.CloseEx();

                mhx.SetData(file_path, 2);
                mhx.PassMH();
                mhx.CloseEx();
                //處理座標
                List<List<XYZ>> pos_list = new List<List<XYZ>>();

                //存xyz string
                List<string> pos_data = new List<string>();
                foreach (List<string> rows in mhx.MHdata)
                {
                    pos_data.Add(rows[7]);
                }
                //處理string
                foreach (string pos_string in pos_data)
                {
                    string[] pos_row = pos_string.Replace(";", ",").Split(',');
                    string tx = "", ty = "", tz = "";
                    List<XYZ> pos_Row = new List<XYZ>();

                    //刪除重複
                    for (int i = 0; i != pos_row.Length; i = i + 3)
                    {
                        string x = pos_row[i], y = pos_row[i + 1], z = pos_row[i + 2];
                        if (x == tx && y == ty && z == tz || x == "")
                            continue;
                        XYZ pos = new XYZ(Double.Parse(x), Double.Parse(y), Double.Parse(z));
                        pos_Row.Add(pos);

                        tx = x; ty = y; tz = z;
                    }
                    pos_list.Add(pos_Row);
                }

                double xshift = xyz_shift[0];
                double yshift = xyz_shift[1];
                double zshift = xyz_shift[2];

                //開始Sweep
                Sweep sb_instance = null;
                Sweep sc_instance = null;

                int index = 0;

                //設置警告、錯誤訊息處理
                WarningSwallower warningSwallower = new WarningSwallower();
                warningSwallower.Case_name = file_path;
                warningSwallower.Pipe_number_list = new List<string>();
                //開始建置
                foreach (List<string> rows in mhx.MHdata)
                {
                    //UIDocument edit_uidoc = app.OpenAndActivateDocument(@"E:\2018研發案\01_pipe\family\edit.rfa");
                    //UIDocument edit_uidoc = app.OpenAndActivateDocument(@"C:\中興企劃案\2019版\edit.rfa");
                    UIDocument edit_uidoc = app.OpenAndActivateDocument(@"C:\Users\TsaiWeiLun\OneDrive\Revit_SinoPipe\edit.rfa");

                    Document edit_doc = edit_uidoc.Document;
                    var rvtname = edit_doc.PathName;
                    ICollection<FamilySymbol> familyinstance = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                    ICollection<Family> edit_doc_families = new FilteredElementCollector(edit_doc).OfClass(typeof(Family)).Cast<Family>().ToList();
                    ICollection<Material> materials = new FilteredElementCollector(edit_doc).OfClass(typeof(Material)).Cast<Material>().ToList();

                    ElementId concrete_m_id = (from x in materials where x.Name == "混凝土" select x.Id).First<ElementId>();
                    int count = pos_list[index].Count;

                    using (Transaction t = new Transaction(edit_doc, "Create sphere direct shape"))
                    {
                        t.Start();

                        FamilySymbol proBase = null;
                        FamilySymbol proCore = null;

                        //掃掠
                        try
                        {
                            //斷面選擇
                            if (rows[4] == "圓形")
                            {

                                if (rows[14] == "1" && rows[13] == "1")
                                {
                                    //圓形1x1
                                    proCore = familyinstance.Where(x => x.Name == "normal").First();
                                    proCore.LookupParameter("D").SetValueString((double.Parse(rows[8]) * 1000).ToString());
                                    proCore.LookupParameter("t").SetValueString((double.Parse(rows[11]) * 1000).ToString());

                                }
                                else
                                {
                                    //圓形2x1, 3x1 , ......Nx1
                                    t.Commit();

                                    Family detailArc = new FilteredElementCollector(edit_doc).OfClass(typeof(Family)).Cast<Family>().ToList().Where(x => x.Name == "profile_circle_core_origin_Edit").First();
                                    Document fa_doc = edit_doc.EditFamily(detailArc);
                                    List<string> para_s = new List<string> { "Xn", "Yn", "壁厚", "內徑" };
                                    List<string> parValue = new List<string> { rows[13].ToString(), rows[14].ToString(), (double.Parse(rows[11]) * 1000).ToString(), (double.Parse(rows[12].Split('x')[0]) / 2).ToString(), rows[12].Split('x')[1].ToString() };// Xn, Yn, t, radius, real num

                                    ProduceProfile(ref fa_doc, ref para_s, ref parValue);
                                    DrawCurve_Inner(ref fa_doc, ref parValue);
                                    DrawCurve(ref fa_doc, ref parValue);
                                    fa_doc.LoadFamily(edit_doc, new FamilyOption());
                                    fa_doc.Close(false);

                                    FamilySymbol familyinstance_new = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "profile_circle_core_origin_Edit").First();
                                    proCore = familyinstance_new;

                                    t.Start("Restart");
                                }
                            }
                            else if (rows[4] == "方形")
                            {
                                if (double.Parse(rows[8]) == 0)
                                {
                                    if (rows[13] == "1" && rows[14] == "1")
                                    {
                                        //方形1X1
                                        proCore = familyinstance.Where(x => x.Name == "box").First();
                                        proCore.LookupParameter("B").SetValueString((double.Parse(rows[9]) * 1000).ToString());
                                        proCore.LookupParameter("H").SetValueString((double.Parse(rows[10]) * 1000).ToString());
                                        proCore.LookupParameter("t").SetValueString((double.Parse(rows[11]) * 1000).ToString());

                                    }
                                    else
                                    {
                                        //方形2x1, 3x1 , ......Nx1
                                        t.Commit();

                                        Family detailArc = new FilteredElementCollector(edit_doc).OfClass(typeof(Family)).Cast<Family>().ToList().Where(x => x.Name == "Box_Nx1").First();
                                        Document fa_doc = edit_doc.EditFamily(detailArc);
                                        List<string> para_s = new List<string> { "Xn", "nB", "nH", "t" };
                                        List<string> parValue = new List<string> { rows[13].ToString(), rows[12].Split('x')[0].ToString(), rows[12].Split('x')[1].ToString(), (double.Parse(rows[11]) * 1000).ToString() };// x, y, dia, real num

                                        ProduceProfile(ref fa_doc, ref para_s, ref parValue);
                                        DrawBox(ref fa_doc, ref parValue);
                                        fa_doc.LoadFamily(edit_doc, new FamilyOption());
                                        fa_doc.Close(false);

                                        FamilySymbol familyinstance_new = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "Box_Nx1").First();
                                        proCore = familyinstance_new;

                                        t.Start("Restart");

                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        //方形混凝土包含內管
                                        t.Commit();

                                        int Xn = int.Parse(rows[13]);
                                        int Yn = int.Parse(rows[14]);
                                        double B = double.Parse(rows[9]) * 1000;
                                        double H = double.Parse(rows[10]) * 1000;
                                        double thick = double.Parse(rows[11]) * 1000; //壁厚
                                        double radius = double.Parse(rows[12].Split('x')[0]) / 2; //內徑
                                        int amount = int.Parse(rows[12].Split('x')[1]);

                                        double PB1 = (B - 310 * (Xn - 1)) / 2;
                                        double PB2 = (H - 310 * (Yn - 1)) / 2;

                                        if (PB1 <= radius * 2)
                                        {
                                            PB1 = B / (Xn + 1);
                                        }
                                        if (PB2 <= radius * 2)
                                        {
                                            PB2 = H / (Yn + 1);
                                        }

                                        //core
                                        Family detailArc = new FilteredElementCollector(edit_doc).OfClass(typeof(Family)).Cast<Family>().ToList().Where(x => x.Name == "profile_core_origin_Edit").First();
                                        Document fa_doc = edit_doc.EditFamily(detailArc);
                                        List<string> para_s = new List<string> { "Xn", "Yn", "B", "H", "PB1", "PB2", "t", "內徑"};
                                        List<string> parValue = new List<string> { Xn.ToString(), Yn.ToString(), B.ToString(), H.ToString(), PB1.ToString(), PB2.ToString(), thick.ToString(), radius.ToString(), amount.ToString() };

                                        ProduceProfile(ref fa_doc, ref para_s, ref parValue);
                                        DrawCurve_Inner(ref fa_doc, ref parValue);
                                        DrawCurve(ref fa_doc, ref parValue);
                                        fa_doc.LoadFamily(edit_doc, new FamilyOption());
                                        fa_doc.Close(false);

                                        FamilySymbol familyinstance_new = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "profile_core_origin_Edit").First();
                                        proCore = familyinstance_new;
                                        
                                        //base
                                        Family detailArc_base = new FilteredElementCollector(edit_doc).OfClass(typeof(Family)).Cast<Family>().ToList().Where(x => x.Name == "profile_base_origin_Edit").First();
                                        Document fa_doc_base = edit_doc.EditFamily(detailArc_base);

                                        ProduceProfile(ref fa_doc_base, ref para_s, ref parValue);
                                        DrawCurve(ref fa_doc_base, ref parValue);
                                        fa_doc_base.LoadFamily(edit_doc, new FamilyOption());
                                        fa_doc_base.Close(false);
                                        
                                        FamilySymbol familyinstance_base_new = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "profile_base_origin_Edit").First();
                                        proBase = familyinstance_base_new;

                                        t.Start("Restart");
                                    }
                                    catch (Exception e) { TaskDialog.Show("asd", e.Message); }
                                }

                            }
                            //建置掃略
                            ReferenceArray reff = new ReferenceArray();
                            List<ElementId> list_mc = new List<ElementId>();

                            double pipedeep = 0;

                            try { pipedeep = double.Parse(rows[16]); } catch { pipedeep = 0; };

                            for (int j = 0; j != count - 1; j++)
                            {
                                XYZ start = (pos_list[index][j] - new XYZ(xshift, yshift, zshift + pipedeep)) * 1000 / 304.8;
                                XYZ end = (pos_list[index][j + 1] - new XYZ(xshift, yshift, zshift + pipedeep)) * 1000 / 304.8;

                                Curve cv = Line.CreateBound(start, end);
                                ModelCurve m_curve = edit_doc.FamilyCreate.NewModelCurve(cv, Sketch_plain(edit_doc, start, end));
                                reff.Append(m_curve.GeometryCurve.Reference);
                                list_mc.Add(m_curve.Id);
                            }

                            ElementId material_id = (from x in materials where x.Name == rows[2] select x.Id).First<ElementId>();

                            if (rows[4] == "方形")
                            {
                                if (double.Parse(rows[8]) != 0)
                                {
                                    try
                                    {
                                        SweepProfile swb = edit_doc.Application.Create.NewFamilySymbolProfile(proBase);
                                        sb_instance = edit_doc.FamilyCreate.NewSweep(true, reff, swb, 0, ProfilePlaneLocation.Start);
                                        sb_instance.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM).Set(concrete_m_id);
                                        sb_instance.get_Parameter(BuiltInParameter.PROFILE_ANGLE).SetValueString("-90");

                                        FamilyManager familyManager = edit_doc.FamilyManager;
                                        FamilyParameter familyParameter = (from x in familyManager.GetParameters() where x.Definition.Name == "VIS" select x).First<FamilyParameter>();

                                        familyManager.AssociateElementParameterToFamilyParameter(sb_instance.get_Parameter(BuiltInParameter.IS_VISIBLE_PARAM), familyParameter);

                                    }
                                    catch (Exception e)
                                    {
                                        TaskDialog.Show("error", e.Message + e.StackTrace + e.Source);
                                    }
                                }
                            }

                            try
                            {
                                SweepProfile swc = edit_doc.Application.Create.NewFamilySymbolProfile(proCore);

                                sc_instance = edit_doc.FamilyCreate.NewSweep(true, reff, swc, 0, ProfilePlaneLocation.Start);

                                sc_instance.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM).Set(material_id);
                                sc_instance.get_Parameter(BuiltInParameter.PROFILE_ANGLE).SetValueString("-90");
                            }
                            catch (Exception e)
                            {
                                TaskDialog.Show("error", e.Message + e.StackTrace + e.Source);
                            }

                            edit_doc.Delete(list_mc);
                            FailureHandlingOptions failOpt = t.GetFailureHandlingOptions();

                            warningSwallower.Pipe_number = rows[1];
                            failOpt.SetFailuresPreprocessor(warningSwallower);
                            t.SetFailureHandlingOptions(failOpt);
                        }
                        catch (Exception e)
                        {
                            TaskDialog.Show("error", e.Message + e.StackTrace + e.Source);
                        };

                        t.Commit();

                        //判斷是否掃掠成功
                        IList<Sweep> sweeps = new FilteredElementCollector(edit_doc).OfClass(typeof(Sweep)).Cast<Sweep>().ToList();
                        if (sweeps.Count == 0)
                        {
                            warningSwallower.Pipe_number_list.Add(warningSwallower.Pipe_number);
                        }

                        SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true, MaximumBackups = 1 };

                        edit_doc.SaveAs(@"C:\Users\TsaiWeiLun\OneDrive\Revit_SinoPipe\edit" + rows[1] + ".rfa", saveAsOptions);
                        app.OpenAndActivateDocument(doc.PathName);
                        edit_doc.Close();

                    }

                    using (Transaction t = new Transaction(doc, "load family"))
                    {
                        t.Start();

                        try
                        {
                            doc.LoadFamily(@"C:\Users\TsaiWeiLun\OneDrive\Revit_SinoPipe\edit" + rows[1] + ".rfa");
                        }
                        catch
                        {
                            ICollection<Family> families = new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>().ToList();
                            string family_name = rows[1];
                            Family family = (from x in families where x.Name == family_name select x).First<Family>();
                            doc.EditFamily(family).LoadFamily(Directory.GetCurrentDirectory() + "\\edit" + rows[1] + ".rfa");
                        }

                        ICollection<FamilySymbol> familysyb = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                        foreach (FamilySymbol fmsy in familysyb)
                        {
                            if (fmsy.Name == "edit" + rows[1])
                            {
                                fmsy.Activate();
                                FamilyInstance instance = doc.Create.NewFamilyInstance(XYZ.Zero, fmsy, StructuralType.NonStructural);
                                instance.LookupParameter("類別碼").Set(rows[0]);
                                instance.LookupParameter("管線編號").Set(rows[1]);
                                instance.LookupParameter("管線類型").Set(rows[2]);
                                instance.LookupParameter("管線材質").Set(rows[3]);
                                instance.LookupParameter("管線型式").Set(rows[4]);
                                instance.LookupParameter("管線長度").Set(rows[15]);
                                instance.LookupParameter("管線總類代碼").Set(rows[17]);
                                instance.LookupParameter("管路規格").Set(rows[12]);
                                instance.LookupParameter("附註").Set(rows[18]);
                                try
                                {
                                    instance.LookupParameter("XY").Set(rows[13] + "," + rows[14]);
                                }
                                catch
                                {
                                    TaskDialog.Show("message", "XY_failure");
                                }

                                if (rows[16] == "")
                                {
                                    rows[16] = "0";
                                }
                                instance.LookupParameter("埋管深度").Set(rows[16]);
                            }
                        }
                        t.Commit();
                    }

                    index++;
                }
                
                //建置套頭
                FilteredElementCollector collector1 = new FilteredElementCollector(doc);
                ICollection<Element> collection = collector1.WhereElementIsNotElementType().ToElements();
                ICollection<FamilySymbol> fsym = null;

                fsym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                FamilySymbol familySymbol_con = fsym.Where(x => x.Name == "接頭").ToList().First();
                
                //起訖點
                XYZ shift = new XYZ(xyz_shift[0], xyz_shift[1], xyz_shift[2]);
                int con_i = 0;
                Transaction con_trans = new Transaction(doc, "con");

                con_trans.Start();
                IList<int> store_con_index = new List<int>();
                IList<string> store_con_name = new List<string>();
                IList<string> store_con_d = new List<string>();
                IList<string> store_con_D = new List<string>();
                IList<string> store_con_t = new List<string>();
                IList<string> store_con_T = new List<string>();
                IList<double> pipedeeps = new List<double>();
                List<List<string>> con_data = new List<List<string>>();

                //儲存所有管線接頭資料
                foreach (List<string> rows in mhx.MHdata)
                {
                    if (rows[19] != "" && store_con_name.Contains(rows[20]) == false)
                    {
                        store_con_index.Add(con_i);
                        store_con_name.Add(rows[20]);
                        store_con_d.Add(rows[21]);
                        store_con_t.Add(rows[11]);
                        store_con_D.Add("");
                        store_con_T.Add("");
                        try { pipedeeps.Add(double.Parse(rows[16])); } catch { pipedeeps.Add(0); };
                        con_data.Add(new List<string> { con_i.ToString(), rows[19], rows[20], rows[21], rows[22] });
                    }
                    else if (rows[19] != "" && store_con_name.Contains(rows[20]) == true)
                    {
                        int number = store_con_name.IndexOf(rows[20]);
                        store_con_D[number] = rows[21];
                        store_con_T[number] = rows[11];
                        con_data.Add(new List<string> {con_i.ToString(), rows[19], rows[20], double.Parse(rows[21]).ToString(), rows[22] });
                    }

                    con_i += 1;
                }

                //建立接頭
                for (int i = 0; i < store_con_name.Count; i++)
                {
                    XYZ center_con = pos_list[store_con_index[i]][0] - new XYZ(xshift, yshift, zshift + pipedeeps[i]);
                    XYZ end_con = pos_list[store_con_index[i]][1] - new XYZ(xshift, yshift, zshift + pipedeeps[i]);
                    XYZ vector = center_con - end_con;

                    double theda = Math.Atan(vector.Y / vector.X) * 180 / Math.PI;

                    if (vector.X < 0) { theda += 180; }

                    FamilyInstance instance = doc.Create.NewFamilyInstance(center_con * 1000 / 304.8, familySymbol_con, StructuralType.NonStructural);

                    instance.LookupParameter("d").SetValueString((Double.Parse(store_con_d[i]) * 1000 + Double.Parse(store_con_t[i]) * 1000 * 2 + 7.7 * 2).ToString());
                    instance.LookupParameter("D").SetValueString((Double.Parse(store_con_D[i]) * 1000 + Double.Parse(store_con_T[i]) * 1000 * 2 + 7.7 * 2).ToString());
                    instance.LookupParameter("方位角").SetValueString(theda.ToString());

                    instance.LookupParameter("附註").Set(con_data[i][0] + "/" + con_data[i + 1][0]);
                    instance.LookupParameter("接頭編號").Set(con_data[i][1] + "/" + con_data[i + 1][1]);
                    instance.LookupParameter("接頭點位").Set(con_data[i][2] + "/" + con_data[i + 1][2]);
                    instance.LookupParameter("接頭管徑").Set(con_data[i][3] + "/" + con_data[i + 1][3]);
                    instance.LookupParameter("接頭連接").Set(con_data[i][4] + "/" + con_data[i + 1][4]);
                }

                con_trans.Commit();
                
                var rebirthdoc = app.OpenAndActivateDocument(filename);
                TaskDialog.Show("Done", "管線建置完畢");
                warningSwallower.Set_failue(warningSwallower.Pipe_number_list);
            }
            catch (Exception e) { TaskDialog.Show("error", e.Message + e.StackTrace); }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
        //訊息處理
        public class WarningSwallower : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(
              FailuresAccessor a)
            {
                IList<FailureMessageAccessor> failureMessageAccessors = a.GetFailureMessages();
                foreach (FailureMessageAccessor fma in failureMessageAccessors)
                {
                    if (fma.GetDescriptionText() == "無法建立掃掠")
                    {
                        a.ResolveFailure(fma);
                        return FailureProcessingResult.ProceedWithCommit;
                    }

                }

                a.DeleteAllWarnings();
                return FailureProcessingResult.Continue;
            }
            public string Pipe_number
            {
                get;
                set;
            }
            public IList<string> Pipe_number_list
            {
                get;
                set;
            }
            public string Case_name
            {
                get;
                set;
            }
            public void Set_failue(IList<string> number_list)
            {
                //excel
                ExReader mhx = new ExReader();
                mhx.SetData(Case_name, 2);
                try
                {
                    if (number_list.Count != 0)
                    {
                        foreach (string number in number_list)
                        {
                            mhx.Change_color(mhx.FindAddress(number));
                        }
                    }
                    mhx.Save_excel(Case_name.Replace(Case_name.Split('\\').Last(), "") + "Result_" + Case_name.Split('\\').Last());
                    mhx.CloseEx();
                }
                catch { mhx.CloseEx(); };


            }
        }

        public SketchPlane Sketch_plain(Document doc, XYZ start, XYZ end)
        {
            SketchPlane sk = null;

            XYZ v = end - start;
            double dxy = Math.Abs(v.X) + Math.Abs(v.Y);
            XYZ w = (dxy > 0.00000001)
              ? XYZ.BasisZ
              : XYZ.BasisY;

            XYZ norm = v.CrossProduct(w).Normalize();
            Plane geomPlane = Plane.CreateByNormalAndOrigin(norm, start);
            sk = SketchPlane.Create(doc, geomPlane);

            return sk;
        }

        public void ProduceProfile(ref Document doc, ref List<string> vs, ref List<string> pv)
        {
            using (Transaction t = new Transaction(doc, "setting profile parameters"))
            {
                t.Start();
                FamilyManager familyManager = doc.FamilyManager;

                for (int i = 0; i < vs.Count; i++)
                {
                    FamilyParameter familyParameter = familyManager.get_Parameter(vs[i]);

                    try
                    {
                        familyManager.SetValueString(familyParameter, pv[i]);
                    }
                    catch
                    {
                        familyManager.Set(familyParameter, double.Parse(pv[i]));
                    }
                }

                t.Commit();
            }
        }

        public void DrawCurve(ref Document doc, ref List<string> xy)
        {
            using (Transaction t = new Transaction(doc, "make profile xn * yn"))
            {
                t.Start();

                CurveElement detailArc = new FilteredElementCollector(doc).OfClass(typeof(CurveElement)).Cast<CurveElement>().ToList().Where(x => x.GeometryCurve.IsBound == false).First();
                ViewPlan viewplan = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().ToList().Where(x => x.Name == "參考樓層").First();

                CurveElement Ndetail = detailArc;
                IList<CurveElement> curveElements = new List<CurveElement>();
                curveElements.Add(Ndetail);
                LocationCurve lc = Ndetail.Location as LocationCurve;

                //distance of the pipe center
                double center_dist1 = 0;
                double center_dist2 = 0;
                if (xy.Count == 5) //圓形
                {
                    //radius * 2 * 2
                    center_dist1 = double.Parse(xy[xy.Count - 2]) * 2 * 2;
                    center_dist2 = center_dist1;
                }
                else if (xy.Count == 9) //方形
                {
                    center_dist1 = 310;
                    center_dist2 = center_dist1;

                    double PB1 = (double.Parse(xy[2]) - 310 * (int.Parse(xy[0]) - 1)) / 2;
                    double PB2 = (double.Parse(xy[3]) - 310 * (int.Parse(xy[1]) - 1)) / 2;

                    //if PB1 <= 內徑 * 2
                    if (PB1 <= double.Parse(xy[xy.Count - 2]) * 2)
                    {
                        center_dist1 = double.Parse(xy[2]) / (int.Parse(xy[0]) + 1);
                    }
                    //if PB2 <= 內徑 * 2
                    if (PB2 <= double.Parse(xy[xy.Count - 2]) * 2)
                    {
                        center_dist2 = double.Parse(xy[3]) / (int.Parse(xy[1]) + 1);
                    }
                }

                for (int i = 0; i < int.Parse(xy[1]); i++)
                {
                    for (int j = 0; j < int.Parse(xy[0]); j++)
                    {
                        if (i != 0 || j != 0)
                        {
                            Transform trans2 = Transform.CreateTranslation(new XYZ(center_dist1 * j / 304.8, center_dist2 * (-i) / 304.8, 0));
                            Curve curve_y = lc.Curve.CreateTransformed(trans2);
                            CurveElement c = doc.FamilyCreate.NewDetailCurve(viewplan, curve_y) as CurveElement;

                            curveElements.Add(c);
                        }
                    }
                }

                //Xn * Yn - amount of pipe
                int n = int.Parse(xy[0]) * int.Parse(xy[1]) - int.Parse(xy[xy.Count - 1]);

                for (int i = 0; i < n; i++)
                {
                    doc.Delete(curveElements[i].Id);
                }

                t.Commit();
            }
        }
        public void DrawCurve_Inner(ref Autodesk.Revit.DB.Document doc, ref List<string> xy)
        {
            using (Transaction t = new Transaction(doc, "make inner profile xn * yn"))
            {
                t.Start();

                CurveElement detailArc = new FilteredElementCollector(doc).OfClass(typeof(CurveElement)).Cast<CurveElement>().ToList().Where(x => x.GeometryCurve.IsBound == false).ToList()[1];
                ViewPlan viewplan = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().ToList().Where(x => x.Name == "參考樓層").First();

                CurveElement Ndetail = detailArc;
                IList<CurveElement> curveElements = new List<CurveElement>();
                curveElements.Add(Ndetail);
                LocationCurve lc = Ndetail.Location as LocationCurve;

                //distance of the pipe center
                double center_dist1 = 0;
                double center_dist2 = 0;
                if (xy.Count == 5) //圓形
                {
                    //radius * 2 * 2
                    center_dist1 = double.Parse(xy[xy.Count - 2]) * 2 * 2;
                    center_dist2 = center_dist1;
                }
                else if (xy.Count == 9) //方形
                {
                    center_dist1 = 310;
                    center_dist2 = center_dist1;

                    double PB1 = (double.Parse(xy[2]) - 310 * (int.Parse(xy[0]) - 1)) / 2;
                    double PB2 = (double.Parse(xy[3]) - 310 * (int.Parse(xy[1]) - 1)) / 2;

                    //if PB1 <= 內徑 * 2
                    if (PB1 <= double.Parse(xy[xy.Count - 2]) * 2)
                    {
                        center_dist1 = double.Parse(xy[2]) / (int.Parse(xy[0]) + 1);
                    }
                    //if PB2 <= 內徑 * 2
                    if (PB2 <= double.Parse(xy[xy.Count - 2]) * 2)
                    {
                        center_dist2 = double.Parse(xy[3]) / (int.Parse(xy[1]) + 1);
                    }
                }

                for (int i = 0; i < int.Parse(xy[1]); i++)
                {
                    for (int j = 0; j < int.Parse(xy[0]); j++)
                    {
                        if (i != 0 || j != 0)
                        {
                            Transform trans2 = Transform.CreateTranslation(new XYZ(center_dist1 * (j) / 304.8, center_dist2 * (-i) / 304.8, 0));
                            Curve curve_y = lc.Curve.CreateTransformed(trans2);
                            CurveElement c = doc.FamilyCreate.NewDetailCurve(viewplan, curve_y) as CurveElement;

                            curveElements.Add(c);
                        }
                    }

                }

                //Xn * Yn - amount of the pipe
                int n = int.Parse(xy[0]) * int.Parse(xy[1]) - int.Parse(xy[xy.Count - 1]);

                for (int i = 0; i < n; i++)
                {
                    doc.Delete(curveElements[i].Id);
                }

                t.Commit();
            }
        }

        class FamilyOption : IFamilyLoadOptions
        {

            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                source = FamilySource.Family;
                return true;
            }
        }

        public void DrawBox(ref Autodesk.Revit.DB.Document doc, ref List<string> xy)
        {
            using (Transaction t = new Transaction(doc, "make profile xn * yn"))
            {

                t.Start();

                IList<CurveElement> detailArc = new FilteredElementCollector(doc).OfClass(typeof(CurveElement)).Cast<CurveElement>().ToList().Where(x => x.GeometryCurve.IsBound == true).ToList();
                ViewPlan viewplan = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().ToList().Where(x => x.Name == "參考樓層").First();
                IList<CurveElement> Ndetails = detailArc;

                double td = double.Parse(xy[3]) + double.Parse(xy[1]);

                for (int j = 4; j < 8; j++)
                {
                    LocationCurve lc = Ndetails[j].Location as LocationCurve;
                    for (int i = 0; i < int.Parse(xy[0]); i++)
                    {
                        if (i != 0)
                        {
                            Transform trans2 = Transform.CreateTranslation(new XYZ(td * (i) / 304.8, 0, 0));
                            Curve curve_y = lc.Curve.CreateTransformed(trans2);
                            CurveElement c = doc.FamilyCreate.NewDetailCurve(viewplan, curve_y) as CurveElement;
                        }
                    }
                }
                t.Commit();
            }
        }
    }
}
