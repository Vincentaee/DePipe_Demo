using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace UngroupGroups
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class SecCreate : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true, MaximumBackups = 1 };

            var file = doc.PathName;

            int dia = 50;     //管徑
            double t = 4.1;   //壁厚

            bool baseOrCore = true;  // true:base  false:core

            double pb1 = 0.2 * 1000;
            double pb2 = 1000 * dia / 500;

            string file_name;
            string file_path;

            List<int> ray = new List<int> { 2, 3, 4 };
            List<string> vs = new List<string> { "PB1", "PB2", "Xn", "Yn", "內徑", "壁厚" };

            for (int i = 0; i!= ray.Count; i++)
            {
                for(int j = 0; j!= ray.Count; j++)
                {
                    int xn = ray[i];
                    int yn = ray[j];
                    int num = xn * yn;
                    List<string> parValue = new List<string> { pb1.ToString(), pb2.ToString(), xn.ToString(), yn.ToString(), (dia/2).ToString(), t.ToString() };

                    ProduceProfile(ref doc,ref vs,ref parValue);
                    UnGroupAll(ref doc);
                    DeleteLines(ref doc);

                    if (baseOrCore)
                    {
                        if (xn >= yn)
                            file_name = dia + "x" + num + "h_base";
                        else
                            file_name = dia + "x" + num + "v_base";
                    }

                    else
                    {
                        if (xn >= yn)
                            file_name = dia + "x" + num + "h_core";
                        else
                            file_name = dia + "x" + num + "v_core";
                    }



                    file_path = @"D:\Sino\Code\PipeV2\profile\base\" + file_name + ".rfa";
                    doc.SaveAs(file_path, saveAsOptions);

                    var redoc = commandData.Application.OpenAndActivateDocument(file);
                    doc.Close(true);
                    doc = redoc.Document;


                }
            }

            return Result.Succeeded;
        }

        public void ProduceProfile(ref Document doc, ref List<string> vs, ref List<string> pv)
        {
            using (Transaction t = new Transaction(doc, "make profile xn * yn"))
            {

                t.Start();
                FamilyManager familyManager = doc.FamilyManager;

                for (int i = 0; i != vs.Count; i++)
                {
                    FamilyParameter familyParameter = familyManager.get_Parameter(vs[i]);
                    try
                    {
                        familyManager.SetValueString(familyParameter, pv[i]);
                    }
                    catch
                    {
                        familyManager.Set(familyParameter, int.Parse(pv[i]));
                    }
                }


                t.Commit();
            }
         }


        public void UnGroupAll(ref Document doc)
        {
            using (Transaction t = new Transaction(doc, "Ungroup All Groups"))
            {
                t.Start();

                foreach (GroupType groupType in new FilteredElementCollector(doc).OfClass(typeof(GroupType))
                       .Cast<GroupType>()
                       .Where(gt => gt.Category.Name == "詳圖群組" && gt.Groups.Size > 0))
                {
                    foreach (Group group in groupType.Groups)
                    {
                        group.UngroupMembers();
                    }
                }

                t.Commit();
            }

        }

       public void DeleteLines(ref Document doc)
        {
            using (Transaction t = new Transaction(doc, "Delete unnecessary lines"))
            {
                t.Start();


                List<ElementId> list_detail = new List<ElementId>();
                ICollection<CurveElement> lines = new FilteredElementCollector(doc).OfClass(typeof(CurveElement)).Cast<CurveElement>().ToList();

                foreach (CurveElement line in lines)
                {
                    if (line.LineStyle.Name == "<不可見的線>")
                    {
                        list_detail.Add(line.Id);
                    }
                }

                doc.Delete(list_detail);
                t.Commit();
            }
        }
    }

}
