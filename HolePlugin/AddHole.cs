using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [TransactionAttribute(TransactionMode.Manual)]
    //[RegenerationAttribute(RegenerationOption.)]
    public class AddHole : IExternalCommand
    {        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.EndsWith("_ОВ")).SingleOrDefault();
            if (ovDoc==null)
            {
                message = "Ошибка! Не найден ОВ-файл.";
                return Result.Failed;
            }

            FamilySymbol holeFS = new FilteredElementCollector(arDoc)
                //.OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .SingleOrDefault();

            if (holeFS==null)
            {
                message = "Ошибка! Семейство \"Отверстие\" не загружено в АР-файл.";
                return Result.Failed;
            }

            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x=>!x.IsTemplate)
                .SingleOrDefault();
            
            if (view3D==null)
            {
                message = "Ошибка! АР-файл не содержит 3D вид.";
                return Result.Failed;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction ts = new Transaction(arDoc, "Hole Insert Transaction");
            ts.Start();

            if (!holeFS.IsActive)
                holeFS.Activate();

            int ductHoleQTY = 0;
            int pipeHoleQTY = 0;

            foreach (Duct d in ducts)
            {
                Line curve = (d.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer()) // оставляет единственный 
                    .ToList();
                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall hostWall = arDoc.GetElement(reference.ElementId) as Wall;

                    Level level = arDoc.GetElement(hostWall.LevelId) as Level;

                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, holeFS, hostWall, level, StructuralType.NonStructural);

                    ductHoleQTY += 1;

                    Parameter width = hole.LookupParameter("Ширина");
                    width.Set(d.Diameter);
                    Parameter height = hole.LookupParameter("Высота");
                    height.Set(d.Diameter);
                }
            }

            foreach (Pipe p in pipes)
            {
                Line curve = (p.Location as LocationCurve).Curve as Line;

                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer()) // оставляет единственный 
                    .ToList();

                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall hostWall = arDoc.GetElement(reference.ElementId) as Wall;

                    Level level = arDoc.GetElement(hostWall.LevelId) as Level;

                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, holeFS, hostWall, level, StructuralType.NonStructural);

                    pipeHoleQTY += 1;

                    Parameter width = hole.LookupParameter("Ширина");
                    width.Set(p.Diameter);
                    Parameter height = hole.LookupParameter("Высота");
                    height.Set(p.Diameter);
                }
            }

            //arDoc.Regenerate(); // Не помогло устранить проблему с отображением отверстий на 3D

            ts.Commit();

            //commandData.Application.ActiveUIDocument.RefreshActiveView(); // Не помогло устранить проблему с отображением отверстий на 3D
            //commandData.Application.ActiveUIDocument.UpdateAllOpenViews(); // Не помогло устранить проблему с отображением отверстий на 3D

            TaskDialog.Show("Выполнено", $"Создано отверстий:{Environment.NewLine}  для воздуховодов - {ductHoleQTY};{Environment.NewLine}  для труб - {pipeHoleQTY}.");
            return Result.Succeeded;
        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
