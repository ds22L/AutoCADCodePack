using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Dreambuild.AutoCAD.Internal;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Dreambuild.AutoCAD
{
    /// <summary>
    /// Tests and samples
    /// </summary>
    public class CodePackTest
    {
        #region Commands that you can provide out of the box in your application

        /// <summary>
        /// View or edit custom dictionaries of DWG.
        /// </summary>
        [CommandMethod("ViewGlobalDict")]
        public static void ViewGlobalDict()
        {
            var dv = new DictionaryViewer(
                CustomDictionary.GetDictionaryNames,
                CustomDictionary.GetEntryNames,
                CustomDictionary.GetValue,
                CustomDictionary.SetValue
            );
            Application.ShowModalWindow(dv);
        }

        /// <summary>
        /// View or edit custom dictionaries of entity.
        /// </summary>
        [CommandMethod("ViewObjectDict")]
        public static void ViewObjectDict()
        {
            var id = Interaction.GetEntity("\nSelect entity");
            if (id == ObjectId.Null)
            {
                return;
            }
            var dv = new DictionaryViewer(  // Currying
                () => CustomObjectDictionary.GetDictionaryNames(id),
                dict => CustomObjectDictionary.GetEntryNames(id, dict),
                (dict, key) => CustomObjectDictionary.GetValue(id, dict, key),
                (dict, key, value) => CustomObjectDictionary.SetValue(id, dict, key, value)
            );
            Application.ShowModalWindow(dv);
        }

        /// <summary>
        /// Eliminates zero-length polylines.
        /// </summary>
        [CommandMethod("PolyClean0", CommandFlags.UsePickSet)]
        public static void PolyClean0()
        {
            var ids = Interaction.GetSelection("\nSelect polyline", "LWPOLYLINE");
            int n = 0;
            ids.QForEach<Polyline>(poly =>
            {
                if (poly.Length == 0)
                {
                    poly.Erase();
                    n++;
                }
            });
            Interaction.WriteLine("{0} eliminated.", n);
        }

        /// <summary>
        /// Removes duplicate vertices on polyline.
        /// </summary>
        [CommandMethod("PolyClean", CommandFlags.UsePickSet)]
        public static void PolyClean()
        {
            var ids = Interaction.GetSelection("\nSelect polyline", "LWPOLYLINE");
            int m = 0;
            int n = 0;
            ids.QForEach<Polyline>(poly =>
            {
                int count = Algorithms.PolyClean_RemoveDuplicatedVertex(poly);
                if (count > 0)
                {
                    m++;
                    n += count;
                }
            });
            Interaction.WriteLine("{1} vertex removed from {0} polyline.", m, n);
        }

        private static double _polyClean2Epsilon = 1;

        /// <summary>
        /// Removes vertices close to others on polyline.
        /// </summary>
        [CommandMethod("PolyClean2", CommandFlags.UsePickSet)]
        public static void PolyClean2()
        {
            double epsilon = Interaction.GetValue("\nEpsilon", _polyClean2Epsilon);
            if (double.IsNaN(epsilon))
            {
                return;
            }
            _polyClean2Epsilon = epsilon;

            var ids = Interaction.GetSelection("\nSelect polyline", "LWPOLYLINE");
            int m = 0;
            int n = 0;
            ids.QForEach<Polyline>(poly =>
            {
                int count = Algorithms.PolyClean_ReducePoints(poly, epsilon);
                if (count > 0)
                {
                    m++;
                    n += count;
                }
            });
            Interaction.WriteLine("{1} vertex removed from {0} polyline.", m, n);
        }

        /// <summary>
        /// Fits arc segs of polyline with line segs.
        /// </summary>
        [CommandMethod("PolyClean3", CommandFlags.UsePickSet)]
        public static void PolyClean3()
        {
            double value = Interaction.GetValue("\nNumber of segs to fit an arc, 0 for smart determination", 0);
            if (double.IsNaN(value))
            {
                return;
            }
            int n = (int)value;

            var ids = Interaction.GetSelection("\nSelect polyline", "LWPOLYLINE");
            var entsToAdd = new List<Polyline>();
            ids.QForEach<Polyline>(poly =>
            {
                var pts = poly.GetPolylineFitPoints(n);
                var poly1 = NoDraw.Pline(pts);
                poly1.Layer = poly.Layer;
                try
                {
                    poly1.ConstantWidth = poly.ConstantWidth;
                }
                catch
                {
                }
                poly1.XData = poly.XData;
                poly.Erase();
                entsToAdd.Add(poly1);
            });
            entsToAdd.ToArray().AddToCurrentSpace();
            Interaction.WriteLine("{0} handled.", entsToAdd.Count);
        }

        /// <summary>
        /// Regulates polyline direction.
        /// </summary>
        [CommandMethod("PolyClean4", CommandFlags.UsePickSet)]
        public static void PolyClean4()
        {
            double value = Interaction.GetValue("\nDirection：1-R to L；2-B to T；3-L to R；4-T to B");
            if (double.IsNaN(value))
            {
                return;
            }
            int n = (int)value;
            if (!new int[] { 1, 2, 3, 4 }.Contains(n))
            {
                return;
            }
            Algorithms.Direction dir = (Algorithms.Direction)n;

            var ids = Interaction.GetSelection("\nSelect polyline", "LWPOLYLINE");
            int m = 0;
            ids.QForEach<Polyline>(poly =>
            {
                if (Algorithms.PolyClean_SetTopoDirection(poly, dir))
                {
                    m++;
                }
            });
            Interaction.WriteLine("{0} handled.", m);
        }

        /// <summary>
        /// Removes unnecessary colinear vertices on polyline.
        /// </summary>
        [CommandMethod("PolyClean5", CommandFlags.UsePickSet)]
        public static void PolyClean5()
        {
            Interaction.WriteLine("Not implemented yet");
            var ids = Interaction.GetSelection("\nSelect polyline", "LWPOLYLINE");
            ids.QForEach<Polyline>(poly =>
            {
                Algorithms.PolyClean_RemoveColinearPoints(poly);
            });
        }

        /// <summary>
        /// Breaks polylines at their intersecting points.
        /// </summary>
        [CommandMethod("PolySplit", CommandFlags.UsePickSet)]
        public static void PolySplit()
        {
            var ids = Interaction.GetSelection("\nSelect polyline", "LWPOLYLINE");
            var newPolys = new List<Polyline>();
            var pm = new ProgressMeter();
            pm.Start("Processing...");
            pm.SetLimit(ids.Length);
            ids.QOpenForWrite<Polyline>(list =>
            {
                foreach (var poly in list)
                {
                    var intersectPoints = new Point3dCollection();
                    foreach (var poly1 in list)
                    {
                        if (poly1 != poly)
                        {
                            poly.IntersectWith3264(poly1, Intersect.OnBothOperands, intersectPoints);
                        }
                    }
                    var ipParams = intersectPoints
                        .Cast<Point3d>()
                        .Select(ip => poly.GetParamAtPointX(ip))
                        .OrderBy(param => param)
                        .ToArray();
                    if (intersectPoints.Count > 0)
                    {
                        var curves = poly.GetSplitCurves(new DoubleCollection(ipParams));
                        foreach (var curve in curves)
                        {
                            newPolys.Add(curve as Polyline);
                        }
                    }
                    else // mod 20130227 Add to newPolys regardless of whether an intersection exists, otherwise dangling lines would be gone.
                    {
                        newPolys.Add(poly.Clone() as Polyline);
                    }
                    pm.MeterProgress();
                }
            });
            pm.Stop();
            if (newPolys.Count > 0)
            {
                newPolys.ToArray().AddToCurrentSpace();
                ids.QForEach(entity => entity.Erase());
            }
            Interaction.WriteLine("Broke {0} to {1}.", ids.Length, newPolys.Count);
        }

        private static double _polyTrimExtendEpsilon = 20;

        /// <summary>
        /// Handles polylines that are by a small distance longer than, shorter than, or mis-intersecting with each other.
        /// </summary>
        [CommandMethod("PolyTrimExtend", CommandFlags.UsePickSet)]
        public static void PolyTrimExtend() // mod 20130228
        {
            double epsilon = Interaction.GetValue("\nEpsilon", _polyTrimExtendEpsilon);
            if (double.IsNaN(epsilon))
            {
                return;
            }
            _polyTrimExtendEpsilon = epsilon;

            var visibleLayers = DbHelper
                .GetAllLayerIds()
                .QOpenForRead<LayerTableRecord>()
                .Where(layer => !layer.IsHidden && !layer.IsFrozen && !layer.IsOff)
                .Select(layer => layer.Name)
                .ToList();

            var ids = Interaction
                .GetSelection("\nSelect polyline", "LWPOLYLINE")
                .QWhere(pline => visibleLayers.Contains(pline.Layer) && pline.Visible)
                .ToArray(); // newly 20130729

            var pm = new ProgressMeter();
            pm.Start("Processing...");
            pm.SetLimit(ids.Length);
            ids.QOpenForWrite<Polyline>(list =>
            {
                foreach (var poly in list)
                {
                    int[] indices = { 0, poly.NumberOfVertices - 1 };
                    foreach (int index in indices)
                    {
                        var end = poly.GetPoint3dAt(index);
                        foreach (var poly1 in list)
                        {
                            if (poly1 != poly)
                            {
                                var closest = poly1.GetClosestPointTo(end, false);
                                double dist = closest.DistanceTo(end);
                                double dist1 = poly1.StartPoint.DistanceTo(end);
                                double dist2 = poly1.EndPoint.DistanceTo(end);

                                double distance = poly1.GetDistToPoint(end);
                                if (poly1.GetDistToPoint(end) > 0)
                                {
                                    if (dist1 <= dist2 && dist1 <= dist && dist1 < epsilon)
                                    {
                                        poly.SetPointAt(index, new Point2d(poly1.StartPoint.X, poly1.StartPoint.Y));
                                    }
                                    else if (dist2 <= dist1 && dist2 <= dist && dist2 < epsilon)
                                    {
                                        poly.SetPointAt(index, new Point2d(poly1.EndPoint.X, poly1.EndPoint.Y));
                                    }
                                    else if (dist <= dist1 && dist <= dist2 && dist < epsilon)
                                    {
                                        poly.SetPointAt(index, new Point2d(closest.X, closest.Y));
                                    }
                                }
                            }
                        }
                    }
                    pm.MeterProgress();
                }
            });
            pm.Stop();
        }

        /// <summary>
        /// Saves selection for later use.
        /// </summary>
        [CommandMethod("SaveSelection", CommandFlags.UsePickSet)]
        public static void SaveSelection()
        {
            var ids = Interaction.GetPickSet();
            if (ids.Length == 0)
            {
                Interaction.WriteLine("No entity selected.");
                return;
            }
            string name = Interaction.GetString("\nSelection name");
            if (name == null)
            {
                return;
            }
            if (CustomDictionary.GetValue("Selections", name) != string.Empty)
            {
                Interaction.WriteLine("Selection with the same name already exists.");
                return;
            }
            var handles = ids.QSelect(entity => entity.Handle.Value.ToString()).ToArray();
            string dictValue = string.Join("|", handles);
            CustomDictionary.SetValue("Selections", name, dictValue);
        }

        /// <summary>
        /// Loads previously saved selection.
        /// </summary>
        [CommandMethod("LoadSelection")]
        public static void LoadSelection()
        {
            string name = Gui.GetChoice("Which selection to load?", CustomDictionary.GetEntryNames("Selections").ToArray());
            if (name == string.Empty)
            {
                return;
            }
            string dictValue = CustomDictionary.GetValue("Selections", name);
            var handles = dictValue.Split('|').Select(value => new Handle(Convert.ToInt64(value))).ToList();
            var ids = new List<ObjectId>();
            handles.ForEach(value =>
            {
                var id = ObjectId.Null;
                if (HostApplicationServices.WorkingDatabase.TryGetObjectId(value, out id))
                {
                    ids.Add(id);
                }
            });
            Interaction.SetPickSet(ids.ToArray());
        }

        /// <summary>
        /// Converts MText to Text.
        /// </summary>
        [CommandMethod("MT2DT", CommandFlags.UsePickSet)]
        public static void MT2DT() // newly 20130815
        {
            var ids = Interaction.GetSelection("\nSelect MText", "MTEXT");
            var mts = ids.QOpenForRead<MText>().Select(mt =>
            {
                var dt = NoDraw.Text(mt.Text, mt.TextHeight, mt.Location, mt.Rotation, false, mt.TextStyleName);
                dt.Layer = mt.Layer;
                return dt;
            }).ToArray();
            ids.QForEach(mt => mt.Erase());
            mts.AddToCurrentSpace();
        }

        /// <summary>
        /// Converts Text to MText.
        /// </summary>
        [CommandMethod("DT2MT", CommandFlags.UsePickSet)]
        public static void DT2MT() // newly 20130815
        {
            var ids = Interaction.GetSelection("\nSelect Text", "TEXT");
            var dts = ids.QOpenForRead<DBText>().Select(dt =>
            {
                var mt = NoDraw.MText(dt.TextString, dt.Height, dt.Position, dt.Rotation, false);
                mt.Layer = dt.Layer;
                return mt;
            }).ToArray();
            ids.QForEach(dt => dt.Erase());
            dts.AddToCurrentSpace();
        }

        /// <summary>
        /// Shows a rectangle indicating the extents of selected entities.
        /// </summary>
        [CommandMethod("ShowExtents", CommandFlags.UsePickSet)]
        public static void ShowExtents() // newly 20130815
        {
            var ids = Interaction.GetSelection("\nSelect entity");
            var extents = ids.GetExtents();
            var rectId = Draw.Rectang(extents.MinPoint, extents.MaxPoint);
            Interaction.GetString("\nPress ENTER to exit...");
            rectId.QOpenForWrite(rect => rect.Erase());
        }

        /// <summary>
        /// Closes a polyline by adding a vertex at the same position as the start rather than setting IsClosed=true.
        /// </summary>
        [CommandMethod("ClosePolyline", CommandFlags.UsePickSet)]
        public static void ClosePolyline()
        {
            var ids = Interaction.GetSelection("\nSelect polyline", "LWPOLYLINE");
            if (ids.Length == 0)
            {
                return;
            }
            if (Interaction.TaskDialog(
                mainInstruction: ids.Count().ToString() + " polyline(s) selected. Make sure what you select is correct.",
                yesChoice: "Yes, I promise.",
                noChoice: "No, I want to double check.",
                title: "AutoCAD",
                content: "All polylines in selection will be closed.",
                footer: "Abuse can mess up the drawing.",
                expanded: "Commonly used before export."))
            {
                //polys.QForEach(poly => LogManager.Write((poly as Polyline).Closed));
                ids.QForEach<Polyline>(poly =>
                {
                    if (poly.StartPoint.DistanceTo(poly.EndPoint) > 0)
                    {
                        poly.AddVertexAt(poly.NumberOfVertices, poly.StartPoint.ToPoint2d(), 0, 0, 0);
                    }
                });
            }
        }

        /// <summary>
        /// Detects non-simple polylines that intersect with themselves.
        /// </summary>
        [CommandMethod("DetectSelfIntersection")]
        public static void DetectSelfIntersection() // mod 20130202
        {
            var ids = QuickSelection.SelectAll("LWPOLYLINE").ToArray();
            var meter = new ProgressMeter();
            meter.Start("Detecting...");
            meter.SetLimit(ids.Length);
            var results = ids.QWhere(pline =>
            {
                bool result = (pline as Polyline).IsSelfIntersecting();
                meter.MeterProgress();
                System.Windows.Forms.Application.DoEvents();
                return result;
            }).ToList();
            meter.Stop();
            if (results.Count() > 0)
            {
                Interaction.WriteLine("{0} detected.", results.Count());
                Interaction.ZoomHighlightView(results);
            }
            else
            {
                Interaction.WriteLine("0 detected.");
            }
        }

        /// <summary>
        /// Finds entity by handle value.
        /// </summary>
        [CommandMethod("ShowObject")]
        public static void ShowObject()
        {
            var ids = QuickSelection.SelectAll().ToArray();
            double handle1 = Interaction.GetValue("Handle of entity");
            if (double.IsNaN(handle1))
            {
                return;
            }
            long handle2 = Convert.ToInt64(handle1);
            var id = HostApplicationServices.WorkingDatabase.GetObjectId(false, new Handle(handle2), 0);
            var col = new ObjectId[] { id };
            Interaction.HighlightObjects(col);
            Interaction.ZoomObjects(col);
        }

        /// <summary>
        /// Shows the shortest line to link given point to existing lines, polylines, or arcs.
        /// </summary>
        [CommandMethod("PolyLanding")]
        public static void PolyLanding()
        {
            var ids = QuickSelection.SelectAll("*LINE,ARC").ToArray();
            var landingLineIds = new List<ObjectId>();
            while (true)
            {
                var p = Interaction.GetPoint("\nSpecify a point");
                if (p.IsNull())
                {
                    break;
                }
                var landings = ids.QSelect(entity => (entity as Curve).GetClosestPointTo(p, false)).ToArray();
                double minDist = landings.Min(point => point.DistanceTo(p));
                var landing = landings.First(point => point.DistanceTo(p) == minDist);
                Interaction.WriteLine("Shortest landing distance of point ({0:0.00},{1:0.00}) is {2:0.00}.", p.X, p.Y, minDist);
                landingLineIds.Add(Draw.Line(p, landing));
            }
            landingLineIds.QForEach(entity => entity.Erase());
        }

        /// <summary>
        /// Shows vertics info of a polyline.
        /// </summary>
        [CommandMethod("PolylineInfo")]
        public static void PolylineInfo() // mod by WY 20130202
        {
            var id = Interaction.GetEntity("\nSpecify a polyline", typeof(Polyline));
            if (id == ObjectId.Null)
            {
                return;
            }
            var poly = id.QOpenForRead<Polyline>();
            for (int i = 0; i <= poly.EndParam; i++)
            {
                Interaction.WriteLine("[Point {0}] coord: {1}; bulge: {2}", i, poly.GetPointAtParameter(i), poly.GetBulgeAt(i));
            }
            var txtIds = new List<ObjectId>();
            double height = poly.GeometricExtents.MaxPoint.DistanceTo(poly.GeometricExtents.MinPoint) / 50.0;
            for (int i = 0; i < poly.NumberOfVertices; i++)
            {
                txtIds.Add(Draw.MText(i.ToString(), height, poly.GetPointAtParameter(i), 0, true));
            }
            Interaction.GetString("\nPress ENTER to exit");
            txtIds.QForEach(mt => mt.Erase());
        }

        /// <summary>
        /// Selects entities on given layer.
        /// </summary>
        [CommandMethod("SelectByLayer")]
        public static void SelectByLayer()
        {
            var availableLayerNames = DbHelper.GetAllLayerNames();
            var selectedLayerNames = Gui.GetChoices("Specify layers", availableLayerNames);
            if (selectedLayerNames.Length < 1)
            {
                return;
            }

            var ids = QuickSelection
                .SelectAll(FilterList.Create().Layer(selectedLayerNames))
                .ToArray();

            Interaction.SetPickSet(ids);
        }

        /// <summary>
        /// Marks layer names of selected entities on the drawing by MText.
        /// </summary>
        [CommandMethod("ShowLayerName")]
        public static void ShowLayerName()
        {
            double height = 10;
            string[] range = { "By entities", "By layers" };
            int result = Gui.GetOption("Choose one way", range);
            if (result == -1)
            {
                return;
            }
            ObjectId[] ids;
            if (result == 0)
            {
                ids = Interaction.GetSelection("\nSelect entities");
                ids
                    .QWhere(entity => !entity.Layer.Contains("_Label"))
                    .QSelect(entity => entity.Layer)
                    .Distinct()
                    .Select(layer => $"{layer}_Label")
                    .ForEach(labelLayer => DbHelper.GetLayerId(labelLayer));
            }
            else
            {
                var layers = DbHelper.GetAllLayerNames().Where(layer => !layer.Contains("_Label")).ToArray();
                string layerName = Gui.GetChoice("Select a layer", layers);
                ids = QuickSelection
                    .SelectAll(FilterList.Create().Layer(layerName))
                    .ToArray();

                DbHelper.GetLayerId($"{layerName}_Label");
            }
            var texts = new List<MText>();
            ids.QForEach<Entity>(entity =>
            {
                string layerName = entity.Layer;
                if (!layerName.Contains("_Label"))
                {
                    var center = entity.GetCenter();
                    var text = NoDraw.MText(layerName, height, center, 0, true);
                    text.Layer = $"{layerName}_Label";
                    texts.Add(text);
                }
            });
            texts.ToArray().AddToCurrentSpace();
        }

        /// <summary>
        /// Inspects object in property palette.
        /// </summary>
        [CommandMethod("InspectObject")]
        public static void InspectObject()
        {
            var id = Interaction.GetEntity("\nSelect objects");
            if (id.IsNull)
            {
                return;
            }
            Gui.PropertyPalette(id.QOpenForRead());
        }

        #endregion
    }
}
