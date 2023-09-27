using System;
using System.Windows.Forms;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.UI;
using System.Collections;
using Tekla.Structures.Solid;
using System.Linq;
using System.Drawing.Drawing2D;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Policy;
using System.Runtime.InteropServices;

namespace Beam_to_column_Connection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        static void RefreshViews()

        {
            var views = ViewHandler.GetVisibleViews();
            while (views.MoveNext())
            ViewHandler.RedrawView(views.Current);

        }
        private void button1_Click(object sender, EventArgs e)
        {
            Model model = new Model();
            Picker pi = new Picker();
            ModelObject modelObject = pi.PickObject(Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART);

            // Get the coordinate system of the part
            CoordinateSystem PartCoordinate = modelObject.GetCoordinateSystem();

            //Create a new transformation plane defined based on the parts coordinate system
            TransformationPlane PartPlane = new TransformationPlane(PartCoordinate);

            //Set the current transformation plane to the part plane.
            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(PartPlane);
            RefreshViews();

            var P1 = pi.PickObject(Picker.PickObjectEnum.PICK_ONE_OBJECT, ("Pick a Column")) as Part;
            //Beam b1 = new Beam();
            ArrayList MyList = new ArrayList();
            ArrayList MyFaceNormalList = new ArrayList();
            Solid Solid = P1.GetSolid();
            double Count = 0;
            FaceEnumerator MyFaceEnum = Solid.GetFaceEnumerator();
            while (MyFaceEnum.MoveNext())
            {
                Count++;
                Face MyFace = MyFaceEnum.Current as Face;
                if (MyFace != null)
                {
                    MyFaceNormalList.Add(MyFace.Normal);
                    LoopEnumerator MyLoopEnum = MyFace.GetLoopEnumerator();
                    while (MyLoopEnum.MoveNext())
                    {
                        Loop MyLoop = MyLoopEnum.Current as Loop;
                        if (MyLoop != null)
                        {
                            VertexEnumerator MyVertexEnum = MyLoop.GetVertexEnumerator() as VertexEnumerator;
                            while (MyVertexEnum.MoveNext())
                            {
                                Point MyVertex = MyVertexEnum.Current as Point;
                                if (MyVertex != null)
                                {
                                    MyList.Add(MyVertex);
                                }
                            }
                        }
                    }
                }
            }

            MyList.Sort();
            ArrayList spair_Vertises = new ArrayList();
            foreach (Point p in MyList)
            {
                if (spair_Vertises.Contains(p)) { spair_Vertises.Remove(p); }
                else { spair_Vertises.Add(p); };
            }

            int i = 0;
            foreach (Point p in spair_Vertises)
            {
                new GraphicsDrawer().DrawText(p, i.ToString(), new Color(0, 0, 1));
                i++;
            }

            GeometricPlane gp = new GeometricPlane()
            {
                Origin = spair_Vertises[0] as Point,
                Normal = new Vector(0, 1, 0)
            };
            GeometricPlane gp_1 = new GeometricPlane()
            {
                Origin = spair_Vertises[23] as Point,
                Normal = new Vector(0, 1, 0)
            };
            GeometricPlane[] GeometricPlanes_1 = { gp, gp_1 };

            var p2 = pi.PickObject(Picker.PickObjectEnum.PICK_ONE_OBJECT, ("Pick a Beam ")) as Part;
            ArrayList MyList1 = new ArrayList();
            ArrayList MyFaceNormalList1 = new ArrayList();
            Solid Solid1 = p2.GetSolid();
            double Count1 = 0;
            FaceEnumerator MyFaceEnum1 = Solid1.GetFaceEnumerator();
            while (MyFaceEnum1.MoveNext())
            {
                Count1++;
                Face MyFace1 = MyFaceEnum1.Current as Face;
                if (MyFace1 != null)
                {
                    MyFaceNormalList1.Add(MyFace1.Normal);
                    LoopEnumerator MyLoopEnum1 = MyFace1.GetLoopEnumerator();
                    while (MyLoopEnum1.MoveNext())
                    {
                        Loop MyLoop1 = MyLoopEnum1.Current as Loop;
                        if (MyLoop1 != null)
                        {
                            VertexEnumerator MyVertexEnum1 = MyLoop1.GetVertexEnumerator() as VertexEnumerator;
                            while (MyVertexEnum1.MoveNext())
                            {
                                Point MyVertex1 = MyVertexEnum1.Current as Point;
                                if (MyVertex1 != null)
                                {
                                    MyList1.Add(MyVertex1);
                                }
                            }
                        }
                    }
                }
            }
            MyList1.Sort();
            ArrayList spair_Vertises1 = new ArrayList();
            foreach (Point p in MyList1)
            {
                if (spair_Vertises1.Contains(p)) { spair_Vertises1.Remove(p); }
                else { spair_Vertises1.Add(p); };
            }
            int j = 0;
            foreach (Point p in spair_Vertises1)
            {
                new GraphicsDrawer().DrawText(p, j.ToString(), new Color(0, 0, 1));
                j++;
            }
            var Centerline = p2.GetCenterLine(false);
            var intersect = Solid1.Intersect(new LineSegment(Centerline[0] as Point, Centerline[1] as Point));
            Point midPoint = new Point(
           ((Centerline[0] as Point).X + (Centerline[1] as Point).X) / 2,
           ((Centerline[0] as Point).Y + (Centerline[1] as Point).Y) / 2,
           ((Centerline[0] as Point).Z + (Centerline[1] as Point).Z) / 2);

            new GraphicsDrawer().DrawText(midPoint, "MId", new Color());
            double mindista = double.MaxValue;
            Point near1 = null;
            Point near2 = null;
            foreach (GeometricPlane p in GeometricPlanes_1)
            {
                var p1_ = Distance.PointToPlane(midPoint, p);
                if (p1_ < mindista)
                {
                    mindista = p1_;
                    near1 = midPoint;
                    near2 = p.Origin;
                }
            }

            new GraphicsDrawer().DrawText(near2, "NEAR2", new Color());
            ArrayList Beam_1 = new ArrayList();
            ArrayList Beam_2 = new ArrayList();
            foreach (Point p in spair_Vertises1)
            {
                if (p.Y < midPoint.Y)
                {
                    Beam_1.Add(p);
                }
                else
                {
                    Beam_2.Add(p);

                }

            }

            new GraphicsDrawer().DrawText(Centerline[1] as Point, "1", new Color());
            new GraphicsDrawer().DrawText(Centerline[0] as Point, "0", new Color());

            double distance_1 = Distance.PointToPoint(spair_Vertises1[0] as Point, Centerline[1] as Point);
            double distance_2 = Distance.PointToPoint(spair_Vertises1[23] as Point, Centerline[0] as Point);

            if (Convert.ToInt16(distance_1) != Convert.ToInt16(distance_2))
            {
                distance_1 = Distance.PointToPoint(spair_Vertises1[0] as Point, Centerline[0] as Point);
                distance_2 = Distance.PointToPoint(spair_Vertises1[23] as Point, Centerline[1] as Point);
            }


            ArrayList outermost = new ArrayList();
            foreach (Point p in spair_Vertises1)
            {
                var p1_ = Distance.PointToPoint(p, Centerline[1] as Point);
                if (p1_ == distance_1)
                {
                    outermost.Add(p);
                    new GraphicsDrawer().DrawText(p, p.ToString(), new Color());
                    //Point pro1 = Projection.PointToPlane(p, gp);
                    //new GraphicsDrawer().DrawText(pro1, "Projection", new Color(0, 1, 0));
                }
            }
            //Point[] p7 = { outermost[0] as Point, outermost[1] as Point, outermost[2] as Point, outermost[3] as Point };
            Point point = outermost[0] as Point;
            Point point1 = outermost[1] as Point;
            Point point2 = outermost[2] as Point;
            Point point3 = outermost[3] as Point;

            double mindista1 = double.MaxValue;
            Point near4 = null;
            Point near5 = null;
            Point near6 = null;
            Point near7 = null;
            foreach (GeometricPlane p in GeometricPlanes_1)
            {
                var p1_ = Distance.PointToPlane(point, p);
                if (p1_ < mindista1)
                {
                    mindista1 = p1_;
                    //near3 = point;
                    //near4 = p.Origin;
                    near4 = Projection.PointToPlane(point, p);
                    near5 = Projection.PointToPlane(point1, p);
                    near6 = Projection.PointToPlane(point2, p);
                    near7 = Projection.PointToPlane(point3, p);
                }
            }

            new GraphicsDrawer().DrawText(near4, "Projection ", new Color(0, 1, 0));
            new GraphicsDrawer().DrawText(near5, "Projection1", new Color(0, 1, 0));
            new GraphicsDrawer().DrawText(near6, "Projection2", new Color(0, 1, 0));
            new GraphicsDrawer().DrawText(near7, "Projection3", new Color(0, 1, 0));
            ArrayList outermost1 = new ArrayList();
            foreach (Point p in spair_Vertises1)
            {
                var p1_ = Distance.PointToPoint(p, Centerline[0] as Point);
                if (p1_ == distance_2)
                {
                    outermost1.Add(p);
                    new GraphicsDrawer().DrawText(p, p.ToString(), new Color());
                }
            }
            //Point[] p8 = { outermost1[0] as Point, outermost1[1] as Point, outermost1[2] as Point, outermost1[3] as Point };
            Point point4 = outermost1[0] as Point;
            Point point5 = outermost1[1] as Point;
            Point point6 = outermost1[2] as Point;
            Point point7 = outermost1[3] as Point;

            //double mindista2 = double.MaxValue;
            //Point near8 = null;
            //Point near9 = null;
            //Point near10 = null;
            //Point near11 = null;
            //foreach (GeometricPlane p in GeometricPlanes_1)
            //{
            //    var p1_ = Distance.PointToPlane(point4, p);
            //    if (p1_ < mindista2)
            //    {
            //        mindista2 = p1_;
            //        //near3 = point;
            //        //near4 = p.Origin;
            //        near8 = Projection.PointToPlane(point4, p);
            //        near9 = Projection.PointToPlane(point5, p);
            //        near10 = Projection.PointToPlane(point6, p);
            //        near11 = Projection.PointToPlane(point7, p);
            //    }
            //}
            //new GraphicsDrawer().DrawText(near8, "Projection4 ", new Color(0, 1, 0));
            //new GraphicsDrawer().DrawText(near9, "Projection5", new Color(0, 1, 0));
            //new GraphicsDrawer().DrawText(near10, "Projection6", new Color(0, 1, 0));
            //new GraphicsDrawer().DrawText(near11, "Projection7", new Color(0, 1, 0));

            ContourPoint point_ = new ContourPoint(new Point(near4), null);
            ContourPoint point4_ = new ContourPoint(new Point(near5), null);
            ContourPoint point5_ = new ContourPoint(new Point(near7), null);
            ContourPoint point3_ = new ContourPoint(new Point(near6), null);
            ContourPlate CP = new ContourPlate();
            CP.AddContourPoint(point_);
            CP.AddContourPoint(point4_);
            CP.AddContourPoint(point5_);
            CP.AddContourPoint(point3_);
            CP.Profile.ProfileString = "PL6.35";
            CP.Material.MaterialString = "A36";
            //CP.Position.Depth = Position.DepthEnum.BEHIND;
            CP.Insert();


            //boolean

            BooleanPart Beam = new BooleanPart();
            Beam.Father = p2;
            CP.Class = BooleanPart.BooleanOperativeClassName;
            Beam.SetOperativePart(CP);
            if (!Beam.Insert())
                MessageBox.Show("Insert failed!");

            //bolting
            BoltArray b = new BoltArray();
            b.PartToBeBolted = CP;
            b.PartToBoltTo = P1;
            b.FirstPosition = new Point(near6.X,near7.Y,((near6.Z)+(near7.Z))/2);
            b.SecondPosition = new Point(near5.X, near4.Y, ((near5.Z) + (near4.Z)) / 2);
            b.BoltSize = 12.7;
            b.Tolerance = 2;
            b.BoltStandard = "A325N";
            b.BoltType = BoltGroup.BoltTypeEnum.BOLT_TYPE_WORKSHOP;
            b.Length = 150;
            b.ThreadInMaterial = BoltGroup.BoltThreadInMaterialEnum.THREAD_IN_MATERIAL_NO;
            b.Position.Depth = Position.DepthEnum.MIDDLE;
            b.Position.Plane = Position.PlaneEnum.MIDDLE;
            b.Position.Rotation = Position.RotationEnum.FRONT;
            b.AddBoltDistX(127);
            b.AddBoltDistX(127);
            b.StartPointOffset.Dx = 76.2;
            b.AddBoltDistY(127);
           // b.Position.Rotation = Position.RotationEnum.BELOW;
            b.Bolt = true;

            if (!b.Insert())
            {
                MessageBox.Show("bolt is not inserted");
            }
            model.CommitChanges();
        }
    }
}
