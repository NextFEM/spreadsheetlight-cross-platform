using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Drawing;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            SLDocument sl = new SLDocument();

            SLPicture pic = new SLPicture("julia.png");
            // set the top of the picture to be halfway in row 3
            // and the left of the picture to be halfway in column 1
            pic.SetPosition(2.5, 0.5);
            sl.InsertPicture(pic);

            pic = new SLPicture("randomclouds.jpg");
            // set the top of the picture to be flush with the bottom of row 1
            // and the left of the picture to be flush with the right of column 6
            pic.SetPosition(1, 6);
            sl.InsertPicture(pic);

            pic = new SLPicture("mandelbrot.png");
            pic.SetPosition(12, 3);

			// now for something really fancy.

			pic.Reflection.SetFullReflection();
			// width 6pt, height 6pt
			pic.Format3D.SetBevelBottom(DocumentFormat.OpenXml.Drawing.BevelPresetValues.Convex, 6, 6);
            // width 3pt, height 4pt
            pic.Format3D.SetBevelTop(DocumentFormat.OpenXml.Drawing.BevelPresetValues.ArtDeco, 3, 4);
            // extrusion colour, extrusion (or depth) height 15 pt
            pic.Format3D.SetExtrusion(System.Drawing.Color.Green, 15);
            // contour colour tint 0%, contour width 4pt
            pic.Format3D.SetContour(SLThemeColorIndexValues.Accent3Color, 0, 4);
            pic.Format3D.Material = DocumentFormat.OpenXml.Drawing.PresetMaterialTypeValues.TranslucentPowder;
            // 5 pt above "ground"
            pic.Rotation3D.DistanceZ = 5;
			// field of view 105 degrees
			// camera latitude, longitude, revolution in degrees (50, 40, 30)
			// light rig revolution in degrees 30
			pic.Rotation3D.SetCameraPreset(SLCameraPresetValues.PerspectiveFront);
			pic.Rotation3D.Perspective = 105;
			pic.Rotation3D.Y = 50;
			pic.Rotation3D.X = 40;
			pic.Rotation3D.Z = 30;
			pic.Format3D.Lighting = DocumentFormat.OpenXml.Drawing.LightRigValues.Sunrise;
			pic.Format3D.Angle = 30;
			sl.InsertPicture(pic);

            sl.SaveAs("Pictures.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
