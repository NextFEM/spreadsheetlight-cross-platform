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

            pic.SetFullReflection();
            // width 6pt, height 6pt
            pic.Set3DBevelBottom(DocumentFormat.OpenXml.Drawing.BevelPresetValues.Convex, 6, 6);
            // width 3pt, height 4pt
            pic.Set3DBevelTop(DocumentFormat.OpenXml.Drawing.BevelPresetValues.ArtDeco, 3, 4);
            // extrusion colour transparency 0%, extrusion (or depth) height 15 pt
            pic.Set3DExtrusion(System.Drawing.Color.Green, 0, 15);
            // contour colour transparency 40%, contour width 4pt
            pic.Set3DContour(DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent3, 40, 4);
            pic.Set3DMaterialType(DocumentFormat.OpenXml.Drawing.PresetMaterialTypeValues.TranslucentPowder);
            // 5 pt above "ground"
            pic.Set3DZDistance(5);
            // field of view 105 degrees, zoom 100%
            // camera latitude, longitude, revolution in degrees (50, 40, 30)
            // light rig latitude, longitude, revolution in degrees (0, 0, 30)
            pic.Set3DScene(DocumentFormat.OpenXml.Drawing.PresetCameraValues.PerspectiveFront, 105, 100, 50, 40, 30, DocumentFormat.OpenXml.Drawing.LightRigValues.Sunrise, DocumentFormat.OpenXml.Drawing.LightRigDirectionValues.TopLeft, 0, 0, 30);
            sl.InsertPicture(pic);

            sl.SaveAs("Pictures.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
