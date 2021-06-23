using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.ApplicationServices;
using System.IO;

[assembly: ExtensionApplication(typeof(GridSnapToolbar.GridSnap))]
[assembly: CommandClass(typeof(GridSnapToolbar.GridSnap))]
namespace GridSnapToolbar
{
    /// <summary>
    /// GridSnap class manages the toolbar and commands to quick change the grid and snap settings in AutoCAD.
    /// </summary>
    public class GridSnap : IExtensionApplication
    {
        /// <summary>
        /// AutoCAD calls this method to initialize the extension.
        /// </summary>
        public void Initialize()
        {
            AcadApplication cadApp = (AcadApplication)Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
            AcadToolbar tb = cadApp.MenuGroups.Item(0).Toolbars.Add("GridSnap");
            tb.Dock(Autodesk.AutoCAD.Interop.Common.AcToolbarDockStatus.acToolbarDockLeft);
            string basePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            AcadToolbarItem tbButton0 = tb.AddToolbarButton(0, "0.1", "Snap/Grid 0.1mm", "SnapUnit_0_1\n", null);
            tbButton0.SetBitmaps(basePath + "/tbBut_0_1.bmp", basePath + "/tbBut_0_1.bmp");

            AcadToolbarItem tbButton1 = tb.AddToolbarButton(1, "0.5", "Snap/Grid 0.5mm", "SnapUnit_0_5\n", null);
            tbButton1.SetBitmaps(basePath + "/tbBut_0_5.bmp", basePath + "/tbBut_0_5.bmp");

            AcadToolbarItem tbButton2 = tb.AddToolbarButton(2, "1.0", "Snap/Grid 1.0mm", "SnapUnit_1\n", null);
            tbButton2.SetBitmaps(basePath + "/tbBut_1.bmp", basePath + "/tbBut_1.bmp");

            AcadToolbarItem tbButton3 = tb.AddToolbarButton(3, "5.0", "Snap/Grid 5.0mm", "SnapUnit_5\n", null);
            tbButton3.SetBitmaps(basePath + "/tbBut_5.bmp", basePath + "/tbBut_5.bmp");

            AcadToolbarItem tbButton4 = tb.AddToolbarButton(4, "10.0", "Snap/Grid 10.0mm", "SnapUnit_10\n", null);
            tbButton4.SetBitmaps(basePath + "/tbBut_10.bmp", basePath + "/tbBut_10.bmp");

            AcadToolbarItem tbButton5 = tb.AddToolbarButton(5, "50.0", "Snap/Grid 50.0mm", "SnapUnit_50\n", null);
            tbButton5.SetBitmaps(basePath + "/tbBut_50.bmp", basePath + "/tbBut_50.bmp");

            AcadToolbarItem tbButton6 = tb.AddToolbarButton(6, "100.0", "Snap/Grid 100.0mm", "SnapUnit_100\n", null);
            tbButton6.SetBitmaps(basePath + "/tbBut_100.bmp", basePath + "/tbBut_100.bmp");

            AcadToolbarItem tbButton7 = tb.AddToolbarButton(7, "500.0", "Snap/Grid 500.0mm", "SnapUnit_500\n", null);
            tbButton7.SetBitmaps(basePath + "/tbBut_500.bmp", basePath + "/tbBut_500.bmp");

            AcadToolbarItem tbButton8 = tb.AddToolbarButton(8, "1000.0", "Snap/Grid 1000.0mm", "SnapUnit_1000\n", null);
            tbButton8.SetBitmaps(basePath + "/tbBut_1000.bmp", basePath + "/tbBut_1000.bmp");
        }

        /// <summary>
        /// AutoCAD calls this method to terminate the extension.
        /// </summary>
        public void Terminate()
        {
            // Nothing to terminate/dispose.
        }

        /// <summary>
        /// Method to set the grid and snap to 0.1mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_0_1")]
        public void SnapUnit_0_1_Command()
        {
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 0.1 0.1)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 0.1 0.1)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 10) ", true, false, false);
        }

        /// <summary>
        /// Method to set the grid and snap to 0.5mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_0_5")]
        public void SnapUnit_0_5_Command()
        {
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 0.5 0.5)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 0.5 0.5)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 20) ", true, false, false);
        }

        /// <summary>
        /// Method to set the grid and snap to 1mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_1")]
        public void SnapUnit_1_Command()
        {
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 1.0 1.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 1.0 1.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 10) ", true, false, false);
        }

        /// <summary>
        /// Method to set the grid and snap to 5mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_5")]
        public void SnapUnit_5_Command()
        {
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 5.0 5.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 5.0 5.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 20) ", true, false, false);
        }

        /// <summary>
        /// Method to set the grid and snap to 10mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_10")]
        public void SnapUnit_10_Command()
        {
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 10.0 10.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 10.0 10.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 10) ", true, false, false);
        }

        /// <summary>
        /// Method to set the grid and snap to 50mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_50")]
        public void SnapUnit_50_Command()
        {
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 50.0 50.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 50.0 50.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 20) ", true, false, false);
        }

        /// <summary>
        /// Method to set the grid and snap to 100mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_100")]
        public void SnapUnit_100_Command()
        {
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 100.0 100.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 100.0 100.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 10) ", true, false, false);
        }

        /// <summary>
        /// Method to set the grid and snap to 500mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_500")]
        public void SnapUnit_500_Command()
        {
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 500.0 500.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 500.0 500.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 20) ", true, false, false);
        }

        /// <summary>
        /// Method to set the grid and snap to 1000mm. (1m)
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_1000")]
        public void SnapUnit_1000_Command()
        {
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 1000.0 1000.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 1000.0 1000.0)) ", true, false, false);
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 10) ", true, false, false);
        }
    }
}