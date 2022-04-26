using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.ApplicationServices;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.GraphicsSystem;
using System.Collections.Generic;
using System.Threading;

// http://docs.autodesk.com/ACD/2010/ENU/AutoCAD%20.NET%20Developer%27s%20Guide/index.html?url=WS1a9193826455f5ff2566ffd511ff6f8c7ca-4363.htm,topicNumber=d0e5006&_ga=2.108075330.1797667990.1645396265-1467573342.1641428789


[assembly: ExtensionApplication(typeof(GridSnapToolbar.GridSnap))]
[assembly: CommandClass(typeof(GridSnapToolbar.GridSnap))]
namespace GridSnapToolbar
{
    /// <summary>
    /// GridSnap class manages the toolbar and commands to quick change the grid and snap settings in AutoCAD.
    /// </summary>
    public class GridSnap : IExtensionApplication
    {
        string basePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        AcadToolbarItem button0_01mm;
        AcadToolbarItem button0_05mm;
        AcadToolbarItem button0_1mm;
        AcadToolbarItem button0_5mm;
        AcadToolbarItem button1mm;
        AcadToolbarItem button5mm;
        AcadToolbarItem button10mm;
        AcadToolbarItem button50mm;
        AcadToolbarItem button100mm;
        AcadToolbarItem button500mm;
        AcadToolbarItem button1000mm;

        /// <summary>
        /// AutoCAD calls this method to initialize the extension.
        /// </summary>
        public void Initialize()
        {
            AcadApplication cadApp = (AcadApplication)Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
            AcadToolbar tb = cadApp.MenuGroups.Item(0).Toolbars.Add("GridSnap");
            tb.Dock(Autodesk.AutoCAD.Interop.Common.AcToolbarDockStatus.acToolbarDockLeft);
            
            button0_01mm = tb.AddToolbarButton(0, "0.01", "Sets the Snap/Grid 0.01mm", "SnapUnit_0_01\n", null);
            button0_05mm = tb.AddToolbarButton(1, "0.05", "Sets the Snap/Grid 0.05mm", "SnapUnit_0_05\n", null);
            button0_1mm = tb.AddToolbarButton(2, "0.1", "Sets the Snap/Grid 0.1mm", "SnapUnit_0_1\n", null);
            button0_5mm = tb.AddToolbarButton(3, "0.5", "Sets the Snap/Grid 0.5mm", "SnapUnit_0_5\n", null);
            button1mm = tb.AddToolbarButton(4, "1", "Sets the Snap/Grid 1mm", "SnapUnit_1\n", null);
            button5mm = tb.AddToolbarButton(5, "5", "Sets the Snap/Grid 5mm", "SnapUnit_5\n", null);
            button10mm = tb.AddToolbarButton(6, "10", "Sets the Snap/Grid 10mm", "SnapUnit_10\n", null);
            button50mm = tb.AddToolbarButton(7, "50", "Sets the Snap/Grid 50mm", "SnapUnit_50\n", null);
            button100mm = tb.AddToolbarButton(8, "100", "Sets the Snap/Grid 100mm", "SnapUnit_100\n", null);
            button500mm = tb.AddToolbarButton(9, "500", "Sets the Snap/Grid 500mm", "SnapUnit_500\n", null);
            button1000mm = tb.AddToolbarButton(10, "1000", "Sets the Snap/Grid 1000mm", "SnapUnit_1000\n", null);

            SetBitMaps();
        }

        /// <summary>
        /// Sets the bitmaps for the toolbar items.
        /// </summary>
        void SetBitMaps()
        {
            button0_01mm.SetBitmaps(basePath + "/tbBut_0_01.bmp", basePath + "/tbBut_0_01.bmp");
            button0_05mm.SetBitmaps(basePath + "/tbBut_0_05.bmp", basePath + "/tbBut_0_05.bmp");
            button0_1mm.SetBitmaps(basePath + "/tbBut_0_1.bmp", basePath + "/tbBut_0_1.bmp");
            button0_5mm.SetBitmaps(basePath + "/tbBut_0_5.bmp", basePath + "/tbBut_0_5.bmp");
            button1mm.SetBitmaps(basePath + "/tbBut_1.bmp", basePath + "/tbBut_1.bmp");
            button5mm.SetBitmaps(basePath + "/tbBut_5.bmp", basePath + "/tbBut_5.bmp");
            button10mm.SetBitmaps(basePath + "/tbBut_10.bmp", basePath + "/tbBut_10.bmp");
            button50mm.SetBitmaps(basePath + "/tbBut_50.bmp", basePath + "/tbBut_50.bmp");
            button100mm.SetBitmaps(basePath + "/tbBut_100.bmp", basePath + "/tbBut_100.bmp");
            button500mm.SetBitmaps(basePath + "/tbBut_500.bmp", basePath + "/tbBut_500.bmp");
            button1000mm.SetBitmaps(basePath + "/tbBut_1000.bmp", basePath + "/tbBut_1000.bmp");
        }
    
        /// <summary>
        /// AutoCAD calls this method to terminate the extension.
        /// </summary>
        public void Terminate()
        {
            // Nothing to terminate/dispose.
        }

        void SetGrid(double size, short spacing)
        {
            // Get AutoCAD objects.
            Document doc = Application.DocumentManager.MdiActiveDocument;
            KernelDescriptor kd = new KernelDescriptor();
            Database database = doc.Database;
            Autodesk.AutoCAD.EditorInput.Editor ed = doc.Editor;

            // Get the current CVPORT.
            int currentCVPORT = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));

            // Get the other View Ports.
            List<int> viewPortNumbers = new List<int>();
            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                SymbolTable symTable = (SymbolTable)transaction.GetObject(database.ViewportTableId, OpenMode.ForRead);
                foreach (ObjectId id in symTable)
                {
                    ViewportTableRecord symbol = (ViewportTableRecord)transaction.GetObject(id, OpenMode.ForWrite);
                    viewPortNumbers.Add(symbol.Number);

                    symbol.GridEnabled = true;

                    symbol.GridIncrements = new Autodesk.AutoCAD.Geometry.Point2d(size,size);

                    symbol.GridMajor = spacing;

                    symbol.SnapIncrements = new Autodesk.AutoCAD.Geometry.Point2d(size , size);
                }
                transaction.Commit();

                doc.Editor.UpdateTiledViewportsFromDatabase();
            }

            // Switch back to original view port.
            Application.SetSystemVariable("CVPORT", currentCVPORT);
        }

        void SetGridOld(double size, int spacing)
        {
            // Get AutoCAD objects.
            Document doc = Application.DocumentManager.MdiActiveDocument;
            KernelDescriptor kd = new KernelDescriptor();
            Database database = doc.Database;
            Autodesk.AutoCAD.EditorInput.Editor ed = doc.Editor;

            // Get the current CVPORT.
            int currentCVPORT = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));
            int currentCMDECHO = System.Convert.ToInt32(Application.GetSystemVariable("CMDECHO"));
            if (currentCMDECHO != 0)
            {
                Application.SetSystemVariable("CMDECHO", 0);
            }

            ed.WriteMessage($"Grid Size {size} Spacing {spacing}\n");

            //ed.Command(new object[] { "\x03\x03", "" });
            //doc.SendStringToExecute("\x03\x03", false, true, false);

            // Get the other View Ports.
            List<int> viewPortNumbers = new List<int>();
            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                SymbolTable symTable = (SymbolTable)transaction.GetObject(database.ViewportTableId, OpenMode.ForRead);
                foreach (ObjectId id in symTable)
                {
                    ViewportTableRecord symbol = (ViewportTableRecord)transaction.GetObject(id, OpenMode.ForRead);
                    viewPortNumbers.Add(symbol.Number);

                }
                transaction.Commit();
            }

            //// Apply setting to all view ports.
            //using (Transaction transaction = database.TransactionManager.StartTransaction())
            //{
            //    foreach (int viewPortNumber in viewPortNumbers)
            //    {
            //        Application.SetSystemVariable("CVPORT", viewPortNumber);
            //        Thread.Sleep(10);
            //        string sizeString = size.ToString() + ", " + size.ToString();
            //        ed.Command(new object[] { "setvar", "SNAPUNIT", sizeString, "" });
            //        ed.Command(new object[] { "setvar", "GRIDUNIT", sizeString, "" });
            //        ed.Command(new object[] { "setvar", "GRIDMAJOR", spacing.ToString(), "" });
            //    }

            //    transaction.Commit();
            //}

            // Apply setting to all view ports.
            foreach (int viewPortNumber in viewPortNumbers)
            {
                Application.SetSystemVariable("CVPORT", viewPortNumber);

                

                Thread.Sleep(100);
                string sizeString = size.ToString("#.00") + ", " + size.ToString("#");
                ed.Command(new object[] { "._setvar", "SNAPUNIT", sizeString, "" });
                Thread.Sleep(100);

                ed.Command(new object[] { "._setvar", "GRIDUNIT", sizeString, "" });
                Thread.Sleep(100);

                ed.Command(new object[] { "._setvar", "GRIDMAJOR", spacing.ToString(), "" });
                Thread.Sleep(100);
            }

            // Switch back to original view port.
            Application.SetSystemVariable("CVPORT", currentCVPORT);
            Application.SetSystemVariable("CMDECHO", currentCMDECHO);
        }

        /// <summary>
        /// Method to set the grid and snap to 0.01mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_0_01")]
        public void SnapUnit_0_01_Command()
        {
            SetGrid(0.01, 10);
        }

        /// <summary>
        /// Method to set the grid and snap to 0.05mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_0_05")]
        public void SnapUnit_0_05_Command()
        {
            SetGrid(0.05, 20);
        }

        /// <summary>
        /// Method to set the grid and snap to 0.1mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_0_1")]
        public void SnapUnit_0_1_Command()
        {
            SetGrid(0.1, 10);
        }

        /// <summary>
        /// Method to set the grid and snap to 0.5mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_0_5")]
        public void SnapUnit_0_5_Command()
        {
            SetGrid(0.5, 20);
        }

        /// <summary>
        /// Method to set the grid and snap to 1mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_1")]
        public void SnapUnit_1_Command()
        {
            SetGrid(1, 10);
        }

        /// <summary>
        /// Method to set the grid and snap to 5mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_5")]
        public void SnapUnit_5_Command()
        {
            SetGrid(5, 20);
        }

        /// <summary>
        /// Method to set the grid and snap to 10mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_10")]
        public void SnapUnit_10_Command()
        {
            SetGrid(10, 10);
        }

        /// <summary>
        /// Method to set the grid and snap to 50mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_50")]
        public void SnapUnit_50_Command()
        {
            SetGrid(50, 20);
        }

        /// <summary>
        /// Method to set the grid and snap to 100mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_100")]
        public void SnapUnit_100_Command()
        {
            SetGrid(100, 10);
        }

        /// <summary>
        /// Method to set the grid and snap to 500mm.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_500")]
        public void SnapUnit_500_Command()
        {
            SetGrid(500, 20);
        }

        /// <summary>
        /// Method to set the grid and snap to 1000mm. (1m)
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("SNAPUNIT_1000")]
        public void SnapUnit_1000_Command()
        {
            SetGrid(100, 10);
        }
    }
}