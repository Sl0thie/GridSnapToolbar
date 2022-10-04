using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.ApplicationServices;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.GraphicsSystem;
using System.Collections.Generic;
using System.Threading;
using Serilog;

// http://docs.autodesk.com/ACD/2010/ENU/AutoCAD%20.NET%20Developer%27s%20Guide/index.html?url=WS1a9193826455f5ff2566ffd511ff6f8c7ca-4363.htm,topicNumber=d0e5006&_ga=2.108075330.1797667990.1645396265-1467573342.1641428789
// https://help.autodesk.com/view/OARX/2022/ENU/?guid=OARX-ManagedRefGuide-_NET_Migration_Guide

[assembly: ExtensionApplication(typeof(GridSnapToolbar.GridSnap))]
[assembly: CommandClass(typeof(GridSnapToolbar.GridSnap))]
namespace GridSnapToolbar
{
    /// <summary>
    /// GridSnap class manages the toolbar and commands to quick change the grid and snap settings in AutoCAD.
    /// </summary>
    public class GridSnap : IExtensionApplication
    {
        private readonly string basePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        private AcadToolbarItem button0_01mm;
        private AcadToolbarItem button0_05mm;
        private AcadToolbarItem button0_1mm;
        private AcadToolbarItem button0_5mm;
        private AcadToolbarItem button1mm;
        private AcadToolbarItem button5mm;
        private AcadToolbarItem button10mm;
        private AcadToolbarItem button50mm;
        private AcadToolbarItem button100mm;
        private AcadToolbarItem button500mm;
        private AcadToolbarItem button1000mm;

        /// <summary>
        /// AutoCAD calls this method to initialize the extension.
        /// </summary>
        public void Initialize()
        {
            // Setup logging for the application.
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File("GridSnap - .txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();

            Log.Information("Logging started.");

            AcadApplication cadApp = (AcadApplication)Application.AcadApplication;
            AcadToolbar tb = cadApp.MenuGroups.Item(0).Toolbars.Add("GridSnap");
            tb.Dock(Autodesk.AutoCAD.Interop.Common.AcToolbarDockStatus.acToolbarDockLeft);

            // Prefixing \u0003\u0003 to the macro will cancel any others already in progress.
            button0_01mm = tb.AddToolbarButton(0, "0.01", "Sets the Snap/Grid 0.01mm", "\u0003\u0003SnapUnit_0_01\n", null);
            button0_05mm = tb.AddToolbarButton(1, "0.05", "Sets the Snap/Grid 0.05mm", "\u0003\u0003SnapUnit_0_05\n", null);
            button0_1mm = tb.AddToolbarButton(2, "0.1", "Sets the Snap/Grid 0.1mm", "\u0003\u0003SnapUnit_0_1\n", null);
            button0_5mm = tb.AddToolbarButton(3, "0.5", "Sets the Snap/Grid 0.5mm", "\u0003\u0003SnapUnit_0_5\n", null);
            button1mm = tb.AddToolbarButton(4, "1", "Sets the Snap/Grid 1mm", "\u0003\u0003SnapUnit_1\n", null);
            button5mm = tb.AddToolbarButton(5, "5", "Sets the Snap/Grid 5mm", "\u0003\u0003SnapUnit_5\n", null);
            button10mm = tb.AddToolbarButton(6, "10", "Sets the Snap/Grid 10mm", "\u0003\u0003SnapUnit_10\n", null);
            button50mm = tb.AddToolbarButton(7, "50", "Sets the Snap/Grid 50mm", "\u0003\u0003SnapUnit_50\n", null);
            button100mm = tb.AddToolbarButton(8, "100", "Sets the Snap/Grid 100mm", "\u0003\u0003SnapUnit_100\n", null);
            button500mm = tb.AddToolbarButton(9, "500", "Sets the Snap/Grid 500mm", "\u0003\u0003SnapUnit_500\n", null);
            button1000mm = tb.AddToolbarButton(10, "1000", "Sets the Snap/Grid 1000mm", "\u0003\u0003SnapUnit_1000\n", null);

            SetBitMaps();
        }

        /// <summary>
        /// Sets the bitmaps for the toolbar items.
        /// </summary>
        private void SetBitMaps()
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

        private void SetGrid(double size, short spacing)
        {
            // Get AutoCAD objects.
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database database = doc.Database;

            // Get the current CVPORT.
            int currentCVPORT = System.Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("CVPORT"));

            // Get the other View Ports.
            List<int> viewPortNumbers = new List<int>();
            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                SymbolTable symTable = (SymbolTable)transaction.GetObject(database.ViewportTableId, OpenMode.ForRead);
                foreach (ObjectId id in symTable)
                {
                    ViewportTableRecord symbol = (ViewportTableRecord)transaction.GetObject(id, OpenMode.ForWrite);
                    viewPortNumbers.Add(symbol.Number);

                    Log.Information("Viewport Before");
                    Log.Information($"Id {symbol.Id}");
                    Log.Information($"Number {symbol.Number}");
                    Log.Information($"CenterPoint {symbol.CenterPoint.X} {symbol.CenterPoint.Y}");
                    Log.Information($"GridAdaptive {symbol.GridAdaptive}");
                    Log.Information($"GridBoundToLimits {symbol.GridBoundToLimits}");
                    Log.Information($"GridEnabled {symbol.GridEnabled}");
                    Log.Information($"GridFollow {symbol.GridFollow}");
                    Log.Information($"GridIncrements {symbol.GridIncrements.X} {symbol.GridIncrements.Y}");
                    Log.Information($"GridMajor {symbol.GridMajor}");
                    Log.Information($"GridSubdivisionRestricted {symbol.GridSubdivisionRestricted}");
                    Log.Information($"Height {symbol.Height}");
                    Log.Information($"UcsName {symbol.UcsName}");
                    Log.Information($"UcsOrthographic {symbol.UcsOrthographic}");
                    Log.Information($"ViewDirection {symbol.ViewDirection.X} {symbol.ViewDirection.Y} {symbol.ViewDirection.Z}");
                    Log.Information($"Width {symbol.Width}");

                    symbol.GridEnabled = true;
                    symbol.GridIncrements = new Autodesk.AutoCAD.Geometry.Point2d(size,size);
                    symbol.GridMajor = spacing;
                    symbol.SnapIncrements = new Autodesk.AutoCAD.Geometry.Point2d(size , size);

                    Log.Information("Viewport After");
                    Log.Information($"Id {symbol.Id}");
                    Log.Information($"Number {symbol.Number}");
                    Log.Information($"CenterPoint {symbol.CenterPoint.X} {symbol.CenterPoint.Y}");
                    Log.Information($"GridAdaptive {symbol.GridAdaptive}");
                    Log.Information($"GridBoundToLimits {symbol.GridBoundToLimits}");
                    Log.Information($"GridEnabled {symbol.GridEnabled}");
                    Log.Information($"GridFollow {symbol.GridFollow}");
                    Log.Information($"GridIncrements {symbol.GridIncrements.X} {symbol.GridIncrements.Y}");
                    Log.Information($"GridMajor {symbol.GridMajor}");
                    Log.Information($"GridSubdivisionRestricted {symbol.GridSubdivisionRestricted}");
                    Log.Information($"Height {symbol.Height}");
                    Log.Information($"UcsName {symbol.UcsName}");
                    Log.Information($"UcsOrthographic {symbol.UcsOrthographic}");
                    Log.Information($"ViewDirection {symbol.ViewDirection.X} {symbol.ViewDirection.Y} {symbol.ViewDirection.Z}");
                    Log.Information($"Width {symbol.Width}\n");
                }

                transaction.Commit();

                doc.Editor.UpdateTiledViewportsFromDatabase();
            }

            // Switch back to original view port.
            Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("CVPORT", currentCVPORT);
        }

        /// <summary>
        /// Method to set the grid and snap to 0.01mm.
        /// </summary>
        [CommandMethod("SNAPUNIT_0_01", CommandFlags.Modal)]
        public void SnapUnit_0_01_Command()
        {
            SetGrid(0.01, 10);
        }

        /// <summary>
        /// Method to set the grid and snap to 0.05mm.
        /// </summary>
        [CommandMethod("SNAPUNIT_0_05", CommandFlags.Modal)]
        public void SnapUnit_0_05_Command()
        {
            SetGrid(0.05, 20);
        }

        /// <summary>
        /// Method to set the grid and snap to 0.1mm.
        /// </summary>
        [CommandMethod("SNAPUNIT_0_1", CommandFlags.Modal)]
        public void SnapUnit_0_1_Command()
        {
            SetGrid(0.1, 10);
        }

        /// <summary>
        /// Method to set the grid and snap to 0.5mm.
        /// </summary>
        [CommandMethod("SNAPUNIT_0_5", CommandFlags.Modal)]
        public void SnapUnit_0_5_Command()
        {
            SetGrid(0.5, 20);
        }

        /// <summary>
        /// Method to set the grid and snap to 1mm.
        /// </summary>
        [CommandMethod("SNAPUNIT_1", CommandFlags.Modal)]
        public void SnapUnit_1_Command()
        {
            SetGrid(1, 10);
        }

        /// <summary>
        /// Method to set the grid and snap to 5mm.
        /// </summary>
        [CommandMethod("SNAPUNIT_5", CommandFlags.Modal)]
        public void SnapUnit_5_Command()
        {
            SetGrid(5, 20);
        }

        /// <summary>
        /// Method to set the grid and snap to 10mm.
        /// </summary>
        [CommandMethod("SNAPUNIT_10", CommandFlags.Modal)]
        public void SnapUnit_10_Command()
        {
            SetGrid(10, 10);
        }

        /// <summary>
        /// Method to set the grid and snap to 50mm.
        /// </summary>
        [CommandMethod("SNAPUNIT_50", CommandFlags.Modal)]
        public void SnapUnit_50_Command()
        {
            SetGrid(50, 20);
        }

        /// <summary>
        /// Method to set the grid and snap to 100mm.
        /// </summary>
        [CommandMethod("SNAPUNIT_100", CommandFlags.Modal)]
        public void SnapUnit_100_Command()
        {
            SetGrid(100, 10);
        }

        /// <summary>
        /// Method to set the grid and snap to 500mm.
        /// </summary>
        [CommandMethod("SNAPUNIT_500", CommandFlags.Modal)]
        public void SnapUnit_500_Command()
        {
            SetGrid(500, 20);
        }

        /// <summary>
        /// Method to set the grid and snap to 1000mm. (1m)
        /// </summary>
        [CommandMethod("SNAPUNIT_1000", CommandFlags.Modal)]
        public void SnapUnit_1000_Command()
        {
            SetGrid(100, 10);
        }
    }
}