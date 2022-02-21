using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.ApplicationServices;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.GraphicsSystem;
using Autodesk.AutoCAD.Windows;
using System.Runtime;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Threading;


// https://www.theswamp.org/index.php?topic=47340.0

// https://knowledge.autodesk.com/support/autocad/learn-explore/caas/CloudHelp/cloudhelp/2015/ENU/AutoCAD-Core/files/GUID-78E0A321-1927-4A01-B9EB-744E535870DE-htm.html

// https://help.autodesk.com/view/ACD/2015/ENU/?guid=GUID-F6B3122D-A6F4-4112-A838-4662E0E1B30B

// https://forums.autodesk.com/t5/net/bug-in-layout-getviewports/td-p/2514172

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

        void SetGrid(double size, int spacing)
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

            //ed.WriteMessage($"Current View Port Number is {currentCVPORT}\n");

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

            // Apply setting to all view ports.
            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                foreach (int viewPortNumber in viewPortNumbers)
                {
                    Application.SetSystemVariable("CVPORT", viewPortNumber);
                    //ed.WriteMessage($"Number {viewPortNumber}\n");

                    Thread.Sleep(10);

                    //Point2d snapPoint = new Point2d(0.01, 0.01);
                    //await ed.CommandAsync("setvar", "SNAPUNIT", snapPoint);
                    //await ed.CommandAsync("setvar", "GRIDUNIT", snapPoint);
                    //await ed.CommandAsync("setvar", "GRIDMAJOR", 10);

                    //ed.Command("setvar", "SNAPUNIT", snapPoint);
                    //Point2d gridPoint = new Point2d(0.01, 0.01);
                    //ed.Command("setvar", "GRIDUNIT", gridPoint);
                    //ed.Command("setvar", "GRIDMAJOR", 10);

                    //Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 0.01 0.01)) ", true, false, false);
                    //Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 0.01 0.01)) ", true, false, false);
                    //Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 10) ", true, false, false);

                    string sizeString = size.ToString() + ", " + size.ToString();

                    ed.Command(new object[] { "setvar", "SNAPUNIT", sizeString, "" });
                    ed.Command(new object[] { "setvar", "GRIDUNIT", sizeString, "" });
                    ed.Command(new object[] { "setvar", "GRIDMAJOR", spacing.ToString(), "" });
                }

                transaction.Commit();
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

        /// <summary>
        /// Method to execute test 1.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("Test1")]
        public void Test1()
        {
            // Get AutoCAD objects.
            Document doc = Application.DocumentManager.MdiActiveDocument;
            KernelDescriptor kd = new KernelDescriptor();
            //Database database = HostApplicationServices.WorkingDatabase;
            Database database = doc.Database;
            Autodesk.AutoCAD.EditorInput.Editor ed = doc.Editor;
            //Autodesk.AutoCAD.GraphicsSystem.Manager gm = doc.GraphicsManager;

            // Get the current CVPORT.
            int currentCVPORT = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));
            ed.WriteMessage($"Current View Port Number is {currentCVPORT}\n");

            // Get the other View Ports.
            List<int> viewPortNumbers = new List<int>();


            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                SymbolTable symTable = (SymbolTable)transaction.GetObject(database.ViewTableId, OpenMode.ForRead);
                foreach (ObjectId id in symTable)
                {
                    ViewTableRecord rec = (ViewTableRecord)transaction.GetObject(id, OpenMode.ForRead);

                    ed.WriteMessage($"ObjectId {rec.ObjectId} OwnerId {rec.OwnerId} ViewAssociatedToViewport {rec.ViewAssociatedToViewport}\n");

                }

                transaction.Commit();
            }


            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                ViewTable acViewTbl = (ViewTable)transaction.GetObject(database.ViewTableId, OpenMode.ForRead);
                foreach (ObjectId id in acViewTbl)
                {
                    ViewTableRecord symbol = (ViewTableRecord)transaction.GetObject(id, OpenMode.ForRead);

                    //TODO: Access to the symbol
                    ed.WriteMessage(string.Format("\nName: {0}", symbol.Name));
                }

                transaction.Commit();
            }



            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                SymbolTable symTable = (SymbolTable)transaction.GetObject(database.ViewportTableId, OpenMode.ForRead);
                foreach (ObjectId id in symTable)
                {
                    ViewportTableRecord symbol = (ViewportTableRecord)transaction.GetObject(id, OpenMode.ForRead);
                    viewPortNumbers.Add(symbol.Number);

                    ed.WriteMessage($"Number {symbol.Number} ObjectId {symbol.ObjectId} OwnerId {symbol.OwnerId}.\n");


                    



                    //using (Autodesk.AutoCAD.GraphicsSystem.Manager gm = doc.GraphicsManager)
                    //{
                    //    using (Autodesk.AutoCAD.GraphicsSystem.View vw = gm.ObtainAcGsView(symbol.Number, kd))
                    //    {
                    //        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Adding View Port number {vw.ViewportObjectId}.\n");

                    //        //    Extents3d ext = ent.GeometricExtents;
                    //        //    vw.ZoomExtents(ext.MinPoint, ext.MaxPoint);
                    //        //    vw.Zoom(zoomFactor);
                    //        //    gm.SetViewportFromView(CvId, vw, true, true, false);
                    //    }
                    //}

                }

                transaction.Commit();
            }


            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                foreach(int viewPortNumber in viewPortNumbers)
                {
                    Application.SetSystemVariable("CVPORT", viewPortNumber);

                    

                    //Application.DocumentManager.MdiActiveDocument.SendStringToExecute("\033 " , true, false, false);
                    Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 0.5 0.5)) ", true, false, false);
                    Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 0.5 0.5)) ", true, false, false);
                    Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 20) ", true, false, false);

                    //ed.Command();
                }

                transaction.Commit();
            }

            Application.SetSystemVariable("CVPORT", currentCVPORT);
        }

        /// <summary>
        /// Method to execute test 2.
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("Test2")]
        public void Test2()
        {
            // Get AutoCAD objects.
            Document doc = Application.DocumentManager.MdiActiveDocument;
            KernelDescriptor kd = new KernelDescriptor();
            //Database database = HostApplicationServices.WorkingDatabase;
            Database database = doc.Database;
            Autodesk.AutoCAD.EditorInput.Editor ed = doc.Editor;
            //Autodesk.AutoCAD.GraphicsSystem.Manager gm = doc.GraphicsManager;

            // Get the current CVPORT.
            int currentCVPORT = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));
            ed.WriteMessage($"Current View Port Number is {currentCVPORT}\n");

            // Get the other View Ports.
            List<int> viewPortNumbers = new List<int>();


            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                SymbolTable symTable = (SymbolTable)transaction.GetObject(database.ViewTableId, OpenMode.ForRead);
                foreach (ObjectId id in symTable)
                {
                    ViewTableRecord rec = (ViewTableRecord)transaction.GetObject(id, OpenMode.ForRead);

                    ed.WriteMessage($"ObjectId {rec.ObjectId} OwnerId {rec.OwnerId} ViewAssociatedToViewport {rec.ViewAssociatedToViewport}\n");

                }

                transaction.Commit();
            }


            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                ViewTable acViewTbl = (ViewTable)transaction.GetObject(database.ViewTableId, OpenMode.ForRead);
                foreach (ObjectId id in acViewTbl)
                {
                    ViewTableRecord symbol = (ViewTableRecord)transaction.GetObject(id, OpenMode.ForRead);

                    //TODO: Access to the symbol
                    ed.WriteMessage(string.Format("\nName: {0}", symbol.Name));
                }

                transaction.Commit();
            }



            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                SymbolTable symTable = (SymbolTable)transaction.GetObject(database.ViewportTableId, OpenMode.ForRead);
                foreach (ObjectId id in symTable)
                {
                    ViewportTableRecord symbol = (ViewportTableRecord)transaction.GetObject(id, OpenMode.ForRead);
                    viewPortNumbers.Add(symbol.Number);

                    ed.WriteMessage($"Number {symbol.Number} ObjectId {symbol.ObjectId} OwnerId {symbol.OwnerId}.\n");






                    //using (Autodesk.AutoCAD.GraphicsSystem.Manager gm = doc.GraphicsManager)
                    //{
                    //    using (Autodesk.AutoCAD.GraphicsSystem.View vw = gm.ObtainAcGsView(symbol.Number, kd))
                    //    {
                    //        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Adding View Port number {vw.ViewportObjectId}.\n");

                    //        //    Extents3d ext = ent.GeometricExtents;
                    //        //    vw.ZoomExtents(ext.MinPoint, ext.MaxPoint);
                    //        //    vw.Zoom(zoomFactor);
                    //        //    gm.SetViewportFromView(CvId, vw, true, true, false);
                    //    }
                    //}

                }

                transaction.Commit();
            }


            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                foreach (int viewPortNumber in viewPortNumbers)
                {
                    Application.SetSystemVariable("CVPORT", viewPortNumber);

                    //Application.DocumentManager.MdiActiveDocument.SendStringToExecute("\033 ", true, false, false);
                    Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"SNAPUNIT\" (List 0.1 0.1)) ", true, false, false);
                    Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDUNIT\" (List 0.1 0.1)) ", true, false, false);
                    Application.DocumentManager.MdiActiveDocument.SendStringToExecute("(setvar \"GRIDMAJOR\" 10) ", true, false, false);

                    //ed.Command();
                }

                transaction.Commit();
            }

            Application.SetSystemVariable("CVPORT", currentCVPORT);
        }

        ///// <summary>
        ///// Method to execute test 1.
        ///// </summary>
        //[Autodesk.AutoCAD.Runtime.CommandMethod("Test1")]
        //public void Test1()
        //{
        //    //Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Test 1 ", true, false, false);

        //    Document doc = Application.DocumentManager.MdiActiveDocument;

        //    Database db = HostApplicationServices.WorkingDatabase;

        //    Autodesk.AutoCAD.EditorInput.Editor ed = doc.Editor;

        //    Autodesk.AutoCAD.GraphicsSystem.Manager gm = doc.GraphicsManager;

        //    //ObjectIdCollection vpIds = db.GetViewports(false);

        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        DBDictionary layoutmgr = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

        //        foreach (DBDictionaryEntry entry in layoutmgr)
        //        {

        //            Layout layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;

        //            ObjectIdCollection vids = layout.GetViewports();

        //            if (vids.Count > 0)
        //            {
        //                //ObjectId vid = vids[0];

        //                //Viewport pvport = tr.GetObject(vid, OpenMode.ForWrite, false, false) as Viewport;

        //                //ed.WriteMessage($"Name = {layout.LayoutName} {vid.ToString()}\n");


        //                foreach (ObjectId oi in vids)
        //                {
        //                    Viewport pvport = tr.GetObject(oi, OpenMode.ForWrite, false, false) as Viewport;

        //                    ed.WriteMessage($"Layout = {layout.LayoutName} Number = {pvport.Number} Plot Style = {pvport.PlotStyleName} {pvport.UcsName}\n");
        //                }

        //            }
        //            else
        //            {
        //                ed.WriteMessage($"No View ports found.\n");
        //            }
        //        }
        //    }

        //    ObjectIdCollection vpIds = db.GetViewports(true);


        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        DBDictionary layoutmgr = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

        //        if (vpIds.Count > 0)
        //        {

        //            foreach (ObjectId oi in vpIds)
        //            {
        //                Viewport pvport = tr.GetObject(oi, OpenMode.ForWrite, false, false) as Viewport;

        //                ed.WriteMessage($"pvport = {pvport.Number} {pvport.BlockId} {pvport.AcadObject.ToString()} {pvport.Handle}\n");
        //            }
        //        }
        //    }

        //    //DBDictionary layoutDict = (DBDictionary)db.LayoutDictionaryId.GetObject(OpenMode.ForRead);

        //    //foreach (DBDictionaryEntry de in layoutDict)
        //    //{
        //    //    Layout layout = (Layout)de.Value.GetObject(OpenMode.ForRead);

        //    //    string layoutName = layout.LayoutName;

        //    //    Application.DocumentManager.MdiActiveDocument.SendStringToExecute($"Layout Name = {layoutName}",false, false, false);
        //    //}
        //}

        //    /// <summary>
        //    /// Method to execute test 2.
        //    /// </summary>
        //    [Autodesk.AutoCAD.Runtime.CommandMethod("Test2")]
        //    public void Test2()
        //    {
        //        //Document doc = Application.DocumentManager.MdiActiveDocument;
        //        //Database db = doc.Database;
        //        //Autodesk.AutoCAD.EditorInput.Editor ed = doc.Editor;
        //        //Autodesk.AutoCAD.GraphicsSystem.Manager gm = doc.GraphicsManager;
        //        //int cvport = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));
        //        //ed.WriteMessage($"CVPORT = {cvport}\n");

        //        //Document doc = Application.DocumentManager.MdiActiveDocument;
        //        //Database db = doc.Database;
        //        //Autodesk.AutoCAD.EditorInput.Editor ed = doc.Editor;
        //        //LayoutManager layoutManager = LayoutManager.Current;


        //        //string layoutName = layoutManager.CurrentLayout;
        //        //ObjectId layoutId = layoutManager.GetLayoutId(layoutName);

        //        //ObjectId vpId = ObjectId.Null;
        //        //using (Transaction tr = db.TransactionManager.StartTransaction())
        //        //{
        //        //    Layout layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
        //        //    var vpIds = layout.GetViewports();
        //        //    if (vpIds.Count > 0)
        //        //        vpId = vpIds[1];        // First Viewport is Paperspace itself. Only one Viewport will be evaluated.
        //        //    tr.Commit();
        //        //}

        //        //Document doc = Application.DocumentManager.MdiActiveDocument;
        //        //Database db = doc.Database;
        //        //Autodesk.AutoCAD.EditorInput.Editor ed = doc.Editor;
        //        //Autodesk.AutoCAD.GraphicsSystem.Manager gm = doc.GraphicsManager;


        //        //ObjectIdCollection vpIds = db.GetViewports(true);


        //        //ed.WriteMessage($"Total number of view ports = {vpIds.Count}\n");


        //        //foreach (ObjectId oi in vpIds)
        //        //{
        //        //    ed.WriteMessage($"ID = {oi.OldIdPtr} {oi.ToString()} {oi.Handle}\n");


        //        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //        //    {
        //        //        Viewport vp = (Viewport)tr.GetObject(oi, OpenMode.ForRead);

        //        //        ed.WriteMessage($"VP = {vp.Number} {oi.OldIdPtr} {oi.Handle}\n");

        //        //    }              
        //        //}

        //        //Application.DocumentManager.MdiActiveDocument.SendStringToExecute($"Port = {cvport}", false, false, false);

        //        //Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Test 2 ", true, false, false);

        //        //Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        //        //KernelDescriptor kd = new KernelDescriptor();
        //        //kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
        //        //int CvId = Settings.Variables.CVPORT;
        //        //using (Autodesk.AutoCAD.GraphicsSystem.Manager gm = doc.GraphicsManager)
        //        //using (Autodesk.AutoCAD.GraphicsSystem.View vw = gm.ObtainAcGsView(CvId, kd))
        //        //{
        //        //    Extents3d ext = ent.GeometricExtents;
        //        //    vw.ZoomExtents(ext.MinPoint, ext.MaxPoint);
        //        //    vw.Zoom(zoomFactor);
        //        //    gm.SetViewportFromView(CvId, vw, true, true, false);

        //        //}

        //        Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        //        KernelDescriptor kd = new KernelDescriptor();
        //        Database database = HostApplicationServices.WorkingDatabase;
        //        using (Transaction transaction = database.TransactionManager.StartTransaction())
        //        {
        //            SymbolTable symTable = (SymbolTable)transaction.GetObject(database.ViewportTableId, OpenMode.ForRead);
        //            foreach (ObjectId id in symTable)
        //            {
        //                ViewportTableRecord symbol = (ViewportTableRecord)transaction.GetObject(id, OpenMode.ForRead);

        //                //TODO: Access to the symbol
        //                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\nName: {0}", symbol.Number));

        //                using (Autodesk.AutoCAD.GraphicsSystem.Manager gm = doc.GraphicsManager)
        //                {
        //                    using (Autodesk.AutoCAD.GraphicsSystem.View vw = gm.ObtainAcGsView(symbol.Number, kd))
        //                    {

        //                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\nID: {0}", vw.ViewportObjectId));

        //                    //    Extents3d ext = ent.GeometricExtents;
        //                    //    vw.ZoomExtents(ext.MinPoint, ext.MaxPoint);
        //                    //    vw.Zoom(zoomFactor);
        //                    //    gm.SetViewportFromView(CvId, vw, true, true, false);

        //                    }

        //                }

        //            }

        //            transaction.Commit();
        //        }
        //    }

    }
}