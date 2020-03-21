using System;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using System.Collections.Generic;
using Autodesk.AutoCAD.Geometry;

namespace mavCAD
{
    public class PP_Deviation : FastCAD, IExtensionApplication
    {
        // Отрисовка плановых отклонений от полилиний
        [CommandMethod("PP_Deviation")]
        public void MyCommand()
        {
            FastCAD data = new FastCAD();

            data.listLayers = MyTakeLayers();

            UserPanel window = new UserPanel(data);

            if (Application.ShowModalWindow(window) != true)
                return;

            window.textbox_deviationCreate.Text = "0.500";
            window.textbox_deviationRed.Text = "0.015";


            double toleranceRED = window._data.deviationRed;
            double toleranceCreate = window._data.deviationCreate;
            string layerBlockName = "Плановые отклонения";
            string layerPoly = window.ComboPlines.SelectedValue.ToString();
            string layerPoints = window.ComboPoints.SelectedValue.ToString();
            double blockScale = window._data.blockScale;
            string pathToBlock = @"C:\arrowDinoBlocks.dwg";


            Document doc = Application.DocumentManager.MdiActiveDocument; // получаем ссылку на активный документ
            Editor ed = doc.Editor;
            Database db = doc.Database; // получаем базу данных докумеента

            using (DocumentLock docLoc = doc.LockDocument()) // если не сделать, то в некоторых случаях приводит к возникновению ошибки при изменении базы данных документа
            using (Transaction tr = db.TransactionManager.StartTransaction()) // запускаем транзакцию
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable; // таблица блоков
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord; // пространство модели
                try
                {
                    if ((toleranceRED <= 0) || (toleranceCreate <= 0))
                    {
                        ed.WriteMessage("\n***Вы ввели неверные значения***\n");
                        return;
                    }
                    else
                    {
                        MyDeviationPP(blockScale, toleranceRED, "Плановые отклонения", toleranceCreate, layerPoly, layerPoints, @"C:\\arrowDinoBlocks.dwg");
                        
                        ed.Regen();
                        tr.Commit();
                    }

                }
                catch (ArgumentNullException)
                {
                    ed.WriteMessage("\n***Вы ввели неверные значения***\n");
                }
            }
        }

        public void Initialize()
        {
       
        }

        public void Terminate()
        {
        
        }

    }
}