using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using App = Autodesk.AutoCAD.ApplicationServices;
using AcDb = Autodesk.AutoCAD.DatabaseServices;

namespace mavCAD
{
    public class FastCAD : IExtensionApplication
    {
        public string selectedLayerPlines { get; set; }
        public string selectedLayerPoints { get; set; }
        public List<string> listLayers { get; set; }
        public double deviationRed { get; set; }
        public double deviationRedAngle { get; set; }
        public double deviationCreate { get; set; }
        public double blockScale { get; set; }



        // Метод создания слоя 
        protected void MyNewLayer(string _name)
        {
            string name = _name;
            // получаем ссылку на активный документ
            Document adoc = Application.DocumentManager.MdiActiveDocument;
            if (adoc == null) // если открыта начальная страница (док не активен=0), то выходим из метода
                return;

            Database db = adoc.Database; // получаем базу данных докумеента
            ObjectId layerTableId = db.LayerTableId; // получаем Id таблицы слоев

            using (DocumentLock docLoc = adoc.LockDocument()) // если не сделать, то в некоторых случаях приводит к возникновению ошибки при изменении базы данных документа

            using (Transaction tr = db.TransactionManager.StartTransaction()) // запускаем транзакцию
            {
                // открываем таблицу блоков документа (чтобы добраться до пространства модели)
                BlockTable acBlockTable;
                acBlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                // открываем пространство модели (Model Space) - оно является одной из записей в таблице блоков документа
                BlockTableRecord acBlockTableRec;
                acBlockTableRec = tr.GetObject(acBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // создаем новый слой
                // открываем таблицу слоев документа
                LayerTable acLayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;

                // создаем сам слой и присваиваем имя
                LayerTableRecord acLayerTableRecord = new LayerTableRecord();

                // проверяем нет ли уже слоя с таким именем
                if (acLayerTable.Has(name) == false)
                {
                    acLayerTableRecord.Name = name;

                    // заносим созданный слой таблицу слоев
                    acLayerTable.Add(acLayerTableRecord);

                    // добавляем созданный слой в документ
                    tr.AddNewlyCreatedDBObject(acLayerTableRecord, true);
                }

                // проверка если нет нашего слоя - прекращаем выполнение проги
                if (acLayerTable.Has(name) == false)
                {
                    return;
                }

                // db.Clayer = acLayerTable[name]; // устанавливаем новый слой как текущий

                tr.Commit(); // подтвержение выполнения транзакции, если не поставить то код не выполнится
            }

        }

        // Перегрузка метода создания слоя имя+цвет слоя
        protected void MyNewLayer(string _name, byte _r, byte _g, byte _b)
        {
            byte r = _r;
            byte g = _g;
            byte b = _b;
            string name = _name;
            // получаем ссылку на активный документ
            Document adoc = Application.DocumentManager.MdiActiveDocument;
            if (adoc == null) // если открыта начальная страница (док не активен=0), то выходим из метода
                return;

            Database db = adoc.Database; // получаем базу данных докумеента
            ObjectId layerTableId = db.LayerTableId; // получаем Id таблицы слоев

            using (DocumentLock docLoc = adoc.LockDocument()) // если не сделать, то в некоторых случаях приводит к возникновению ошибки при изменении базы данных документа

            using (Transaction tr = db.TransactionManager.StartTransaction()) // запускаем транзакцию
            {
                // открываем таблицу блоков документа (чтобы добраться до пространства модели)
                BlockTable acBlockTable;
                acBlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                // открываем пространство модели (Model Space) - оно является одной из записей в таблице блоков документа
                BlockTableRecord acBlockTableRec;
                acBlockTableRec = tr.GetObject(acBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // создаем новый слой
                // открываем таблицу слоев документа
                LayerTable acLayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;

                // создаем сам слой и присваиваем имя
                LayerTableRecord acLayerTableRecord = new LayerTableRecord();

                // проверяем нет ли уже слоя с таким именем
                if (acLayerTable.Has(name) == false)
                {
                    acLayerTableRecord.Name = name;
                    acLayerTableRecord.Color = Color.FromRgb(r, g, b);
                    // заносим созданный слой таблицу слоев
                    acLayerTable.Add(acLayerTableRecord);

                    // добавляем созданный слой в документ
                    tr.AddNewlyCreatedDBObject(acLayerTableRecord, true);
                }

                // проверка если нет нашего слоя - прекращаем выполнение проги
                if (acLayerTable.Has(name) == false)
                {
                    return;
                }

                // db.Clayer = acLayerTable[name]; // устанавливаем новый слой как текущий

                tr.Commit(); // подтвержение выполнения транзакции, если не поставить то код не выполнится
            }

        }

        // Метод вставки внешнего блока
        protected void MyInsertBlock(string _path = @"C:\arrowBlocks\arrowDinoBlocks.dwg")
        {
            string path = _path;
            App.Document doc = App.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string blockname = "arrowDinoBlocks";
            using (var outsideDb = new Database(false, true))
            {
                List<BlockTableRecord> listBtr = new List<BlockTableRecord>();
                Editor ed = doc.Editor;
                outsideDb.ReadDwgFile(path, System.IO.FileShare.Read, true, blockname);
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    using (doc.LockDocument())
                    {
                        BlockTable destDbBlockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord destDbCurrentSpace = (BlockTableRecord)db.CurrentSpaceId.GetObject(OpenMode.ForWrite);
                        ObjectId sourceBlockId;
                        sourceBlockId = db.Insert(blockname, outsideDb, true);

                        foreach (ObjectId obId in bt)
                        {
                            BlockTableRecord dbObj = (BlockTableRecord)tr.GetObject(obId, OpenMode.ForWrite);
                            listBtr.Add(dbObj);
                        }

                        foreach (BlockTableRecord btr in listBtr)
                        {
                            if (btr.Name == "arrowDinoBlocks")
                                btr.Erase();
                        }


                    }
                    tr.Commit();
                }

            }
        }

        // Метод создания списка отрезков от точки до ближайшей полилинии
        protected List<Line> MyMakeClosestLines(double _tolerance, string _layerPoly, string _layerPoints, string _layerLines)
        {
            string layerLines = _layerLines; // слой выборки полилиний в чертеже
            double tolerance = _tolerance; // максимальное расстояние при котором будут рисоваться отрезки
            string layerPoly = _layerPoly; // слой выборки кратчайших отрезков в чертеже
            string layerPoints = _layerPoints; // слой выборки точек в чертеже

            //tolerance = 0.030; // 
            List<Line> ClosestLineList = new List<Line>();
            // переменная для выбора слоя с полилиниями string setLayerPoly, string setLayerPoints
            //layerPoly = "Стены";
            // переменная для выбора слоя с точками
            //layerPoints = "Точки";

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // создаем переменную, в которой будут содержаться данные для фильтра
            TypedValue[] filterPoly = new TypedValue[2];

            // первый аргумент (0) указывает, что мы задаем тип объекта 
            // второй аргумент ("POINT") - собственно тип     /// LWPOLYLINE - упрощенная полилиния, исп-ся везде; POLYLINE - 2Д полилиния которая грузит систему, была в старых версиях
            filterPoly[0] = new TypedValue((int)DxfCode.Start, "LWPOLYLINE"); // условие выбора чтоб это была полилиния
            filterPoly[1] = new TypedValue((int)DxfCode.LayerName, layerPoly); // И чтобы находилась в слое "Стены"

            // создаем фильтр 1й выборки
            SelectionFilter selectionFilterPoly = new SelectionFilter(filterPoly);

            // пытаемся получить ссылки на объекты с учетом фильтра
            // ВНИМАНИЕ! Нужно проверить работоспособность метода с замороженными и заблокированными слоями!
            PromptSelectionResult selResPoly = ed.SelectAll(selectionFilterPoly);

            // если произошла ошибка - сообщаем о ней
            if (selResPoly.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nВ выбранном слое отсутствуют полилинии!\n");
                return null;
            }

            // получаем массив ID объектов
            ObjectId[] idsPoly = selResPoly.Value.GetObjectIds();

            // создаем переменную, в которой будут содержаться данные для фильтра
            TypedValue[] filterPoints = new TypedValue[2];

            // первый аргумент (0) указывает, что мы задаем тип объекта
            // второй аргумент ("POINT") - собственно тип
            filterPoints[0] = new TypedValue((int)DxfCode.Start, "POINT"); // условие выбора чтоб это была точка
            filterPoints[1] = new TypedValue((int)DxfCode.LayerName, layerPoints); // И чтобы находилась в слое "Точки"

            // создаем фильтр 2й выборки
            SelectionFilter selectionFilterPoints = new SelectionFilter(filterPoints);

            // пытаемся получить ссылки на объекты с учетом фильтра
            // ВНИМАНИЕ! Нужно проверить работоспособность метода с замороженными и заблокированными слоями!
            PromptSelectionResult selResPoints = ed.SelectAll(selectionFilterPoints);

            // если произошла ошибка - сообщаем о ней
            if (selResPoints.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nВ выбранном слое отсутствуют точки!\n");
                return null;
            }

            // получаем массив ID объектов
            ObjectId[] idsPoints = selResPoints.Value.GetObjectIds();

            // начинаем транзакцию
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {

                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // создаем списки с полилиниями и точками
                List<Polyline> listPlines = new List<Polyline>();
                List<DBPoint> listPoints = new List<DBPoint>();

                foreach (ObjectId idPoly in idsPoly) // заполняем список полилиниями из набора, приобразовав их в тип Polyline
                {
                    Polyline acPoly = (Polyline)tr.GetObject(idPoly, OpenMode.ForRead);

                   // if (acPoly.Closed == true)
                        listPlines.Add(acPoly);
                }

                foreach (ObjectId idPoint in idsPoints) // заполняем список точками из набора, приобразовав их в тип DBPoint
                {
                    DBPoint acPoint = (DBPoint)tr.GetObject(idPoint, OpenMode.ForRead);
                    listPoints.Add(acPoint);
                }

                int ii = 0;
                int jj = 1;
                for (int i = 0; i < listPoints.Count; i++) // перебираем точки
                {
                    for (int j = 0; j < listPlines.Count; j++) // перебераем полилинии для каждой точки
                    {
                        Polyline acPoly = listPlines[j]; // присваиваем jтовую полилинию
                        DBPoint point = listPoints[i]; // присваиваем iтовую точку

                        // определяем координаты ближайшей точки к полилинии
                        Point3d closestPoint = acPoly.GetClosestPointTo(point.Position, false);

                        if (point.ColorIndex == 5) // Если цвет точки синий (низ)
                        {
                            // создаем полилинию где первая координата - это координаты точки, вторая - координаты ближайшей точки на полилинии
                            Line myLine = new Line(new Point3d(point.Position.X, point.Position.Y, 0), new Point3d(closestPoint.X, closestPoint.Y, 0));
                            myLine.Color = Color.FromRgb(0, 0, 5); // задаем цвет 
                            myLine.Layer = layerLines; // задаем слой 
                            if (myLine.Length < tolerance)
                            {
                                ClosestLineList.Add(myLine);
                                ms.AppendEntity(myLine); // закидываем полилинию в пр-во модели
                                tr.AddNewlyCreatedDBObject(myLine, true); // закидываем в транзакцию
                            }
                        }
                        else if (point.ColorIndex == 6) // Если цвет точки фиолетовый (верх)
                        {
                            // создаем полилинию где первая координата - это координаты точки, вторая - координаты ближайшей точки на полилинии
                            Line myLine = new Line(new Point3d(point.Position.X, point.Position.Y, 0), new Point3d(closestPoint.X, closestPoint.Y, 0));
                            myLine.Color = Color.FromRgb(0, 0, 6); // задаем цвет 
                            myLine.Layer = layerLines; // задаем слой 
                            if (myLine.Length < tolerance)
                            {
                                ClosestLineList.Add(myLine);
                                ms.AppendEntity(myLine); // закидываем полилинию в пр-во модели
                                tr.AddNewlyCreatedDBObject(myLine, true); // закидываем в транзакцию
                            }
                        }
                        else
                        {
                            // создаем полилинию где первая координата - это координаты точки, вторая - координаты ближайшей точки на полилинии
                            Line myLine = new Line(new Point3d(point.Position.X, point.Position.Y, 0), new Point3d(closestPoint.X, closestPoint.Y, 0));
                            myLine.Color = Color.FromRgb(0, 0, 7); // задаем цвет 
                            myLine.Layer = layerLines; // задаем слой 
                            if (myLine.Length < tolerance)
                            {
                                ClosestLineList.Add(myLine);
                                ms.AppendEntity(myLine); // закидываем полилинию в пр-во модели
                                tr.AddNewlyCreatedDBObject(myLine, true); // закидываем в транзакцию
                            }
                        }


                    }

                }

                tr.Commit(); // подтверждаем транзакцию
            }
            return ClosestLineList;
        } // method 

        // Метод выбора полилиний в чертеже
        protected List<Polyline> MyTakePolylines(string _layerPoly)
        {
            List<Polyline> listPlines = new List<Polyline>();
            string layerPoly = _layerPoly;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // создаем переменную, в которой будут содержаться данные для фильтра
            TypedValue[] filterPoly = new TypedValue[2];

            // первый аргумент (0) указывает, что мы задаем тип объекта 
            // второй аргумент ("POINT") - собственно тип     /// LWPOLYLINE - упрощенная полилиния, исп-ся везде; POLYLINE - 2Д полилиния которая грузит систему, была в старых версиях
            filterPoly[0] = new TypedValue((int)DxfCode.Start, "LWPOLYLINE"); // условие выбора чтоб это была полилиния
            filterPoly[1] = new TypedValue((int)DxfCode.LayerName, layerPoly); // И чтобы находилась в слое "Стены"

            // создаем фильтр 1й выборки
            SelectionFilter selectionFilterPoly = new SelectionFilter(filterPoly);

            // пытаемся получить ссылки на объекты с учетом фильтра
            // ВНИМАНИЕ! Нужно проверить работоспособность метода с замороженными и заблокированными слоями!
            PromptSelectionResult selResPoly = ed.SelectAll(selectionFilterPoly);

            // если произошла ошибка - сообщаем о ней
            if (selResPoly.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nВ выбранном слое отсутствуют полилинии\n");
                return null;
            }

            // получаем массив ID объектов
            ObjectId[] idsPoly = selResPoly.Value.GetObjectIds();

            // начинаем транзакцию
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                foreach (ObjectId idPoly in idsPoly) // заполняем список полилиниями из набора, приобразовав их в тип Polyline
                {
                    Polyline acPoly = (Polyline)tr.GetObject(idPoly, OpenMode.ForRead);
                    listPlines.Add(acPoly);
                }

                tr.Commit(); // подтверждаем транзакцию
            }
            return listPlines;
        } // method 

        // Метод выбора точек в чертеже
        protected List<DBPoint> MyTakePoints(string _layerPoints)
        {
            List<DBPoint> listPoints = new List<DBPoint>();
            string layerPoints = _layerPoints;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // создаем переменную, в которой будут содержаться данные для фильтра
            TypedValue[] filterPoints = new TypedValue[2];

            // первый аргумент (0) указывает, что мы задаем тип объекта
            // второй аргумент ("POINT") - собственно тип
            filterPoints[0] = new TypedValue((int)DxfCode.Start, "POINT"); // условие выбора чтоб это была точка
            filterPoints[1] = new TypedValue((int)DxfCode.LayerName, layerPoints); // И чтобы находилась в слое "Точки"

            // создаем фильтр 2й выборки
            SelectionFilter selectionFilterPoints = new SelectionFilter(filterPoints);

            // пытаемся получить ссылки на объекты с учетом фильтра
            // ВНИМАНИЕ! Нужно проверить работоспособность метода с замороженными и заблокированными слоями!
            PromptSelectionResult selResPoints = ed.SelectAll(selectionFilterPoints);

            // если произошла ошибка - сообщаем о ней
            if (selResPoints.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nВ выбранном слое отсутствуют точки!\n");
                return null;
            }

            // получаем массив ID объектов
            ObjectId[] idsPoints = selResPoints.Value.GetObjectIds();

            // начинаем транзакцию
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                foreach (ObjectId idPoint in idsPoints) // заполняем список точками из набора, приобразовав их в тип DBPoint
                {
                    DBPoint acPoint = (DBPoint)tr.GetObject(idPoint, OpenMode.ForRead);
                    listPoints.Add(acPoint);
                }

                tr.Commit(); // подтверждаем транзакцию
            }
            return listPoints;
        } // method 

        // Метод выбора слоев в чертеже
        protected List<string> MyTakeLayers() 
        {
            Document adoc = Application.DocumentManager.MdiActiveDocument; // получаем ссылку на активный документ
            if (adoc == null) // если открыта начальная страница (док не активен=0), то выходим из метода
                return null;

            Database db = adoc.Database; // получаем базу данных докумеента

            ObjectId layerTableId = db.LayerTableId; // получаем Id таблицы слоев

            List<string> layersNames = new List<string>(); // список для сбора имен слоев

            using (Transaction tr = db.TransactionManager.StartTransaction()) // запускаем транзакцию
            {
                LayerTable layerTable = tr.GetObject(layerTableId, OpenMode.ForRead) as LayerTable; // получаем объект таблицы слоев

                foreach (ObjectId layerTableRecordId in layerTable) // проходим по таблице слоев
                {
                    LayerTableRecord layerTableRecord = tr.GetObject(layerTableRecordId, OpenMode.ForRead) as LayerTableRecord; // получаем объект - запись о слое в таблице
                    layersNames.Add(layerTableRecord.Name); // получаем имя слоя и сохраняем его в результирующий список
                }
                tr.Commit();
            }

            return layersNames;
        }

        // Метод проверки находится ли точка внутри\снаружи\на контуре
        protected bool MyIsInside(Point3d _point, Polyline _poly)
        {
            Point3d point = _point;
            Polyline poly = _poly;
            bool isInside = true;
            SystemObjects.DynamicLinker.LoadModule("AcMPolygonObj" + Application.Version.Major + ".dbx", false, false);
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //PromptPointOptions pr = new PromptPointOptions("\nУкажите точку: ");
            //PromptPointResult res = ed.GetPoint(pr);
            //if (res.Status != PromptStatus.OK) return;
            //PromptEntityOptions pre = new PromptEntityOptions("\nУкажите контур (замкнутую полилинию): ");
            //pre.SetRejectMessage("Это не контур (полилиния)");
            //pre.AddAllowedClass(typeof(Polyline), true);
            //PromptEntityResult rese = ed.GetEntity(pre);
            //if (rese.Status != PromptStatus.OK) return;
            using (MPolygon mp = new MPolygon())
            {

                using (poly.ObjectId.Open(OpenMode.ForRead) as Polyline)
                {
                    try
                    {
                        mp.AppendLoopFromBoundary(poly, true, Tolerance.Global.EqualPoint);
                        if (mp.IsPointOnLoopBoundary(point, 0, Tolerance.Global.EqualPoint))
                        {
                            isInside = true;
                            // ed.WriteMessage("\nТочка на границе контура!");
                        }
                        else if (mp.IsPointInsideMPolygon(point, Tolerance.Global.EqualPoint).Count > 0)
                        {
                            isInside = true;
                            // ed.WriteMessage("\nТочка внутри контура!");
                        }
                        else
                        {
                            isInside = false;
                            // ed.WriteMessage("\nТочка вне контура!");
                        }
                    }
                    catch { }



                }
            }
            return isInside;
        } // method

        // Метод исправляющий угол на гостовский чтоб текст читался слева направо снизу вверх
        protected double MyCorrectAngle(double _myAngleRad)
        {
            double myAngle = _myAngleRad;
            double correctAngle;


            correctAngle = (myAngle) * (180 / Math.PI);  // угол получаем в радианах, для удобного вычисления переводим в градусы
                                                         // делаем проверку на поворот отрезка, чтобы текст поворачивался по госту (чтение слева направо, снизу вверх)
            if (correctAngle > 135 && correctAngle <= 315)
            {
                correctAngle = correctAngle + 180;

                if (correctAngle >= 360)
                {
                    correctAngle = correctAngle - 360;
                }

                if (correctAngle < 0)
                {
                    correctAngle = correctAngle * -1;
                }
            }
            return correctAngle * (Math.PI / 180);
        } // method

        // Метод создания отклонений от плиты перекрытия
        protected void MyDeviationPP(double _blockScale = 0.25, double _toleranceRED = 0.020, string _layerBlockName = "Плановые отклонения", double _toleranceCreateClosestLines = 0.1, string _layerPoly = "Стены", string _layerPoints = "Точки", string _pathToBlock = @"C:\arrowBlocks\arrowDinoBlocks.dwg")
        {
            double toleranceRED = _toleranceRED;
            string layerBlockName = _layerBlockName;
            double toleranceCreateClosestLines = _toleranceCreateClosestLines;
            string layerPoly = _layerPoly;
            string layerPoints = _layerPoints;
            string pathToBlock = _pathToBlock;
            double blockScale = _blockScale;

            Document doc = Application.DocumentManager.MdiActiveDocument; // получаем ссылку на активный документ
            Editor ed = doc.Editor;
            Database db = doc.Database; // получаем базу данных докумеента

            using (DocumentLock docLoc = doc.LockDocument()) // если не сделать, то в некоторых случаях приводит к возникновению ошибки при изменении базы данных документа
            using (Transaction tr = db.TransactionManager.StartTransaction()) // запускаем транзакцию
            {

                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable; // таблица блоков
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord; // пространство модели
                MyNewLayer(layerBlockName); // создаем новый слой
                //MyNewLayer("Красный", 255,0,0);
                MyInsertBlock(pathToBlock); // всталяем наш DWG с блоками

                List<Line> closestLines = new List<Line>(MyMakeClosestLines(toleranceCreateClosestLines, layerPoly, layerPoints, layerBlockName)); // вставляем список кратчайших отрезков от точки до полилинии
                List<Polyline> selectedPolylines = new List<Polyline>(MyTakePolylines(layerPoly)); // вставляем список из полилиний из чертежа

                foreach (Line l in closestLines)
                {
                    Matrix3d curUCSMatrix = ed.CurrentUserCoordinateSystem; // создаем матрицу координат
                    CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d; // содаем текущую систему координат

                    BlockTableRecord blockDef = bt["arrowDinoBlockRPL"].GetObject(OpenMode.ForRead) as BlockTableRecord;
                    BlockReference br = new BlockReference(l.EndPoint, bt["arrowDinoBlockRPL"]);

                    ms.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    foreach (ObjectId id in blockDef) // Вставляем в блок атрибуты с отклонениями
                    {
                        DBObject obj = id.GetObject(OpenMode.ForRead);
                        AttributeDefinition attDef = obj as AttributeDefinition;
                        if ((attDef != null) && (!attDef.Constant))
                        {
                            // Это неконстантный AttributeDefinition
                            // Создаём новый AttributeReference
                            using (AttributeReference attRef = new AttributeReference())
                            {
                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                if (attRef.Tag == "ОТКЛ")
                                    attRef.TextString = string.Format("{0:0}", l.Length * 1000);
                                //else if (attRef.Tag == "ОТКЛНИЗ")
                                //    attRef.TextString = "Niz";
                                // Добавляем AttributeReference к BlockReference
                                br.AttributeCollection.AppendAttribute(attRef);
                                tr.AddNewlyCreatedDBObject(attRef, true);
                            } // using attRef
                        } // if 
                    } // foreach attribute


                    // Меняем направление стрелки в динамическом блоке
                    if ((l.Angle * (180 / Math.PI)) > 90.0 && (l.Angle * (180 / Math.PI)) <= 270.0)
                    {

                    }
                    else
                    {
                        foreach (DynamicBlockReferenceProperty prop in br.DynamicBlockReferencePropertyCollection)
                        {
                            object[] values = prop.GetAllowedValues();

                            if (prop.PropertyName == "mirror" && !prop.ReadOnly)
                            {
                                if (prop.Value.ToString() == values[0].ToString())
                                    prop.Value = values[1];

                                else
                                    prop.Value = values[0];
                            }
                        }
                    }  // Меняем направление стрелки в динамическом блоке

                    br.TransformBy(Matrix3d.Rotation((MyCorrectAngle(l.Angle)), curUCS.Zaxis, l.EndPoint)); // поворачиваем блок на вычисленный угол, переведя обратно в радианы
                    br.TransformBy(Matrix3d.Scaling(blockScale, l.EndPoint)); // изменяем масштаб блока

                    if (l.Length > toleranceRED)
                    {
                        MyNewLayer("!Вне допуска");
                        Circle c = new Circle();
                        c.Center = l.EndPoint;
                        c.Radius = 2.4 * blockScale;
                        c.Layer = "!Вне допуска";
                        c.Color = Color.FromColorIndex(ColorMethod.None, 1);
                        c.LineWeight = (LineWeight)30;
                        ms.AppendEntity(c);
                        tr.AddNewlyCreatedDBObject(c, true);
                        br.Color = Color.FromColorIndex(ColorMethod.None, 1);
                    }

                    br.Layer = layerBlockName;

                    // Коллекция точек пересечения блока и полилинии
                    Point3dCollection intersectionPoints = new Point3dCollection();

                    // Проверям не заползает ли блок на полилинию, если да, то смещаем на длину блока
                    foreach (Polyline pl in selectedPolylines)
                    {
                        Line disLine = new Line(new Point3d(0, 0, 0), new Point3d(2.4 * blockScale, 0, 0));


                        Vector3d acVec1 = disLine.StartPoint.GetVectorTo(l.EndPoint);
                        disLine.TransformBy(Matrix3d.Displacement(acVec1));
                        disLine.TransformBy(Matrix3d.Rotation((MyCorrectAngle(l.Angle)), curUCS.Zaxis, disLine.StartPoint));

                        disLine.IntersectWith(pl, Intersect.OnBothOperands, new Plane(), intersectionPoints, IntPtr.Zero, IntPtr.Zero);
                        
                        if (intersectionPoints.Count > 1)
                        {
                            Vector3d acVec2 = disLine.EndPoint.GetVectorTo(disLine.StartPoint);
                            disLine.TransformBy(Matrix3d.Displacement(acVec2));

                            Vector3d acVec3 = br.Position.GetVectorTo(disLine.StartPoint);
                            br.TransformBy(Matrix3d.Displacement(acVec3));
                        }

                    }

                    l.Erase(); // удаляем все кратчайшие отрезки 
                } // foreach closestlines

                tr.Commit();
            }
        } // method

        // Метод создания отклонений от стен в верхнем и нижнем сечении
        protected void MyDeviationST(double _blockScale = 0.25, double _tolerancePlaneRED = 0.015, double _toleranceAngleRED = 0.015, string _layerBlockName = "Плановые отклонения", double _toleranceCreateClosestLines = 0.1, string _layerPoly = "Стены", string _layerPoints = "Точки", string _pathToBlock = @"C:\arrowBlocks\arrowDinoBlocks.dwg")
        {
            double tolerancePlaneRED = _tolerancePlaneRED;
            double toleranceAngleRED = _toleranceAngleRED;
            string layerBlockName = _layerBlockName;
            double toleranceCreateClosestLines = _toleranceCreateClosestLines;
            string layerPoly = _layerPoly;
            string layerPoints = _layerPoints;
            string pathToBlock = _pathToBlock;
            double blockScale = _blockScale;

            Document doc = Application.DocumentManager.MdiActiveDocument; // получаем ссылку на активный документ
            Editor ed = doc.Editor;
            Database db = doc.Database; // получаем базу данных докумеента

            using (DocumentLock docLoc = doc.LockDocument()) // если не сделать, то в некоторых случаях приводит к возникновению ошибки при изменении базы данных документа
            using (Transaction tr = db.TransactionManager.StartTransaction()) // запускаем транзакцию
            {

                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable; // таблица блоков
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord; // пространство модели
                MyNewLayer(layerBlockName); // создаем новый слой
                //MyNewLayer("Красный", 255,0,0);
                MyInsertBlock(pathToBlock); // всталяем наш DWG с блоками

                List<Line> allClosestLines = new List<Line>(MyMakeClosestLines(toleranceCreateClosestLines, layerPoly, layerPoints, layerBlockName)); // вставляем список кратчайших отрезков от точки до полилинии
                List<Polyline> selectedPolylines = new List<Polyline>(MyTakePolylines(layerPoly)); // вставляем список из полилиний из чертежа

                // Создаем два списка чтобы разбить список на 2
                List<Line> tempClosestLines = new List<Line>();
                List<Line> closestLinesDown = new List<Line>();
                List<Line> closestLinesUp = new List<Line>();

                List<Line> closestLinesDownToUp = new List<Line>();

                double min = 0;
                string disDown = "??";
                string disUp = "??";
                bool isRightDown = true;
                bool isRightUp = true;
                double tempDisUpTolerance = 0;
                double tempDisAngleTolerance = 0;

                Matrix3d curUCSMatrix = ed.CurrentUserCoordinateSystem; // создаем матрицу координат
                CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d; // содаем текущую систему координат

                foreach (Line l in allClosestLines)
                {
                    if (l.Color == Color.FromRgb(0, 0, 5)) // если цвет синий
                    {
                        closestLinesDown.Add(l);
                    }
                    else if (l.Color == Color.FromRgb(0, 0, 6)) // если цвет фиолетовый
                    {
                        closestLinesUp.Add(l);
                    }
                }

                for (int i = 0; i < closestLinesDown.Count; i++) // перебираем точки
                {
                    for (int j = 0; j < closestLinesUp.Count; j++) // перебераем полилинии для каждой точки
                    {
                        Line up = closestLinesUp[j]; // присваиваем jтовую полилинию
                        Line down = closestLinesDown[i]; // присваиваем iтовую точку

                        // создаем полилинию где первая координата - это координаты точки, вторая - координаты ближайшей точки на полилинии
                        if (down.StartPoint.DistanceTo(up.StartPoint) < 0.5)
                        {
                            Line line1 = new Line(down.StartPoint, up.StartPoint);
                            tempClosestLines.Add(line1);
                            ms.AppendEntity(line1);
                            tr.AddNewlyCreatedDBObject(line1, true);
                        }

                    } // for up


                    for (int k = 0; k < tempClosestLines.Count; k++)
                    {
                        min = tempClosestLines[0].Length;
                        if (tempClosestLines[k].Length < min)
                        {
                            min = tempClosestLines[k].Length;
                        }

                    }

                    foreach (Line l in tempClosestLines)
                    {
                        if (l.Length == min)
                        {
                            l.ColorIndex = 1;
                            closestLinesDownToUp.Add(l);
                        }
                        else
                        {
                            l.Erase();
                        }
                    }

                    tempClosestLines.Clear();


                } // for down

                ///////////////////////////////////////////////////////////////////////////////////////////
                ///



                foreach (Line ld in closestLinesDown)
                {
                    if ((ld.Angle * (180 / Math.PI)) > 90.0 && (ld.Angle * (180 / Math.PI)) <= 270.0)
                    {
                        isRightDown = true;
                    }
                    else
                    {
                        isRightDown = false;
                    }

                    foreach (Line cl in closestLinesDownToUp)
                    {
                        if (cl.StartPoint == ld.StartPoint)
                        {
                            foreach (Line lu in closestLinesUp)
                            {
                                if (lu.StartPoint == cl.EndPoint)
                                {
                                    if ((lu.Angle * (180 / Math.PI)) > 90.0 && (lu.Angle * (180 / Math.PI)) <= 270.0)
                                    {
                                        isRightUp = true;
                                    }
                                    else
                                    {
                                        isRightUp = false;
                                    }
                                }
                            }
                        }
                    }

                    if (isRightDown == isRightUp)
                    {
                        BlockTableRecord blockDefR = bt["arrowDinoBlockR"].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        BlockReference brR = new BlockReference(ld.EndPoint, bt["arrowDinoBlockR"]);

                        ms.AppendEntity(brR);
                        tr.AddNewlyCreatedDBObject(brR, true);

                        disUp = "??";

                        foreach (Line cl in closestLinesDownToUp)
                        {
                            if (cl.StartPoint == ld.StartPoint)
                            {
                                foreach (Line lu in closestLinesUp)
                                {
                                    if (lu.StartPoint == cl.EndPoint)
                                        disUp = string.Format("{0:0}", lu.Length * 1000);
                                }
                            }
                        }

                        disDown = string.Format("{0:0}", ld.Length * 1000);

                        foreach (ObjectId id in blockDefR) // Вставляем в блок атрибуты с отклонениями
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDefR = obj as AttributeDefinition;
                            if ((attDefR != null) && (!attDefR.Constant))
                            {
                                // Это неконстантный AttributeDefinition
                                // Создаём новый AttributeReference
                                using (AttributeReference attRef = new AttributeReference())
                                {
                                    attRef.SetAttributeFromBlock(attDefR, brR.BlockTransform);
                                    if (attRef.Tag == "ОТКЛВЕРХ")
                                        attRef.TextString = disUp;
                                    else if (attRef.Tag == "ОТКЛНИЗ")
                                        attRef.TextString = disDown;
                                    //Добавляем AttributeReference к BlockReference
                                    brR.AttributeCollection.AppendAttribute(attRef);
                                    tr.AddNewlyCreatedDBObject(attRef, true);

                                } // using attRef
                            } // if 
                        } // foreach attribute

                        // Меняем направление стрелки в динамическом блоке
                        if ((ld.Angle * (180 / Math.PI)) > 90.0 && (ld.Angle * (180 / Math.PI)) <= 270.0)
                        {

                        }
                        else
                        {
                            foreach (DynamicBlockReferenceProperty prop in brR.DynamicBlockReferencePropertyCollection)
                            {
                                object[] values = prop.GetAllowedValues();

                                if (prop.PropertyName == "mirror" && !prop.ReadOnly)
                                {
                                    if (prop.Value.ToString() == values[0].ToString())
                                        prop.Value = values[1];

                                    else
                                        prop.Value = values[0];
                                }
                            }
                        }  // Меняем направление стрелки в динамическом блоке

                        brR.TransformBy(Matrix3d.Rotation((MyCorrectAngle(ld.Angle)), curUCS.Zaxis, ld.EndPoint)); // поворачиваем блок на вычисленный угол, переведя обратно в радианы
                        brR.TransformBy(Matrix3d.Scaling(blockScale, ld.EndPoint)); // изменяем масштаб блока

                        brR.Layer = layerBlockName;

                        // Коллекция точек пересечения блока и полилинии
                        Point3dCollection intersectionPoints = new Point3dCollection();

                        // Проверям не заползает ли блок на полилинию, если да, то смещаем на длину блока
                        foreach (Polyline pl in selectedPolylines)
                        {
                            Line disLine = new Line(new Point3d(0, 0, 0), new Point3d(2.4 * blockScale, 0, 0));

                            Vector3d acVec1 = disLine.StartPoint.GetVectorTo(ld.EndPoint);
                            disLine.TransformBy(Matrix3d.Displacement(acVec1));
                            disLine.TransformBy(Matrix3d.Rotation((MyCorrectAngle(ld.Angle)), curUCS.Zaxis, disLine.StartPoint));

                            disLine.IntersectWith(pl, Intersect.OnBothOperands, new Plane(), intersectionPoints, IntPtr.Zero, IntPtr.Zero);

                            if (intersectionPoints.Count > 1)
                            {
                                Vector3d acVec2 = disLine.EndPoint.GetVectorTo(disLine.StartPoint);
                                disLine.TransformBy(Matrix3d.Displacement(acVec2));

                                Vector3d acVec3 = brR.Position.GetVectorTo(disLine.StartPoint);
                                brR.TransformBy(Matrix3d.Displacement(acVec3));
                            }
                        } // intersect

                        // Выделяем косяки
                        foreach (Line cl in closestLinesDownToUp)
                        {
                            if (cl.StartPoint == ld.StartPoint)
                            {
                                foreach (Line lu in closestLinesUp)
                                {
                                    if (lu.StartPoint == cl.EndPoint)
                                        tempDisUpTolerance = lu.Length;
                                }
                            }
                        }

                        if (ld.Length > tolerancePlaneRED || tempDisUpTolerance > tolerancePlaneRED)
                        {
                            MyNewLayer("!Вне допуска");
                            Circle c = new Circle();
                            c.Center = ld.EndPoint;
                            c.Radius = 2.4*blockScale;
                            c.Layer = "!Вне допуска";
                            c.Color = Color.FromColorIndex(ColorMethod.None, 1);
                            c.LineWeight = (LineWeight)30;
                            ms.AppendEntity(c);
                            tr.AddNewlyCreatedDBObject(c, true);
                            brR.Color = Color.FromColorIndex(ColorMethod.None, 1);
                        }
                    }
                    else
                    {
                        BlockTableRecord blockDefRL = bt["arrowDinoBlockRL"].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        BlockReference brRL = new BlockReference(ld.EndPoint, bt["arrowDinoBlockRL"]);

                        ms.AppendEntity(brRL);
                        tr.AddNewlyCreatedDBObject(brRL, true);

                        disUp = "??";

                        foreach (Line cl in closestLinesDownToUp)
                        {
                            if (cl.StartPoint == ld.StartPoint)
                            {
                                foreach (Line lu in closestLinesUp)
                                {
                                    if (lu.StartPoint == cl.EndPoint)
                                        disUp = string.Format("{0:0}", lu.Length * 1000);
                                }
                            }
                        }

                        disDown = string.Format("{0:0}", ld.Length * 1000);

                        foreach (ObjectId id in blockDefRL) // Вставляем в блок атрибуты с отклонениями
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDefRL = obj as AttributeDefinition;
                            if ((attDefRL != null) && (!attDefRL.Constant))
                            {
                                // Это неконстантный AttributeDefinition
                                // Создаём новый AttributeReference
                                using (AttributeReference attRef = new AttributeReference())
                                {
                                    attRef.SetAttributeFromBlock(attDefRL, brRL.BlockTransform);
                                    if (attRef.Tag == "ОТКЛВЕРХ")
                                        attRef.TextString = disUp;
                                    else if (attRef.Tag == "ОТКЛНИЗ")
                                        attRef.TextString = disDown;
                                    //Добавляем AttributeReference к BlockReference
                                    brRL.AttributeCollection.AppendAttribute(attRef);
                                    tr.AddNewlyCreatedDBObject(attRef, true);

                                } // using attRef
                            } // if 
                        } // foreach attribute

                        // Меняем направление стрелки в динамическом блоке
                        if (isRightDown == false)
                        {

                        }
                        else
                        {
                            foreach (DynamicBlockReferenceProperty prop in brRL.DynamicBlockReferencePropertyCollection)
                            {
                                object[] values = prop.GetAllowedValues();

                                if (prop.PropertyName == "mirror" && !prop.ReadOnly)
                                {
                                    if (prop.Value.ToString() == values[0].ToString())
                                        prop.Value = values[1];

                                    else
                                        prop.Value = values[0];
                                }
                            }
                        }  // Меняем направление стрелки в динамическом блоке

                        brRL.TransformBy(Matrix3d.Rotation((MyCorrectAngle(ld.Angle)), curUCS.Zaxis, ld.EndPoint)); // поворачиваем блок на вычисленный угол, переведя обратно в радианы
                        brRL.TransformBy(Matrix3d.Scaling(blockScale, ld.EndPoint)); // изменяем масштаб блока

                        brRL.Layer = layerBlockName;

                        // Коллекция точек пересечения блока и полилинии
                        Point3dCollection intersectionPoints = new Point3dCollection();

                        // Проверям не заползает ли блок на полилинию, если да, то смещаем на длину блока
                        foreach (Polyline pl in selectedPolylines)
                        {
                            Line disLine = new Line(new Point3d(0, 0, 0), new Point3d(2.4 * blockScale, 0, 0));

                            Vector3d acVec1 = disLine.StartPoint.GetVectorTo(ld.EndPoint);
                            disLine.TransformBy(Matrix3d.Displacement(acVec1));
                            disLine.TransformBy(Matrix3d.Rotation((MyCorrectAngle(ld.Angle)), curUCS.Zaxis, disLine.StartPoint));

                            disLine.IntersectWith(pl, Intersect.OnBothOperands, new Plane(), intersectionPoints, IntPtr.Zero, IntPtr.Zero);

                            if (intersectionPoints.Count > 1)
                            {
                                Vector3d acVec2 = disLine.EndPoint.GetVectorTo(disLine.StartPoint);
                                disLine.TransformBy(Matrix3d.Displacement(acVec2));

                                Vector3d acVec3 = brRL.Position.GetVectorTo(disLine.StartPoint);
                                brRL.TransformBy(Matrix3d.Displacement(acVec3));
                            }
                        } // intersect

                        // Выделяем косяки
                        foreach (Line cl in closestLinesDownToUp)
                        {
                            if (cl.StartPoint == ld.StartPoint)
                            {
                                foreach (Line lu in closestLinesUp)
                                {
                                    if (lu.StartPoint == cl.EndPoint)
                                        tempDisUpTolerance = lu.Length;
                                }
                            }
                        }

                        tempDisAngleTolerance = (tempDisUpTolerance + ld.Length);

                        if (ld.Length > tolerancePlaneRED || tempDisUpTolerance > tolerancePlaneRED || tempDisAngleTolerance > toleranceAngleRED)
                        {
                            MyNewLayer("!Вне допуска");
                            Circle c = new Circle();
                            c.Center = ld.EndPoint;
                            c.Radius = 2.4*blockScale;
                            c.Layer = "!Вне допуска";
                            c.Color = Color.FromColorIndex(ColorMethod.None, 1);
                            c.LineWeight = (LineWeight)30;
                            ms.AppendEntity(c);
                            tr.AddNewlyCreatedDBObject(c, true);
                            brRL.Color = Color.FromColorIndex(ColorMethod.None, 1);
                        }
                    } // else

                    foreach (Line l in allClosestLines)
                        l.Erase();

                    foreach (Line l in closestLinesDownToUp)
                        l.Erase();


                } // foreach cl
                
                tr.Commit();
            }
        } // method

        public void Initialize()
        {
           
        }

        public void Terminate()
        {
           
        }
    }
}

