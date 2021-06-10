//#define ACAD

using System;
using System.Collections.Generic;

#if ACAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
#else
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;
#endif

namespace mavCAD
{
    public class FastCAD
    {
        const double PI = Math.PI;

        // Свойства для связи с интерфейсом
        public string       selectedLayerPlines { get; set; }
        public string       selectedLayerPoints { get; set; }
        public List<string> listLayers          { get; set; }
        public double       toleranceRed        { get; set; }
        public double       toleranceRedAngle   { get; set; }
        public double       toleranceCreate     { get; set; }
        public double       blockScale          { get; set; }

        /// <summary>
        /// Cоздание слоя 
        /// </summary>
        public void MyNewLayer(string name)
        {
            Document adoc = Application.DocumentManager.MdiActiveDocument;// получаем ссылку на активный документ

            if (adoc == null) // если открыта начальная страница (док не активен=0), то выходим из метода
                return;

            Database db = adoc.Database; // получаем базу данных докумеента
            ObjectId layerTableId = db.LayerTableId; // получаем Id таблицы слоев

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
                if (!acLayerTable.Has(name))
                {
                    acLayerTableRecord.Name = name;

                    // заносим созданный слой таблицу слоев
                    acLayerTable.Add(acLayerTableRecord);

                    // добавляем созданный слой в документ
                    tr.AddNewlyCreatedDBObject(acLayerTableRecord, true);
                }

                // проверка если нет нашего слоя - прекращаем выполнение проги
                if (!acLayerTable.Has(name))
                {
                    return;
                }

                // db.Clayer = acLayerTable[name]; // устанавливаем новый слой как текущий

                tr.Commit(); // подтвержение выполнения транзакции, если не поставить то код не выполнится
            }
        }

        /// <summary>
        /// Cоздание слоя
        /// </summary>
        public void MyNewLayer(string name, byte r, byte g, byte b)
        {
            // получаем ссылку на активный документ
            Document adoc = Application.DocumentManager.MdiActiveDocument;

            if (adoc == null) // если открыта начальная страница (док не активен=0), то выходим из метода
                return;

            Database db = adoc.Database; // получаем базу данных докумеента
            ObjectId layerTableId = db.LayerTableId; // получаем Id таблицы слоев

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
                if (!acLayerTable.Has(name))
                {
                    acLayerTableRecord.Name = name;
                    acLayerTableRecord.Color = Color.FromRgb(r, g, b);
                    // заносим созданный слой таблицу слоев
                    acLayerTable.Add(acLayerTableRecord);

                    // добавляем созданный слой в документ
                    tr.AddNewlyCreatedDBObject(acLayerTableRecord, true);
                }

                // проверка если нет нашего слоя - прекращаем выполнение проги
                if (!acLayerTable.Has(name))
                {
                    return;
                }

                // db.Clayer = acLayerTable[name]; // устанавливаем новый слой как текущий

                tr.Commit(); // подтвержение выполнения транзакции, если не поставить то код не выполнится
            }

        }

        /// <summary>
        /// Вставка блока из внешнего файла
        /// </summary>
        public void MyInsertBlock(string path = @"C:\arrowBlocks\arrowDinoBlocks.dwg")
        {           
            Document doc = Application.DocumentManager.MdiActiveDocument;
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

                            if (dbObj.Name == "arrowDinoBlocks")
                                dbObj.Erase();
                        }
                    }
                    tr.Commit();
                }
            }
        }

        /// <summary>
        /// Cоздание списка отрезков от точки до ближайшей полилинии (цвета в формате RGB 005 - для низа (синие точки), 006 - для верха (фиолетовые точки), 007 - остальные)
        /// </summary>
        public List<Line> MyMakeClosestLines(double tolerance, string layerPoints, string layerLines)
        {
            //  tolerance     максимальное расстояние при котором будут рисоваться отрезки
            //  layerPoints   слой выборки точек в чертеже
            //  layerLines    слой выборки полилиний в чертеже

            List<Line> ClosestLineList = new List<Line>();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // начинаем транзакцию
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {

                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // создаем списки с полилиниями и точками
                List<Polyline> listPlines = MyTakePolylines(layerLines);
                List<DBPoint> listPoints = MyTakePoints(layerPoints);

                foreach(DBPoint point in listPoints)
                {
                    foreach (Polyline acPoly in listPlines)
                    {
                        // определяем координаты ближайшей точки к полилинии
                        Point3d closestPoint = acPoly.GetClosestPointTo(point.Position, false);

                        if (point.ColorIndex == 5) // Если цвет точки синий (низ)
                        {
                            // создаем полилинию где первая координата - это координаты точки, вторая - координаты ближайшей точки на полилинии
                            Line myLine = new Line(new Point3d(point.Position.X, point.Position.Y, 0), new Point3d(closestPoint.X, closestPoint.Y, 0));
                            myLine.Color = Color.FromRgb(0, 0, 5); // задаем уникальный цвет 
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
                            myLine.Color = Color.FromRgb(0, 0, 6); // задаем уникальный цвет 
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
                            myLine.Color = Color.FromRgb(0, 0, 7); // задаем уникальный цвет 
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
        }

        /// <summary>
        /// Получить все полилинии в четреже (и в скрытых слоях и в заблокированных)
        /// </summary>
        public List<Polyline> MyTakePolylines(string layerPoly)
        {
            List<Polyline> listPlines = new List<Polyline>();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // создаем переменную, в которой будут содержаться данные для фильтра
            TypedValue[] filterPoly = new TypedValue[2];

            // первый аргумент (0) указывает, что мы задаем тип объекта 
            // второй аргумент ("POINT") - собственно тип  // LWPOLYLINE - упрощенная полилиния, исп-ся везде; POLYLINE - 2Д полилиния которая грузит систему, была в старых версиях
            filterPoly[0] = new TypedValue((int)DxfCode.Start, "LWPOLYLINE"); // условие выбора чтоб это была полилиния
            filterPoly[1] = new TypedValue((int)DxfCode.LayerName, layerPoly); // И чтобы находилась в слое "Стены"

            // создаем фильтр 1й выборки
            SelectionFilter selectionFilterPoly = new SelectionFilter(filterPoly);

            // пытаемся получить ссылки на объекты с учетом фильтра
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
                foreach (ObjectId idPoly in idsPoly) // заполняем список полилиниями из набора, приобразовав их в тип Polyline
                {
                    Polyline acPoly = (Polyline)tr.GetObject(idPoly, OpenMode.ForRead);
                    listPlines.Add(acPoly);
                }
                tr.Commit(); // подтверждаем транзакцию
            }
            return listPlines;
        } // method 

        /// <summary>
        /// Получить все точки в четреже (и в скрытых слоях и в заблокированных)
        /// </summary>
        public List<DBPoint> MyTakePoints(string layerPoints)
        {
            List<DBPoint> listPoints = new List<DBPoint>();

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

        /// <summary>
        /// Получить все имена слоёв в чертеже
        /// </summary>
        public List<string> MyTakeLayers()
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

        /// <summary>
        /// Корректируем угол на ГОСТ'овский. Чтобы текст читался слева направо, снизу вверх
        /// </summary>
        public double MyCorrectAngle(double radianAngle)
        {
            double correctAngle;

            correctAngle = (radianAngle) * (180 / PI);  // угол получаем в радианах, для удобного вычисления переводим в градусы

            if (correctAngle > 135 && correctAngle <= 315)
            {
                correctAngle = correctAngle + 180;

                if (correctAngle >= 360)
                    correctAngle = correctAngle - 360;

                if (correctAngle < 0)
                    correctAngle = correctAngle * -1;
            }
            return correctAngle * (PI / 180);
        }

        /// <summary>
        /// Проверка в каком направлении расположен объект
        /// </summary>
        public bool isRightDirection(double radAngle)
        {
            double angle = radAngle * (180 / PI);
            return (angle > 90.0) && (angle <= 270.0);
        }

        /// <summary>
        /// Отклонение от проекта для плит перекрытий
        /// </summary>
        public void MyDeviationPP(string layerBlockName = "Плановые отклонения",
                                  string pathToBlock = @"C:\arrowDinoBlocks\arrowDinoBlocks.dwg")
        {
            Document doc = Application.DocumentManager.MdiActiveDocument; // получаем ссылку на активный документ
            Editor ed = doc.Editor;
            Database db = doc.Database; // получаем базу данных докумеента

            using (Transaction tr = db.TransactionManager.StartTransaction()) // запускаем транзакцию
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable; // таблица блоков
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord; // пространство модели
                MyNewLayer(layerBlockName); // создаем новый слой
                MyInsertBlock(pathToBlock); // всталяем наши блоки из DWG

                List<Line> closestLines = MyMakeClosestLines(toleranceCreate, selectedLayerPoints, selectedLayerPlines); // получаем список кратчайших отрезков от точки до полилинии
                List<Polyline> selectedPolylines = MyTakePolylines(selectedLayerPlines); // получаем список полилиний из чертежа

                foreach (Line clLine in closestLines) // вставляем блоки на каждой линии
                {
                    Matrix3d curUCSMatrix = ed.CurrentUserCoordinateSystem; // получаем текущую матрицу координат
                    CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d; // получаем текущую пользовательскую систему координат

                    BlockTableRecord blockDef = bt["arrowDinoBlockRPL"].GetObject(OpenMode.ForRead) as BlockTableRecord; // получаем определение блока в таблице с именем "arrowDinoBlockRPL"
                    BlockReference blockRef = new BlockReference(clLine.EndPoint, bt["arrowDinoBlockRPL"]); // вставляем блок в пространство модели в точку конца отрезка

                    ms.AppendEntity(blockRef);
                    tr.AddNewlyCreatedDBObject(blockRef, true);

                    // Вставляем в блок атрибуты с отклонениями
                    foreach (ObjectId id in blockDef)
                    {
                        DBObject obj = id.GetObject(OpenMode.ForRead);
                        AttributeDefinition attDef = obj as AttributeDefinition;
                        if ((attDef != null) && (!attDef.Constant))
                        {
                            // Создаём новый AttributeReference
                            using (AttributeReference attRef = new AttributeReference())
                            {
                                attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                                if (attRef.Tag == "ОТКЛ")
                                    attRef.TextString = string.Format("{0:0}", clLine.Length * 1000); // устанавливаем значение атрибута ОТКЛОНЕНИЕ в мм
                                // Добавляем AttributeReference к BlockReference
                                blockRef.AttributeCollection.AppendAttribute(attRef);
                                tr.AddNewlyCreatedDBObject(attRef, true);
                            }
                        }
                    }

                    // Меняем направление стрелки в динамическом блоке
                    if (!isRightDirection(clLine.Angle))
                    {
                        // Перебираем коллекцию всех параметров динамического блока
                        foreach (DynamicBlockReferenceProperty prop in blockRef.DynamicBlockReferencePropertyCollection)
                        {
                            object[] values = prop.GetAllowedValues();

                            if (prop.PropertyName == "mirror" && !prop.ReadOnly) // меняем местами значение "Отражение" для стрелки
                            {
                                if (prop.Value.ToString() == values[0].ToString())
                                    prop.Value = values[1];
                                else
                                    prop.Value = values[0];
                            }
                        }
                    }

                    blockRef.TransformBy(Matrix3d.Rotation((MyCorrectAngle(clLine.Angle)), curUCS.Zaxis, clLine.EndPoint)); // поворачиваем блок
                    blockRef.TransformBy(Matrix3d.Scaling(blockScale, clLine.EndPoint)); // изменяем масштаб блока

                    blockRef.Layer = layerBlockName;

                    // Коллекция точек пересечения блока и полилинии
                    Point3dCollection intersectionPoints = new Point3dCollection();

                    // Проверям не заползает ли блок на полилинию, если да, то смещаем на длину блока
                    foreach (Polyline pl in selectedPolylines)
                    {
                        Line tempLine = new Line(new Point3d(0, 0, 0), new Point3d(2.4 * blockScale, 0, 0)); // получаем отрезок длинной в стрелку

                        // перемещаем и поворачиваем отрезок на блок
                        Vector3d acVec1 = tempLine.StartPoint.GetVectorTo(clLine.EndPoint);
                        tempLine.TransformBy(Matrix3d.Displacement(acVec1));
                        tempLine.TransformBy(Matrix3d.Rotation((MyCorrectAngle(clLine.Angle)), curUCS.Zaxis, tempLine.StartPoint));

                        // проверяем пересечается ли отрезок с полилинией и выводит точки пересечения в intersectionPoints
                        tempLine.IntersectWith(pl, Intersect.OnBothOperands, new Plane(), intersectionPoints, IntPtr.Zero, IntPtr.Zero);

                        // Если блок пересечает стену, смещаем его на его же длину
                        if (intersectionPoints.Count > 1)
                        {
                            Vector3d acVec2 = tempLine.EndPoint.GetVectorTo(tempLine.StartPoint);
                            tempLine.TransformBy(Matrix3d.Displacement(acVec2));

                            Vector3d acVec3 = blockRef.Position.GetVectorTo(tempLine.StartPoint);
                            blockRef.TransformBy(Matrix3d.Displacement(acVec3));
                        }
                    }

                    clLine.Erase(); // удаляем все кратчайшие отрезки 

                   //----------------------------------ВЫДЕЛЯЕМ КОСЯКИ ---------------------------------------------------------------------------------

                    if (clLine.Length > toleranceRed) // Выделяем косяки
                    {
                        MyNewLayer("!Вне допуска");
                        Circle c = new Circle();
                        c.Center = clLine.EndPoint;
                        c.Radius = 2.4 * blockScale; // 2.4 - длина стрелки
                        c.Layer = "!Вне допуска";
                        c.Color = Color.FromColorIndex(ColorMethod.None, 1);
                        c.LineWeight = (LineWeight)30;
                        ms.AppendEntity(c);
                        tr.AddNewlyCreatedDBObject(c, true);
                        blockRef.Color = Color.FromColorIndex(ColorMethod.None, 1);
                    }
                } // foreach closestlines
                tr.Commit();
            }
        }

        /// <summary>
        /// Отклонение от проекта для стен в верхнем и нижнем сечении
        /// </summary>
        public void MyDeviationST(string layerBlockName = "Плановые отклонения [В] [Н]",
                                  string pathToBlock = @"C:\arrowDinoBlocks\arrowDinoBlocks.dwg")
        {
            Document doc = Application.DocumentManager.MdiActiveDocument; // получаем ссылку на активный документ
            Editor ed = doc.Editor;
            Database db = doc.Database; // получаем базу данных документа

            using (Transaction tr = db.TransactionManager.StartTransaction()) // запускаем транзакцию
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable; // таблица блоков
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord; // пространство модели
                MyNewLayer(layerBlockName); // создаем новый слой
                MyInsertBlock(pathToBlock); // всталяем наши блоки из DWG

                List<Line> allClosestLines = MyMakeClosestLines(toleranceCreate, selectedLayerPoints, selectedLayerPlines); // получаем список кратчайших отрезков от точки до полилинии
                List<Polyline> selectedPolylines = MyTakePolylines(selectedLayerPlines); // получаем список полилиний из чертежа

                // Создаем два списка чтобы разбить список на верх и низ               
                List<Line> closestLinesDown = new List<Line>();
                List<Line> closestLinesUp = new List<Line>();

                List<Line> closestLinesDownToUp = new List<Line>();

                string disUp = "??";                // отклонение верхней стрелки
                string disDown = "??";              // отклонение нижней стрелки
                bool isRightDown = true;            // нижняя стрелка направлена вправо
                bool isRightUp = true;              // верхняя стрелка направлена вправо
                double tempDisUpTolerance = 0;      // временное значение отклонения верхней стрелки
                double tempDisAngleTolerance = 0;   // фактическое отклонение верхней + нижней стрелки (уклон) 

                Matrix3d curUCSMatrix = ed.CurrentUserCoordinateSystem;         // создаем матрицу координат
                CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d;    // получаем текущую пользовательскую систему координат

                // Заполняем списки отрезками по верху и низу
                foreach (Line l in allClosestLines)
                {
                    if (l.Color == Color.FromRgb(0, 0, 5)) // если цвет синий                    
                        closestLinesDown.Add(l);

                    else if (l.Color == Color.FromRgb(0, 0, 6)) // если цвет фиолетовый                    
                        closestLinesUp.Add(l);
                }

                foreach (Line down in closestLinesDown) // перебираем нижние, заполняем closestLinesDownToUp кратчайшие отрезки от начала нижней до начала верхней
                {
                    List<Line> tempClosestLines = new List<Line>();

                    foreach (Line up in closestLinesUp)
                    {
                        // Проверяем насколько далеко (в плане, без учета высоты) верхняя расположена относительно нижней
                        // Если меньше 0.5, то создаем отрезок и помещаем в tempClosestLines
                        if (new Point3d(down.StartPoint.X, down.StartPoint.Y, 0).DistanceTo(new Point3d(up.StartPoint.X, up.StartPoint.Y, 0)) < 0.5)
                        {
                            Line line = new Line(down.StartPoint, up.StartPoint);
                            tempClosestLines.Add(line);
                            ms.AppendEntity(line);
                            tr.AddNewlyCreatedDBObject(line, true);
                        }
                    } // верх

                    double minDownToUpLength = 0;       // для вычисления кратчайшей линии от начала нижней до начала верхней

                    for (int i = 0; i < tempClosestLines.Count; i++) // ищем самую короткую
                    {
                        minDownToUpLength = tempClosestLines[0].Length;

                        if (tempClosestLines[i].Length < minDownToUpLength)
                            minDownToUpLength = tempClosestLines[i].Length;
                    }

                    // добавляем в closestLinesDownToUp только самый короткий отрезок от линии, остальные удаляем
                    foreach (Line line in tempClosestLines)
                    {
                        if (line.Length == minDownToUpLength)
                        {
                            line.ColorIndex = 1;
                            closestLinesDownToUp.Add(line);
                        }
                        else
                            line.Erase();
                    }
                    tempClosestLines.Clear();
                } // заполняем closestLinesDownToUp

                //---------------------------------- ОСНОВНОЙ КОД ---------------------------------------------------------------------------------

                foreach (Line lineDowntoUp in closestLinesDownToUp) // Проходимся по всем отрезкам у которых найдены соседи
                {                  
                    foreach (Line lDown in closestLinesDown)
                    {
                        if (lDown.StartPoint != lineDowntoUp.StartPoint)
                            continue;

                        isRightDown = isRightDirection(lDown.Angle);

                        foreach (Line lUp in closestLinesUp)
                        {
                            if (lUp.StartPoint != lineDowntoUp.EndPoint)
                                continue;

                            isRightUp = isRightDirection(lUp.Angle);

                            BlockTableRecord blockDef;
                            BlockReference blockRef;
                            if (isRightDown == isRightUp)
                            {
                                blockDef = bt["arrowDinoBlockR"].GetObject(OpenMode.ForRead) as BlockTableRecord;
                                blockRef = new BlockReference(lDown.EndPoint, bt["arrowDinoBlockR"]);

                                ms.AppendEntity(blockRef);
                                tr.AddNewlyCreatedDBObject(blockRef, true);                             
                            }
                            else
                            {
                                blockDef = bt["arrowDinoBlockRL"].GetObject(OpenMode.ForRead) as BlockTableRecord;
                                blockRef = new BlockReference(lDown.EndPoint, bt["arrowDinoBlockRL"]);

                                ms.AppendEntity(blockRef);
                                tr.AddNewlyCreatedDBObject(blockRef, true);
                            }

                            disUp   = string.Format("{0:0}", lUp.Length * 1000);
                            disDown = string.Format("{0:0}", lDown.Length * 1000);

                            foreach (ObjectId id in blockDef) // Вставляем в блок атрибуты с отклонениями
                            {
                                DBObject obj = id.GetObject(OpenMode.ForRead);
                                AttributeDefinition attrDef = obj as AttributeDefinition;
                                if ((attrDef != null) && (!attrDef.Constant))
                                {
                                    using (AttributeReference attRef = new AttributeReference())
                                    {
                                        attRef.SetAttributeFromBlock(attrDef, blockRef.BlockTransform);
                                        if (attRef.Tag == "ОТКЛВЕРХ")
                                            attRef.TextString = disUp;
                                        else if (attRef.Tag == "ОТКЛНИЗ")
                                            attRef.TextString = disDown;
                                        //Добавляем AttributeReference к BlockReference
                                        blockRef.AttributeCollection.AppendAttribute(attRef);
                                        tr.AddNewlyCreatedDBObject(attRef, true);
                                    } // using attRef
                                } // if 
                            } // foreach attribute

                            if (!isRightDirection(lDown.Angle)) // Меняем направление стрелки в динамическом блоке
                            {
                                foreach (DynamicBlockReferenceProperty prop in blockRef.DynamicBlockReferencePropertyCollection)
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
                            }

                            blockRef.TransformBy(Matrix3d.Rotation((MyCorrectAngle(lDown.Angle)), curUCS.Zaxis, lDown.EndPoint)); // поворачиваем блок на вычисленный угол
                            blockRef.TransformBy(Matrix3d.Scaling(blockScale, lDown.EndPoint)); // изменяем масштаб блока
                            blockRef.Layer = layerBlockName;

                            // Коллекция точек пересечения блока и полилинии
                            Point3dCollection intersectionPoints = new Point3dCollection();

                            // Проверям не заползает ли блок на полилинию, если да, то смещаем на длину блока
                            foreach (Polyline plWall in selectedPolylines)
                            {
                                Line tempLine = new Line(new Point3d(0, 0, 0), new Point3d(2.4 * blockScale, 0, 0)); // 2.4 - длина стрелки в блоке

                                Vector3d acVec1 = tempLine.StartPoint.GetVectorTo(lDown.EndPoint);                                                         
                                tempLine.TransformBy(Matrix3d.Displacement(acVec1));                                                                       
                                tempLine.TransformBy(Matrix3d.Rotation((MyCorrectAngle(lDown.Angle)), curUCS.Zaxis, tempLine.StartPoint));                 
                                                                                                                                                           
                                tempLine.IntersectWith(plWall, Intersect.OnBothOperands, new Plane(), intersectionPoints, IntPtr.Zero, IntPtr.Zero);       
                                                                                                                                                           
                                if (intersectionPoints.Count > 1)                                                                                          
                                {                                                                                                                          
                                    Vector3d acVec2 = tempLine.EndPoint.GetVectorTo(tempLine.StartPoint);                                                  
                                    tempLine.TransformBy(Matrix3d.Displacement(acVec2));                                                                   
                                                                                                                                                           
                                    Vector3d acVec3 = blockRef.Position.GetVectorTo(tempLine.StartPoint);                                                  
                                    blockRef.TransformBy(Matrix3d.Displacement(acVec3));                                                                   
                                }                                                                                                                            
                            }

                            //----------------------------------ВЫДЕЛЯЕМ КОСЯКИ ---------------------------------------------------------------------------------

                            if (lineDowntoUp.StartPoint == lDown.StartPoint)
                            {
                                if (lUp.StartPoint == lineDowntoUp.EndPoint)
                                    tempDisUpTolerance = lUp.Length;
                            }


                            bool isBadDeviation = false;

                            if (isRightDown == isRightUp)
                            {
                                if (lDown.Length > toleranceRed || tempDisUpTolerance > toleranceRed)
                                    isBadDeviation = true;
                            }
                            else
                            {
                                tempDisAngleTolerance = (tempDisUpTolerance + lDown.Length);

                                if (lDown.Length > toleranceRed || tempDisUpTolerance > toleranceRed || tempDisAngleTolerance > toleranceRedAngle)
                                    isBadDeviation = true;
                            }

                            if (isBadDeviation)
                            {
                                MyNewLayer("!Вне допуска");
                                Circle c = new Circle();
                                c.Center = lDown.EndPoint;
                                c.Radius = 2.4 * blockScale;
                                c.Layer = "!Вне допуска";
                                c.Color = Color.FromColorIndex(ColorMethod.None, 1);
                                c.LineWeight = (LineWeight)30;
                                ms.AppendEntity(c);
                                tr.AddNewlyCreatedDBObject(c, true);
                                blockRef.Color = Color.FromColorIndex(ColorMethod.None, 1);
                            }
                        }
                    }
                } // Проходимся по всем отрезкам у которых найдены соседи

                foreach (Line clLine in allClosestLines) // Проходимся по одиноким точкам, без соседей
                {
                    if (clLine.Color == Color.FromRgb(0, 0, 7)) // пропускаем точки без цвета верх\низ               
                        continue;

                    bool isSingle = true;

                    foreach (Line clLineDownToUp in closestLinesDownToUp) 
                    {
                        if ((clLine.StartPoint == clLineDownToUp.StartPoint) || (clLine.StartPoint == clLineDownToUp.EndPoint))
                            isSingle = false;
                    }

                    if(!isSingle)
                        continue;

                    disUp = "??";
                    disDown = "??";

                    BlockTableRecord blockDef;
                    BlockReference blockRef;

                    blockDef = bt["arrowDinoBlockR"].GetObject(OpenMode.ForRead) as BlockTableRecord;
                    blockRef = new BlockReference(clLine.EndPoint, bt["arrowDinoBlockR"]);

                    ms.AppendEntity(blockRef);
                    tr.AddNewlyCreatedDBObject(blockRef, true);

                    if (clLine.Color == Color.FromRgb(0, 0, 5)) // синий
                    {
                        disDown = string.Format("{0:0}", clLine.Length * 1000);
                    }
                    else if (clLine.Color == Color.FromRgb(0, 0, 6)) // фиолетовый
                    {
                        disUp = string.Format("{0:0}", clLine.Length * 1000);
                    }

                    foreach (ObjectId id in blockDef) // Вставляем в блок атрибуты с отклонениями
                    {
                        DBObject obj = id.GetObject(OpenMode.ForRead);
                        AttributeDefinition attrDef = obj as AttributeDefinition;
                        if ((attrDef != null) && (!attrDef.Constant))
                        {
                            using (AttributeReference attRef = new AttributeReference())
                            {
                                attRef.SetAttributeFromBlock(attrDef, blockRef.BlockTransform);
                                if (attRef.Tag == "ОТКЛВЕРХ")
                                    attRef.TextString = disUp;
                                else if (attRef.Tag == "ОТКЛНИЗ")
                                    attRef.TextString = disDown;
                                //Добавляем AttributeReference к BlockReference
                                blockRef.AttributeCollection.AppendAttribute(attRef);
                                tr.AddNewlyCreatedDBObject(attRef, true);
                            } // using attRef
                        } // if 
                    } // foreach attribute

                    if (!isRightDirection(clLine.Angle)) // Меняем направление стрелки в динамическом блоке
                    {
                        foreach (DynamicBlockReferenceProperty prop in blockRef.DynamicBlockReferencePropertyCollection)
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
                    }

                    blockRef.TransformBy(Matrix3d.Rotation((MyCorrectAngle(clLine.Angle)), curUCS.Zaxis, clLine.EndPoint)); // поворачиваем блок на вычисленный угол
                    blockRef.TransformBy(Matrix3d.Scaling(blockScale, clLine.EndPoint)); // изменяем масштаб блока
                    blockRef.Layer = layerBlockName;

                    // Коллекция точек пересечения блока и полилинии
                    Point3dCollection intersectionPoints = new Point3dCollection();

                    // Проверям не заползает ли блок на полилинию, если да, то смещаем на длину блока
                    foreach (Polyline plWall in selectedPolylines)
                    {
                        Line tempLine = new Line(new Point3d(0, 0, 0), new Point3d(2.4 * blockScale, 0, 0)); // 2.4 - длина стрелки в блоке

                        Vector3d acVec1 = tempLine.StartPoint.GetVectorTo(clLine.EndPoint);                                                        
                        tempLine.TransformBy(Matrix3d.Displacement(acVec1));                                                                       
                        tempLine.TransformBy(Matrix3d.Rotation((MyCorrectAngle(clLine.Angle)), curUCS.Zaxis, tempLine.StartPoint));                
                                                                                                                                                   
                        tempLine.IntersectWith(plWall, Intersect.OnBothOperands, new Plane(), intersectionPoints, IntPtr.Zero, IntPtr.Zero);       
                                                                                                                                                   
                        if (intersectionPoints.Count > 1)                                                                                          
                        {                                                                                                                          
                            Vector3d acVec2 = tempLine.EndPoint.GetVectorTo(tempLine.StartPoint);                                                  
                            tempLine.TransformBy(Matrix3d.Displacement(acVec2));                                                                   
                                                                                                                                                   
                            Vector3d acVec3 = blockRef.Position.GetVectorTo(tempLine.StartPoint);                                                  
                            blockRef.TransformBy(Matrix3d.Displacement(acVec3));                                                                   
                        }
                    }

                    //---------------------------------- ВЫДЕЛЯЕМ КОСЯКИ ---------------------------------------------------------------------------------

                    if (clLine.Length > toleranceRed)
                    {
                        MyNewLayer("!Вне допуска");
                        Circle c = new Circle();
                        c.Center = clLine.EndPoint;
                        c.Radius = 2.4 * blockScale;
                        c.Layer = "!Вне допуска";
                        c.Color = Color.FromColorIndex(ColorMethod.None, 1);
                        c.LineWeight = (LineWeight)30;
                        ms.AppendEntity(c);
                        tr.AddNewlyCreatedDBObject(c, true);
                        blockRef.Color = Color.FromColorIndex(ColorMethod.None, 1);
                    }
                } // Проходимся по одиноким точкам, без соседей

                // Удаляем вспомогательные линии 
                foreach (Line l in allClosestLines)
                    l.Erase();

                foreach (Line l in closestLinesDownToUp)
                    l.Erase();
                // Подтверждаем транзакцию
                tr.Commit();
            } // Transaction
        } // MyDeviationST
    }

}

