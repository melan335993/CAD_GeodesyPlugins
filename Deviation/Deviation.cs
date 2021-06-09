//#define ACAD

#if ACAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
#else
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using Teigha.Runtime;
#endif

namespace mavCAD
{
    public class Deviation
    {
        [CommandMethod("PP_Deviation")]
        public void PP_Deviation()
        {
            FastCAD data = new FastCAD();

            data.listLayers = data.MyTakeLayers();

            PPDialog window = new PPDialog(data);

            if (Application.ShowModalWindow(window) != true)
                return;

            //window.textbox_deviationCreate.Text = "0.500";
            //window.textbox_deviationRed.Text = "0.015";

            double toleranceRED = window._data.deviationRed;                    // допустимое значение по отклонениям в плане
            double toleranceCreate = window._data.deviationCreate;              // максимальное расстояние при котором будут создаваться стрелки
            string layerBlockName = "Плановые отклонения";                      // имя нового слоя под стрелки
            string layerPoly = window.ComboPlines.SelectedValue.ToString();     // слой с полилиниями
            string layerPoints = window.ComboPoints.SelectedValue.ToString();   // слой с точками
            double blockScale = window._data.blockScale;                        // масштаб блока
            string pathToBlock = @"C:\arrowDinoBlocks\arrowDinoBlocks.dwg";     // путь до dwg с блоками стрелками

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                data.MyDeviationPP(blockScale, toleranceRED, layerBlockName, toleranceCreate, layerPoly, layerPoints, pathToBlock);
                ed.WriteMessage("\n***Отклонения успешно построены***\n");
                ed.Regen();

            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\n***Что-то пошло не так***\n{ex.Message}\n");
            }
        }

        // Отрисовка плановых отклонений от полилиний
        [CommandMethod("ST_Deviation")]
        public void ST_Deviation()
        {
            FastCAD data = new FastCAD();

            data.listLayers = data.MyTakeLayers();

            STDialog window = new STDialog(data);

            if (Application.ShowModalWindow(window) != true)
                return;

            double toleranceRED = window._data.deviationRed;                    // допустимое значение по отклонениям в плане
            double toleranceREDAngle = window._data.deviationRedAngle;          // допустимое значение по отклонениям от плоскости (уклон стены)
            double toleranceCreate = window._data.deviationCreate;              // максимальное расстояние при котором будут создаваться стрелки
            string layerBlockName = "Плановые отклонения [В] [Н]";              // имя нового слоя под стрелки
            string layerPoly = window.ComboPlines.SelectedValue.ToString();     // слой с полилиниями
            string layerPoints = window.ComboPoints.SelectedValue.ToString();   // слой с точками
            double blockScale = window._data.blockScale;                        // масштаб блока
            string pathToBlock = @"C:\arrowDinoBlocks\arrowDinoBlocks.dwg";     // путь до dwg с блоками стрелками

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                data.MyDeviationST(blockScale, toleranceRED, toleranceREDAngle, layerBlockName, toleranceCreate, layerPoly, layerPoints, pathToBlock);
                ed.WriteMessage("\n***Отклонения успешно построены***\n");
                ed.Regen();
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\n***Что-то пошло не так***\n{ex.Message}\n");
            }
        }
    }
}
