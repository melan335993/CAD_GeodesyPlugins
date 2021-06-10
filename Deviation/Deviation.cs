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
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

        [CommandMethod("PP_Deviation")]
        public void PP_Deviation()
        {
            FastCAD data = new FastCAD();

            data.listLayers = data.MyTakeLayers();

            PPDialog window = new PPDialog(data);

            if (Application.ShowModalWindow(window) != true)
                return;

            try
            {
                data.MyDeviationPP();
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

            try
            {
                data.MyDeviationST();
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
