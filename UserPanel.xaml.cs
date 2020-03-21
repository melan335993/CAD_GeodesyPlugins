using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace mavCAD
{
    /// <summary>
    /// Логика взаимодействия для UserControl1.xaml
    /// </summary>
    public partial class UserPanel : Window
    {
        public FastCAD _data;

        public UserPanel(FastCAD data)
        {
            InitializeComponent();
            _data = data;
            this.DataContext = data;

            ////////////////////////////////////////

            ComboPlines.Text = "Выбери стены";
            ComboPoints.Text = "Выбери точки";

            foreach (string l in _data.listLayers)
            {
                ComboPlines.Items.Add(l);
                ComboPoints.Items.Add(l);
            }

            ComboPlines.SelectedIndex = 0;
            ComboPoints.SelectedIndex = 0;
        }



        private void Btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            _data.selectedLayerPlines = ComboPlines.Text;
            _data.selectedLayerPoints = ComboPoints.Text;

            {
                string isNum;
                double num;
                isNum = textbox_deviationRed.Text;

                if (textbox_deviationRed.Text == null)
                    num = 0;

                if (double.TryParse(isNum, out num))
                {
                    num = double.Parse(isNum);
                }
                else
                {
                    num = 0;
                }

                _data.deviationRed = num;
            }

            {
                string isNum;
                double num;
                isNum = textbox_deviationRedAngle.Text;

                if (textbox_deviationRedAngle.Text == null)
                    num = 0;


                if (double.TryParse(isNum, out num))
                {
                    num = double.Parse(isNum);
                }
                else
                {
                    num = 0;
                }

                _data.deviationRedAngle = num;
            }

            {
                string isNum;
                double num;
                isNum = textbox_deviationCreate.Text;

                if (textbox_deviationCreate.Text == null)
                    num = 0;


                if (double.TryParse(isNum, out num))
                {
                    num = double.Parse(isNum);
                }
                else
                {
                    num = 0;
                }

                _data.deviationCreate = num;
            }

            {
                string isNum;
                double num;
                isNum = textbox_blockScale.Text;

                if (textbox_blockScale.Text == null)
                    num = 0;

                if (double.TryParse(isNum, out num))
                {
                    num = double.Parse(isNum);
                }
                else
                {
                    num = 0;
                }

                _data.blockScale = num;
            }

            this.DialogResult = true;
            this.Close();
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
               
    }
}
