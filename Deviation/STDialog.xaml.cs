using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace mavCAD
{
    /// <summary>
    /// Логика взаимодействия для UserControl1.xaml
    /// </summary>
    public partial class STDialog : Window
    {
        public FastCAD _data;

        public STDialog(FastCAD data)
        {
            InitializeComponent();
            _data = data;
            this.DataContext = data;

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

            _data.deviationRed = ParseToDoubleDotOrComma(textbox_deviationRed.Text);
            _data.deviationRedAngle = ParseToDoubleDotOrComma(textbox_deviationRedAngle.Text);
            _data.deviationCreate = ParseToDoubleDotOrComma(textbox_deviationCreate.Text);
            _data.blockScale = ParseToDoubleDotOrComma(textbox_blockScale.Text);
            
            this.DialogResult = true;
            this.Close();
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        /// <summary>
        /// Обработчик ввода с клавиатуры, разрешается вводить только цифры
        /// </summary>
        private void textbox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            TextBox textBox = sender as TextBox;

            if (!(Char.IsDigit(e.Text, 0) || ((e.Text == ".") || (e.Text == ","))
                && ((!textBox.Text.Contains(".") || !textBox.Text.Contains(",")) && textBox.Text.Length != 0)))
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// Парсим строку в дабл независимо точка или запятая как разделитель
        /// </summary>
        double ParseToDoubleDotOrComma(string inputStr)
        {
            double result;

            if (!double.TryParse(inputStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("ru-RU"), out result))
            {
                if (!double.TryParse(inputStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("en-US"), out result))
                {
                    return double.NaN;
                }
            }

            return result;
        }

    }
}
