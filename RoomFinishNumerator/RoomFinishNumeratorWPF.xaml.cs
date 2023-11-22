using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace RoomFinishNumerator
{
    /// <summary>
    /// Логика взаимодействия для RoomFinishNumeratorWPF.xaml
    /// </summary>
    public partial class RoomFinishNumeratorWPF : Window
    {
        public string RoomFinishNumberingSelectedName;

        public bool ConsiderCeilings;
        public bool ConsiderOpenings;

        public RoomFinishNumeratorWPF()
        {
            InitializeComponent();
        }

        private void btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            RoomFinishNumberingSelectedName = (groupBox_RoomFinishNumbering.Content as System.Windows.Controls.Grid)
                .Children.OfType<RadioButton>()
                .FirstOrDefault(rb => rb.IsChecked.Value == true)
                .Name;
            ConsiderCeilings = checkBox_ConsiderCeilings.IsChecked.Value;
            ConsiderOpenings = checkBox_ConsiderOpenings.IsChecked.Value;

            DialogResult = true;
            Close();
        }

        private void btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        private void RoomFinishNumeratorWPF_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Space)
            {
                RoomFinishNumberingSelectedName = (groupBox_RoomFinishNumbering.Content as System.Windows.Controls.Grid)
                    .Children.OfType<RadioButton>()
                    .FirstOrDefault(rb => rb.IsChecked.Value == true)
                    .Name;
                ConsiderCeilings = checkBox_ConsiderCeilings.IsChecked.Value;
                ConsiderOpenings = checkBox_ConsiderOpenings.IsChecked.Value;

                DialogResult = true;
                Close();
            }

            else if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }
    }
}
