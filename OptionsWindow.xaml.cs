using System;
using System.Collections.Generic;
using System.Globalization;
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
using System.Windows.Shapes;

namespace ExcelToPPT
{
    /// <summary>
    /// Interaction logic for OptionsWindow.xaml
    /// </summary>
    public partial class OptionsWindow : Window
    {

        private Dictionary<string, string> formats;

        public OptionsWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            this.formats = new Dictionary<string, string>();

            CultureInfo culture = new CultureInfo("en-US");

            string[] strFormats =
                DateTimeFormatInfo
                                .CurrentInfo
                                .GetAllDateTimePatterns('d')
                                .Union(
                                    DateTimeFormatInfo
                                    .CurrentInfo
                                    .GetAllDateTimePatterns('D'))
                                .Union(
                                    DateTimeFormatInfo
                                    .CurrentInfo
                                    .GetAllDateTimePatterns('r'))
                                .Union(
                                    DateTimeFormatInfo
                                    .CurrentInfo
                                    .GetAllDateTimePatterns('R'))
                                .Union(
                                    DateTimeFormatInfo
                                    .CurrentInfo
                                    .GetAllDateTimePatterns('m'))
                                .Union(
                                    DateTimeFormatInfo
                                    .CurrentInfo
                                    .GetAllDateTimePatterns('M'))
                                .Union(
                                    new String[]{ "MMM. d, yyyy"}
                                    )
                                .ToArray();
            
            foreach (var customString in strFormats)
            {
                string date = DateTime.Today.ToString(customString, culture);
                this.formats[date] = customString;
                this.TBDateFormat.Items.Add(date);
//                Console.WriteLine("   {0}", customString);
            }
            this.TBDateFormat.SelectedItem = DateTime.Today.ToString( MySettings.Default.SettingDateFormat, culture);
        }

        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            MySettings.Default.SettingDateFormat = formats[this.TBDateFormat.SelectedItem.ToString()];
            MySettings.Default.Save();
            this.Close();
        }

        private void ButtonCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
