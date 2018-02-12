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
using System.Windows.Shapes;
using FinanceAnalyticsReporting.ExcelWorksheetTypes;


namespace FinanceAnalyticsReporting.MVVM_Assets
{
    /// <summary>
    /// Interaction logic for ReportSettingsWindow.xaml
    /// </summary>
    public partial class ReportSettingsWindow : Window
    {
        public ReportSettingsWindow()
        {
            InitializeComponent();
        }
        public ReportSettingsWindow(ReportWorksheet reportWorksheet) : this()
        {
            this.ReportPropertyGrid.SelectedObject = reportWorksheet;
            this.ReportPropertyGrid.ShowSearchBox = false;
            this.ReportPropertyGrid.ShowSortOptions = false;
            this.ReportPropertyGrid.ShowAdvancedOptions = true;
        }

    }
}
