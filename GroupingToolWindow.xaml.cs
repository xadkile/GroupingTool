using Microsoft.Office.Interop.Excel;
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
using System.Collections.ObjectModel;
using System.ComponentModel;
using GroupingTool.model;
using System.Collections.Immutable;

namespace GroupingTool {


    public class SortLabelItem : INotifyPropertyChanged {
        public string content { get; set; }

        private bool _isSelected = false;
        public bool isSelected {
            get { return _isSelected; }
            set {
                _isSelected = value;
                this.OnPropertyChanged("IsSelected");
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string strPropertyName) {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(strPropertyName));
        }
        public static SortLabelItem fromStr(string content) {
            return new SortLabelItem { content = content, isSelected = false };
        }
    }
    /// <summary>
    /// Interaction logic for GroupingToolWindow.xaml
    /// </summary>
    public partial class GroupingToolWindow : System.Windows.Window {
        Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
        ObservableCollection<string> dataSheetComboBoxLabelList = new ObservableCollection<string>();
        ObservableCollection<string> sortByComboBoxLabelList = new ObservableCollection<string>();
        ObservableCollection<string> groupByComboBoxLabelList = new ObservableCollection<string>();

        public GroupingToolWindow(){
            InitializeComponent();
            this.dataSheetComboBox.ItemsSource = this.dataSheetComboBoxLabelList;
            this.groupByComboBox.ItemsSource = this.groupByComboBoxLabelList;
            this.sortByListBox.ItemsSource = this.sortByComboBoxLabelList;
            populateDataSheetList();
        }

        private Workbook getCurrentWorkbook() {
            return this.app.ActiveWorkbook;
        }

        private Tuple<bool,string> validateInput() {
            StringBuilder builder = new StringBuilder();
            bool hasProblem = false;
            if (!this.validateDataSheet()) {
                builder.Append("-Invalid data sheet\n");
                hasProblem = true;
            }

            if (!this.validateLabelRange()) {
                builder.Append("-Invalid label range\n");
                hasProblem = true;
            }

            if (!this.validFromRow()) {
                builder.Append("-Invalid from row\n");
                hasProblem = true;
            }

            if (!this.validToRow()) {
                builder.Append("-Invalid to row\n");
                hasProblem = true;
            }

            if (!this.validateGroupByComboBox()) {
                builder.Append("-Invalid group-by field\n");
                hasProblem = true;
            }
            
            return new Tuple<bool, string>(hasProblem, builder.ToString());
        }

        private bool validateDataSheet() {
            return this.validateComboBox(this.dataSheetComboBox);
        }

        private bool validateGroupByComboBox() {
            return this.validateComboBox(this.groupByComboBox);
        }

        private bool validFromRow() {
            return this.validateTextNumber(this.fromRowTextBox);
        }

        private bool validToRow() {
            return this.validateTextNumber(this.toRowTextBox);
        }

        private bool validateLabelRange() {
            return !String.IsNullOrEmpty(this.labelRangeTextBox.Text);
        }

        private bool validateTextNumber(System.Windows.Controls.TextBox tb) {
            string textNumber = tb.Text;
            try {
                Int32.Parse(textNumber);
                return true;
            }catch(Exception exception) {
                return false;
            }
        }

        private bool validateComboBox(ComboBox cb) {
            return !(cb.SelectedIndex == -1);
        }

        private void okButton_Click(object sender, RoutedEventArgs e) {
            Tuple<bool, string> check = this.validateInput();
            bool hasProblem = check.Item1;
            if (hasProblem) {
                MessageBox.Show(check.Item2);
            } else {
                //TODO run the logic
                string dataSheetName = this.dataSheetComboBoxLabelList[this.dataSheetComboBox.SelectedIndex];
                string labelRangeAddress = this.labelRangeTextBox.Text;
                List<string> sortFlagList = new List<string>();
                string groupByFlag = this.groupByComboBox.SelectedValue.ToString();
                foreach (String item in this.sortByListBox.SelectedItems) {
                    sortFlagList.Add(item);
                }
                //ImmutableList<string> sortFlagList = this.sortByComboBoxLabelList.ToImmutableList();
                int fromRow = this.getFromRow();
                int toRow = this.getToRow();
                //int trueFromRow = this.getTrueFromRow(fromRow);
                //int trueToRow = this.getTrueToRow(toRow);
                InputFacade inputFacade = new InputFacade(dataSheetName, labelRangeAddress, groupByFlag, sortFlagList.ToImmutableList(), fromRow, toRow);
                Either<Exception,object> runResult = MainLogic.run(inputFacade.toModel());
                if (runResult.isOk()) {
                    this.Close();
                } else {
                    MessageBox.Show(runResult.getException().Message);
                }
            }
        }
       

        private int getNumberFromTextBox(System.Windows.Controls.TextBox textBox,string errMessage) {
            string text = textBox.Text;
            try {
                int rt = Int32.Parse(text);
                return rt;
            } catch (Exception exception) {
                MessageBox.Show(errMessage);
            }
            return 0;
        }

        private int getFromRow() {
            return this.getNumberFromTextBox(this.fromRowTextBox, "Invalid from row value");
        }

        private int getToRow() {
            return this.getNumberFromTextBox(this.toRowTextBox, "Invalid to row value");
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e) {
            this.Close();
        }

        private void groupByComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {

        }

        private void populateDataSheetList() {
            this.dataSheetComboBoxLabelList.Clear();
            foreach(Worksheet ws in this.getCurrentWorkbook().Worksheets) {
                this.dataSheetComboBoxLabelList.Add(ws.Name);
            }
        }

        private void dataSheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            this.refreshForm();
        }

        /**
         * refresh this form
         */ 
        private void refreshForm() {
            
            bool thereAreSheets = this.dataSheetComboBoxLabelList.Count != 0;
            bool oneSheetIsSelected = this.dataSheetComboBox.SelectedIndex != -1;

            if (thereAreSheets && oneSheetIsSelected) {
                // update group by combobox
                this.groupByComboBoxLabelList.Clear();
                this.sortByComboBoxLabelList.Clear();
                // get the selected worksheet
                String selectedWSName = this.dataSheetComboBoxLabelList[this.dataSheetComboBox.SelectedIndex];
                Worksheet selectedSheet = (Worksheet)this.getCurrentWorkbook().Worksheets[selectedWSName];
                try {
                    bool labelRangeIsNotNull = !String.IsNullOrEmpty(this.labelRangeTextBox.Text);
                    if (labelRangeIsNotNull) {
                        Range labelRange = selectedSheet.Range[this.labelRangeTextBox.Text];
                        bool labelRangeRepresentARealRange = labelRange.Rows.Count != 1;
                        if (labelRangeRepresentARealRange) {
                            MessageBox.Show("Label range should be a single row");
                        } else {
                            foreach (Range cell in labelRange.Cells) {
                                string cellValue = Convert.ToString(cell.Value2);
                                if (!String.IsNullOrEmpty(cellValue)) {
                                    this.groupByComboBoxLabelList.Add(cellValue);
                                    this.sortByComboBoxLabelList.Add(cellValue);
                                }
                            }
                        }
                    }
                    
                } catch (Exception exp) {
                    MessageBox.Show("Invalid label range");
                }
            }
        }

        private void labelRangeTextBox_KeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Enter) {
                this.refreshForm();
            }
        }

        private void groupingToolWindowName_KeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Escape) {
                this.Close();
            }
        }

        private void labelRangeTextBox_IsKeyboardFocusedChanged(object sender, DependencyPropertyChangedEventArgs e) {
            this.refreshForm();
        }

        private void refreshFormWithSheetSelectionCheck() {

        }
    }
}
