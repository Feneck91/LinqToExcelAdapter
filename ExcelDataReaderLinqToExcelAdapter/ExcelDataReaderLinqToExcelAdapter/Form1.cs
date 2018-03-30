using ExcelDataReader;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelDataReaderLinqToExcelAdapter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            listView1.Columns.Clear();

            // Add a column with width 20 and left alignment.
            listView1.Columns.Add("Sheet"     , 100, HorizontalAlignment.Center);
            listView1.Columns.Add("First Name", 120, HorizontalAlignment.Left);
            listView1.Columns.Add("Last Name" , 120, HorizontalAlignment.Left);
            listView1.Columns.Add("Age"       , 50, HorizontalAlignment.Left);
            listView1.Columns.Add("Sexe"      , 100, HorizontalAlignment.Left);
            listView1.Columns.Add("Comment"   , 250, HorizontalAlignment.Left);
            listView1.View = View.Details;
        }

        private void OpenButton_Click(object sender, EventArgs e)
        {
            // Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog.Filter = "Excel File (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm|All Files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.InitialDirectory = Path.Combine(GetExePath(), "..", "..", "Excel");

            DialogResult res = openFileDialog.ShowDialog();

            // Process input if the user clicked OK.
            if (res == DialogResult.OK)
            {
                try
                {
                    var excelReader = new LinqToExcelAdapter.ExcelQueryFactory(openFileDialog.FileName);

                    int iHeaderBeginAt = -1;

                    // Normal binding
                    excelReader.AddMapping<Person>(c => c.FirstName, "First Name");
                    excelReader.AddMapping<Person>(c => c.LastName,  "Last Name");
                    excelReader.AddMapping<Person>(c => c.Comment,   "Comment");
                    // Binding with automatic conversion from double to int
                    excelReader.AddMapping<Person>(c => c.Age,       "Age");
                    //excelReader.AddMapping<Person>(c => c.Age,       "Age", null, true); <-- For version 3 : if replace by this line, if the column is not found, no rows are returns (sheet number 3)
                    // Binding with manual conversion from M/F to enum
                    excelReader.AddMapping<Person>(c => c.Sexe,      "Sexe",        value => (value.ToUpper() == "M") ? Person.eSexe.eSexeMale : (value.ToUpper() == "F" ? Person.eSexe.eSexeFemale : Person.eSexe.eSexeUnknown));

                    listView1.Items.Clear();
                    foreach (var strSheetName in excelReader.GetWorksheetNames())
                    {
                        if (strSheetName == "DataSheet")
                        {
                            excelReader.IsAutoDetectFirstRowForMapping = false;
                            // Find the row where begin the header
                            for (int iBeginLine = 1;iBeginLine < 100;iBeginLine++)
                            {
                                var query = from x
                                            in excelReader.WorksheetRange<Person>(String.Format("A{0}",iBeginLine),String.Format("AZ{0}",iBeginLine + 1), "DataSheet")
                                            select x;

                                Person line = query.First();
                                if (line != null && line.FirstName != null && line.FirstName.Length > 0)
                                {   // Found!
                                    iHeaderBeginAt = iBeginLine;
                                    break;
                                }
                                // Not found, new line
                            }

                            if (iHeaderBeginAt != -1)
                            {
                                // Read all
                                foreach (Person excelRow in (from x in excelReader.WorksheetRange<Person>(String.Format("A{0}",iHeaderBeginAt),"AZ65535","DataSheet")
                                                             where x.Age < 40 || x.Age > 50
                                                             select x))
                                {
                                    ListViewItem lvi = new ListViewItem(strSheetName);
                                    lvi.SubItems.Add(excelRow.FirstName);
                                    lvi.SubItems.Add(excelRow.LastName);
                                    lvi.SubItems.Add(excelRow.Age.ToString());
                                    lvi.SubItems.Add(excelRow.Sexe.ToString());
                                    lvi.SubItems.Add(excelRow.Comment);
                                    listView1.Items.Add(lvi);
                                }
                            }
                        }
                        else
                        {
                            excelReader.IsAutoDetectFirstRowForMapping = true; // <-- If removed, it don't work !
                            foreach (Person excelRow in (from x in excelReader.Worksheet<Person>(strSheetName)
                                                         select x))
                            {
                                ListViewItem lvi = new ListViewItem(strSheetName);
                                lvi.SubItems.Add(excelRow.FirstName);
                                lvi.SubItems.Add(excelRow.LastName);
                                lvi.SubItems.Add(excelRow.Age.ToString());
                                lvi.SubItems.Add(excelRow.Sexe.ToString());
                                lvi.SubItems.Add(excelRow.Comment);
                                listView1.Items.Add(lvi);
                            }
                        }
                    }
                }
                catch(Exception _ex)
                {
                    MessageBox.Show(String.Format("Error: {0}", _ex.Message), "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public static String            GetExePath()
        {
            return Path.GetDirectoryName(GetExeFullPath());
        }

        public static String            GetExeFullPath()
        {
            String strExePath = "";
            Assembly entryAssembly = System.Reflection.Assembly.GetEntryAssembly();
            if (entryAssembly == null)
            {
                strExePath = Process.GetCurrentProcess().MainModule.FileName;
            }
            else
            {
                strExePath = entryAssembly.Location;
            }
            return strExePath;
        }
    }
}
