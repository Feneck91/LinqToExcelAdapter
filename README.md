# LinqToExcelAdapter
Based on ExcelDataReader it allow to use it with linq like LinqToExcel but work in 64 bits

Wrapper classes around ExcelDataReader to make it work like LinqToExcel (not fully implemented)

This project need ExcelDataReader to work, you can find it here : https://github.com/ExcelDataReader/ExcelDataReader
Base on first on source code : http://codereview.stackexchange.com/questions/47227/reading-data-from-excel-sheet-with-exceldatareader
Linq implementation based on : https://msdn.microsoft.com/en-us/library/bb546158.aspx

Problem:
---------
When you compile your software in 64 bits and use LinqToExcel with Office installed in 32 bits mode, you cannot use LinqToExcel.
When you compile your software in 32 bits and use LinqToExcel with Office installed in 64 bits mode, you cannot use LinqToExcel.

Why ? because it use a driver (installed with Office) that must be in the same target a your software.
Of course, installing 32 and 64 bits drivers is not easy (there are some workaroud) and not always possible (in compagny for example), or
for public users.

Idea ? Rewrite all code that use LinqToExcel in my software or used another excel reader and make an adapter to work exactly like LinqToExcel.
It is this second option that I have tried to make.
Of course, for the moment, all LinqToExcel interface is not implemented but 3 functions :
Avalaible:
- WorksheetRangeNoHeader
- WorksheetRange
- GetColumnNames
- GetWorksheetNames

You can use Linq to query into excel, make mapping to automatically fill class on each row.
Mapping with column's name is not always the first row, it can be anywhere and can be found automatically if needed.
You can add mandatories columns header to force it to find at least some colums.

Probably it is a good start to make your library more efficient and simple to use.

To install ExcelDataReader 3.4.0 with Package Manager Console :
   Install-Package ExcelDataReader -Version 3.4.0
   Install-Package ExcelDataReader.DataSet

Source code needed:
-------------------
Only LinqToExcelAdapter.cs file is needed to make it work like LinqToExcel.

Simple sample:
---------------
class Person
{
    public enum eSexe
    {
        eSexeUnknown,
        eSexeMale,
        eSexeFemale
    };

    public String FirstName     { set; get; }
    public String LastName      { set; get; }
    public int Age              { set; get;}
    public eSexe Sexe           { set; get;}
}

try
{
    var excelReader = new LinqToExcelAdapter.ExcelQueryFactory("<file path>");
    // Normal binding
    excelReader.AddMapping<Person>(c => c.FirstName, "First Name", null);
    excelReader.AddMapping<Person>(c => c.LastName,  "Last Name",  null, true); // Name colmumn is mandatory
    excelReader.AddMapping<Person>(c => c.Age,       "Age");
    excelReader.AddMapping<Person>(c => c.Sexe,      "Sexe",        value => (value.ToUpper() == "M") ? Person.eSexe.eSexeMale : (value.ToUpper() == "F" ? Person.eSexe.eSexeFemale : Person.eSexe.eSexeUnknown));

    foreach (var strSheetName in excelReader.GetWorksheetNames())
    {
        excelReader.IsAutoDetectFirstRowForMapping = true; // Auto detect columns if not int the first row (optionnal)
        // Read all except person who have between 40 and 50 years old
        foreach (Person excelRow in (from x in excelReader.WorksheetRange<Person>(String.Format("A{0}",iHeaderBeginAt),"AZ65535","DataSheet")
                                     where x.Age < 40 || x.Age > 50
                                     select x))
        {
            // use excelRow.FirstName / excelRow.LastName / excelRow.Age / excelRow.Sexe
        }
    }
}
catch(Exception _ex)
{
    MessageBox.Show(String.Format("Error: {0}", _ex.Message), "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
}
