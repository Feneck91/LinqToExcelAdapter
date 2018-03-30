using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections;
using System.Reflection;
using System.Linq.Expressions;
using System.ComponentModel;

namespace ExcelDataReader
{
    /// <summary>
    /// Wrapper classes around ExcelDataReader to make it work like LinqToExcel.
    /// 
    /// Base on first on source code : http://codereview.stackexchange.com/questions/47227/reading-data-from-excel-sheet-with-exceldatareader
    /// Linq implementation based on : https://msdn.microsoft.com/en-us/library/bb546158.aspx
    ///
    /// Problem:
    /// When you compile your software in 64 bits and use LinqToExcel with Office installed in 32 bits mode, you cannot use LinqToExcel.
    /// When you compile your software in 32 bits and use LinqToExcel with Office installed in 64 bits mode, you cannot use LinqToExcel.
    ///
    /// Why ? because it use a driver (installed with Office) that must be in the same target a your software.
    /// Of course, installing 32 and 64 bits drivers is not easy (there are some workaroud) and not always possible (in compagny for example), or
    /// for public users.
    /// 
    /// Idea ? Rewrite all code that use LinqToExcel in my software or used another excel reader and make an adapter to work exactly like LinqToExcel.
    /// It is this second option that I have tried to make.
    /// Of course, for the moment, all LinqToExcel interface is not implemented.
    /// Avalaible:
    /// - WorksheetRangeNoHeader
    /// - WorksheetRange
    /// - GetColumnNames
    /// - GetWorksheetNames
    ///
    /// You can use Linq to query into excel, make mapping to automatically fill class on each row.
    /// Mapping with column's name is not always the first row, it can be anywhere and can be found automatically if needed.
    /// You can add mandatories columns header to force it to find at least some colums.
    /// 
    /// Probably it is a good start to make your library more efficient and simple to use.
    /// 
    /// To install ExcelDataReader 3.4.0 with Package Manager Console :
    ///    Install-Package ExcelDataReader -Version 3.4.0
    ///    Install-Package ExcelDataReader.DataSet
    /// 
    /// Version 1.x are based on ExcelDataReader 2.x
    ///
    /// ------------------
    /// Version history :
    /// ------------------
    /// Version 1.0      : http://ti1ca.com/8ci1u2xm-ExcelDataReaderLinqToExcelAdapter-ExcelDataReaderLinqToExcelAdapter.zip.html
    /// Version 1.1 (v2) : http://ti1ca.com/b8v7pika-ExcelDataReaderLinqToExcelAdapter-v-2-ExcelDataReaderLinqToExcelAdapter_v_2.zip.html
    ///                   • Add Function : Worksheet
    ////                  • Throw exception if Worksheet not found.
    /// Version 1.2 (v3) : https://ti1ca.com/6b4wsagf-ExcelDataReaderLinqToExcelAdapter-v-3-ExcelDataReaderLinqToExcelAdapter_v_3.zip.html
    ///                   • Allow to open the excel file even it is already opened into Excel.
    ///                   • Add LinqToExcelAdapter.ExcelQueryFactory.GetWorksheetNames() function that return all Worksheet names.
    ///                   • Allow to auto-detect the header if not into the first row ('IsAutoDetectFirstRowForMapping' parameter into the LinqToExcelAdapter.ExcelQueryFactory function).
    ///                   • Allow to force column to be mapped into LinqToExcelAdapter.ExcelQueryFactory.AddMapping ('IsMandatory' parameter, all rows are ignored while this column is not found, can work with IsAutoDetectFirstRowForMapping parameter (or not)).
    /// Version 2.0 : 
    ///                   • Same as Version 1.2 but work for ExcelDataReader 3.4.0
    /// </summary>
    /// <author>Stéphane Château (Feneck91)</author>
    public static class LinqToExcelAdapter
    {
        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                          ExcelQueryFactory                                                          |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// Classe use to simulate LinqToExcel.ExcelQueryFactory
        /// </summary>
        public class ExcelQueryFactory
        {
            /// <summary>
            /// The DataSet.
            /// 
            /// Never call this dataset getter directly, call GetExcelDataDataSet() that check if excel file is correctly opened.
            /// Only GetExcelDataDataSet() call the getter directly.
            /// </summary>
            private DataSet DataSet                             { get; set; }      = null;

            /// <summary>
            /// The mapping list.
            /// </summary>
            private  List<AMappingBase> ListMapping             { get; set; }      = new List<AMappingBase>();

            /// <summary>
            /// Should replace columns names by underscore ?
            /// 
            /// Used to be compatible with LinqToExcel.
            /// </summary>
            public bool ReplaceCariageReturnByUnderscore        { get; set; }

            /// <summary>
            /// Option that allow to auto-detect first row where column are found.
            /// 
            /// This option is not compatible with LinqToExcel. If false, it work like LinqToExcel (by default).
            /// </summary>
            public bool IsAutoDetectFirstRowForMapping          { get; set; }

            /// <summary>
            /// Constructor.
            /// 
            /// The _bReplaceCariageReturnByUnderscore is used to replace all column's name (into mapping) by '_' like in LinqToExcel library).
            /// </summary>
            /// <param name="_strExcelFilePath">Excel filename to load</param>
            /// <param name="_bReplaceCariageReturnByUnderscore">Should replace carriage return column's name by underscore ?</param>
            /// <param name="_bIsAutoDetectFirstRowForMapping">Allow to auto-detect first row where column binding are found, false to work like LinqToExcel.</param>
            public ExcelQueryFactory(string _strExcelFilePath, bool _bReplaceCariageReturnByUnderscore = true, bool _bIsAutoDetectFirstRowForMapping = false)
            {
                if (!File.Exists(_strExcelFilePath))
                {
                    throw new FileNotFoundException(String.Format("The {0} file doesn't exists!", _strExcelFilePath));
                }
                ReplaceCariageReturnByUnderscore = _bReplaceCariageReturnByUnderscore;
                IsAutoDetectFirstRowForMapping   = _bIsAutoDetectFirstRowForMapping;
                DataSet                          = ReadExcelDataSet(_strExcelFilePath);
            }

            /// <summary>
            /// Get the list of worksheet names.
            /// </summary>
            /// <returns>A list with the names of all the worksheet.</returns>
            public List<String> GetWorksheetNames()
            {
                List<String> lstWorksheetNames = null;
                lstWorksheetNames = (from DataTable worksheet in GetExcelDataDataSet().Tables select worksheet.TableName).ToList();

                return lstWorksheetNames;
            }


            /// <summary>
            /// Get the data from worksheet name and range.
            /// </summary>
            /// <param name="_strStartRange">Cell of started range.</param>
            /// <param name="_strEndRange">Cell of ended range.</param>
            /// <param name="_strWorksheetName">Sheet's name.</param>
            /// <returns>A queryable object to be used with Linq</returns>
            public ExcelQueryable<NoRowHeader> WorksheetRangeNoHeader(string _strStartRange, string _strEndRange, string _strWorksheetName)
            {
                ExcelQueryable<NoRowHeader> queryNoHeaderRet = null;
                DataTable dtWorkSheet = GetExcelWorksheet(_strWorksheetName);

                if (dtWorkSheet != null)
                {
                    queryNoHeaderRet = new ExcelQueryable<NoRowHeader>(new ReaderOptions(dtWorkSheet.Rows, new ExcelCellRange(new ExcelCellAddress(_strStartRange), new ExcelCellAddress(_strEndRange)), null, ReplaceCariageReturnByUnderscore, IsAutoDetectFirstRowForMapping));
                }

                return queryNoHeaderRet;
            }

            /// <summary>
            /// Get a list of automatically binded from colum's name.
            /// 
            /// The first row of range is read to create the colum's mapping, the data are read from next line. 
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="_strStartRange">Start range of excel rows / columns to read.</param>
            /// <param name="_strEndRange">End range of excel rows / columns to read.</param>
            /// <param name="_strWorksheetName">Worksheet name.</param>
            /// <returns>A queryable object to be used with Linq</returns>
            public QueryableLinqToExcelAdapterData<T> WorksheetRange<T>(string _strStartRange, string _strEndRange, string _strWorksheetName) where T: class
            {
                QueryableLinqToExcelAdapterData<T> queryRet = null;
                DataTable dtWorkSheet = GetExcelWorksheet(_strWorksheetName);

                if (dtWorkSheet != null)
                {
                    // Clearing the binding, it can be the binding of another sheet
                    ListMapping.ForEach(item => item.ClearBinding());
                    queryRet = new QueryableLinqToExcelAdapterData<T>(new ReaderOptions(dtWorkSheet.Rows, new ExcelCellRange(new ExcelCellAddress(_strStartRange), new ExcelCellAddress(_strEndRange)), ListMapping, ReplaceCariageReturnByUnderscore, IsAutoDetectFirstRowForMapping));
                }
                else
                {
                    throw new Exception(String.Format("The excel worksheet {0} doesn't exists!", _strWorksheetName));
                }

                return queryRet;
            }

            /// <summary>
            /// Get a list of automatically binded from colum's name.
            /// 
            /// The first row of range is read to create the colum's mapping, the data are read from next line. 
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="_strWorksheetName">Worksheet name.</param>
            /// <returns>A queryable object to be used with Linq</returns>
            public QueryableLinqToExcelAdapterData<T> Worksheet<T>(string _strWorksheetName) where T: class
            {
                QueryableLinqToExcelAdapterData<T> queryRet = null;
                DataTable dtWorkSheet = GetExcelWorksheet(_strWorksheetName);

                if (dtWorkSheet != null)
                {
                    // Clearing the binding, it can be the binding of another sheet
                    ListMapping.ForEach(item => item.ClearBinding());
                    queryRet = new QueryableLinqToExcelAdapterData<T>(new ReaderOptions(dtWorkSheet.Rows, null, ListMapping, ReplaceCariageReturnByUnderscore, IsAutoDetectFirstRowForMapping));
                }
                else
                {
                    throw new Exception(String.Format("The excel worksheet {0} doesn't exists!", _strWorksheetName));
                }

                return queryRet;
            }

            /// <summary>
            /// Get a list of colum's name (beginning of the range).
            /// 
            /// The first row of range is read to create the colum's content to return.
            /// The list of columns returned could be on more than one line.
            /// </summary>
            /// <typeparam name="T">Column's type.</typeparam>
            /// <param name="_strWorksheetName">Worksheet name.</param>
            /// <param name="_strtRange">Range of excel rows / columns to read, null to take all.</param>
            /// <returns>A list of column's values.</returns>
            public IEnumerable<string> GetColumnNames(string _strWorksheetName, string _strtRange = null)
            {
                var list = new List<string>();
                DataTable dtWorkSheet = GetExcelWorksheet(_strWorksheetName);

                if (dtWorkSheet != null)
                {
                    ExcelCellRange range = new ExcelCellRange(_strtRange);

                    for (int iIndexRow = (int) range.GetFirstRowOfRange() - 1;
                            iIndexRow <= Math.Min(range.GetLastRowOfRange() - 1, dtWorkSheet.Rows.Count - 1);
                            ++iIndexRow)
                    {
                        var row = dtWorkSheet.Rows[iIndexRow].ItemArray;

                        for (long lIndexColumn = range == null ? 0 : range.GetFirstColumnOfRange() - 1;
                                lIndexColumn <= Math.Min((range == null ? long.MaxValue : range.GetLastColumnOfRange() - 1), row.Length - 1);
                                ++lIndexColumn)
                        {
                            string strValue;
                            strValue = Helper.GetColumnName(row[lIndexColumn], ReplaceCariageReturnByUnderscore);
                            list.Add(strValue);
                        }
                    }
                }

                return list;
            }

            /// <summary>
            /// Add mapping into the list of mapping items.
            /// </summary>
            /// <typeparam name="T">Type of class to map.</typeparam>
            /// <param name="_mapping">Mapping function to map property's class to excel's column data.</param>
            /// <param name="_strColumnName">Excel column's name where data is read.</param>
            /// <param name="_funcConvertValue">Function to convert value if needed (if not directly convert from string to the target field type automatically).</param>
            /// <param name="_bIsMandatory">If the mandatory is true, the binding will not work if this column is not found. false to work like LinqToExcel.</param>
            public void AddMapping<T>(Expression<Func<T, object>> _mapping, string _strColumnName, Func<string, object> _funcConvertValue = null, bool _bIsMandatory= false)
            {
                ListMapping.Add(new Mapping<T>(_mapping, _strColumnName, _funcConvertValue, _bIsMandatory));
            }

            #region Privates functions
            /// <summary>
            /// Get the excel dataset.
            /// </summary>
            /// <param name="_strExcelFilePath">File path.</param>
            /// <returns>The excel datast if no error, exception else.</returns>
            private DataSet ReadExcelDataSet(string _strExcelFilePath)
            {
                using (FileStream fileStream = File.Open(_strExcelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IExcelDataReader dataReader = ExcelReaderFactory.CreateReader(fileStream);
                    return dataReader.AsDataSet(new ExcelDataSetConfiguration() {  ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = false }});
                }
            }

            /// <summary>
            /// The the excel dataset (internal use)
            /// </summary>
            /// <returns>The dataset.</returns>
            private DataSet GetExcelDataDataSet()
            {
                if (DataSet == null)
                {
                    throw new FileNotFoundException("The excel file is not loaded!");
                }

                return DataSet;
            }

            /// <summary>
            /// Get the worksheet from name.
            /// </summary>
            /// <param name="_strWorkSheetName">Worksheet name.</param>
            /// <returns>A datatable if found, null else.</returns>
            private DataTable GetExcelWorksheet(string _strWorkSheetName)
            {
                DataTable workSheet = GetExcelDataDataSet().Tables[_strWorkSheetName];

                return workSheet;
            }
            #endregion
        }

        /// <summary>
        /// Class that record all options in a single class.
        /// </summary>
        public class ReaderOptions
        {
            /// <summary>
            /// Current rows.
            /// </summary>
            public DataRowCollection DataRows                   { get; private set; }

            /// <summary>
            /// Current cells range.
            /// </summary>
            public ExcelCellRange Range                         { get; private set; }

            /// <summary>
            /// The mapping list.
            /// 
            /// Null if not used.
            /// </summary>
            public List<AMappingBase> ListMapping               { get; private set; }

            /// <summary>
            /// Should replace columns names by underscore ?
            /// 
            /// Used to be compatible with LinqToExcel.
            /// </summary>
            public bool ReplaceCariageReturnByUnderscore        { get; private set; }

            /// <summary>
            /// Option that allow to auto-detect first row where column are found.
            /// 
            /// This option is not compatible with LinqToExcel. If false, it work like LinqToExcel (by default).
            /// </summary>
            public bool IsAutoDetectFirstRowForMapping          { get; private set; }

            /// <summary>
            /// Create Options from parameters.
            /// </summary>
            /// <param name="_dataRows">Rows, cannot be null.</param>
            /// <param name="_range">Range, cannot be null.</param>
            /// <param name="_listMapping">Mapping.</param>
            /// <param name="_bReplaceCariageReturnByUnderscore">Should replace carriage return column's name by underscore ?</param>
            /// <param name="_bIsAutoDetectFirstRowForMapping">Allow to auto-detect first row where column binding are found, false to work like LinqToExcel.</param>
            public ReaderOptions(DataRowCollection _dataRows, ExcelCellRange _range, List<AMappingBase> _listMapping, bool _bReplaceCariageReturnByUnderscore, bool _bIsAutoDetectFirstRowForMapping)
            {
                DataRows                         = _dataRows;
                Range                            = _range;
                ListMapping                      = _listMapping;
                ReplaceCariageReturnByUnderscore = _bReplaceCariageReturnByUnderscore;
                IsAutoDetectFirstRowForMapping   = _bIsAutoDetectFirstRowForMapping;
            }
        }
        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                             AMappingBase                                                            |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// Base class for mapping.
        /// 
        /// It let data assign only with Func.
        /// </summary>
        public abstract class AMappingBase
        {
            /// <summary>
            /// Expression used to get the field where to set value.
            /// </summary>
            public Expression           ExpressionGetField                  { get; private set; }

            /// <summary>
            /// Column name.
            /// </summary>
            public String               ColumnName                          { get; private set; }

            /// <summary>
            /// Function to convert value if needed (if not directly convert from string to the target field type).
            /// </summary>
            public Func<string, object> FuncConvertValue                    { get; private set; }

            /// <summary>
            /// Fast property.
            /// 
            /// Cannot use PropertyInfo class to assign velue because it is really to slow (under debugger, VERY slow).
            /// So, I use this FastProperty class.
            /// </summary>
            public FastProperty         FastProperty                        { get; private set; }

            /// <summary>
            /// A converter used to convert value into target type.
            /// Not always used, but is filled at creation to optimize the speed.
            /// </summary>
            public TypeConverter        TypeConverter                       { get; private set; }

            /// <summary>
            /// Set the index of the column for the binding.
            /// 
            /// -1 : no binding found.
            /// </summary>
            public int                  IndexColumnBinding                  { get; private set; } = -1;

            /// <summary>
            /// Is this column binding is mandatory ?.
            /// 
            /// false to work like linqToExcel, else if if true and the column binding is not found, the row is ignored.
            /// </summary>
            public bool                 IsMandatory                         { get; private set; } = false;

            /// <summary>
            /// Check if this binding is correctly found (IndexColumnBinding different from -1)
            /// </summary>
            public bool IsBinded
            {
                get
                {   // Return true if it correctly binded
                    return IndexColumnBinding != -1;
                }
            }

            /// <summary>
            /// Check if this binding is mandatory, it MUST have an IndexColumnBinding different from -1
            /// </summary>
            public bool MandatoryBadBinding
            {
                get
                {   // Return true if it is mandatory and not binded
                    return IsMandatory && IndexColumnBinding == -1;
                }
            }

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="_expGetField">Target Field to fill.</param>
            /// <param name="_strColumnName">Excel column's name where read data.</param>
            /// <param name="_funcConvertValue">Function to convert value if needed (if not directly convert from string to the target field type automatically).</param>
            /// <param name="_bIsMandatory">If the mandatory is true, the binding will not work if this column is not found. false to work like LinqToExcel.</param>
            public AMappingBase(Expression _expGetField, String _strColumnName, Func<string, object> _funcConvertValue, bool _bIsMandatory)
            {
                ExpressionGetField  = _expGetField;
                ColumnName          = _strColumnName;
                FuncConvertValue    = _funcConvertValue;
                IsMandatory         = _bIsMandatory;

                // Init property name.
                ComputeIntrospection(null);
            }

            /// <summary>
            /// Clear the current binding.
            /// </summary>
            public void ClearBinding()
            {
                IndexColumnBinding = -1;
            }

            /// <summary>
            /// Compute the property name of the target field to fill.
            /// </summary>
            /// <param name="_type">Type of target object, filled by derived class. This class call this function this null parameter.</param>
            protected virtual void ComputeIntrospection(Type _type)
            {
                var exp = (LambdaExpression) ExpressionGetField;

                // exp.Body has 2 possible types
                // If the property type is native, then exp.Body == typeof(MemberExpression)
                // If the property type is not native, then exp.Body == typeof(UnaryExpression) in which 
                // case we can get the MemberExpression from its Operand property
                var mExp = (exp.Body.NodeType == ExpressionType.MemberAccess) ?
                    (MemberExpression)exp.Body :
                    (MemberExpression)((UnaryExpression)exp.Body).Operand;

                PropertyInfo pPropertyInfo = _type.GetProperty(mExp.Member.Name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (pPropertyInfo != null && pPropertyInfo.CanWrite)
                {
                    FastProperty = new FastProperty(pPropertyInfo);

                    TypeConverter converter = TypeDescriptor.GetConverter(pPropertyInfo.PropertyType);
                    if (converter != null)
                    {
                        TypeConverter = converter;
                    }
                }
            }

            /// <summary>
            /// Read header.
            /// </summary>
            /// <param name="_row">row</param>
            /// <param name="_options">Options, contains Range and others informations.</param>
            internal void ReadHeader(DataRow _row, ReaderOptions _options)
            {
                var row = _row.ItemArray;

                for (long lIndexColumn = _options.Range == null ? 0 : _options.Range.GetFirstColumnOfRange() - 1;
                     lIndexColumn <= Math.Min((_options.Range == null ? long.MaxValue : _options.Range.GetLastColumnOfRange() - 1), row.Length - 1);
                     ++lIndexColumn)
                {
                    try
                    {
                        if (ColumnName == Helper.GetColumnName(row[lIndexColumn], _options.ReplaceCariageReturnByUnderscore))
                        {
                            IndexColumnBinding = (int) lIndexColumn;
                            break;
                        }
                    }
                    catch
                    {   // Not converted to string, ignore this column
                    }
                }
            }

            /// <summary>
            /// Make the row assignation.
            /// </summary>
            /// <param name="_dataRow">Row that contains datas.</param>
            /// <param name="_value">target instance of T (derived class).</param>
            internal abstract void Assign(DataRow _dataRow, object _value);
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                             Mapping<T>                                                              |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// Derived class for mapping.
        /// </summary>
        /// <typeparam name="T">Type of the class that will receive row data.</typeparam>
        private class Mapping<T> : AMappingBase
        {
            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="_expGetField">Target Field to fill.</param>
            /// <param name="_strColumnName">Excel column's name where data is read.</param>
            /// <param name="_funcConvertValue">Function to convert value if needed (if not directly convert from string to the target field type automatically).</param>
            /// <param name="_bIsMandatory">If the mandatory is true, the binding will not work if this column is not found. false to work like LinqToExcel.</param>
            public Mapping(Expression<Func<T, object>> _expGetField, String _strColumnName, Func<string, object> _funcConvertValue = null, bool _bIsMandatory = false)
                : base(_expGetField, _strColumnName, _funcConvertValue, _bIsMandatory)
            {
            }

            /// <summary>
            /// Compute the all needed fields to be able to correctly assign the target field.
            /// </summary>
            protected override void ComputeIntrospection(Type _type)
            {   // Nothing to do, only call base class.
                base.ComputeIntrospection(typeof(T));
            }

            /// <summary>
            /// Make the row assignation.
            /// </summary>
            /// <param name="_dataRow">Row that contains datas.</param>
            /// <param name="_value">target instance of T.</param>
            internal override void Assign(DataRow _dataRow, object _value)
            {
                if (IndexColumnBinding != -1 && FastProperty != null)
                {
                    object objColumnValue = _dataRow[IndexColumnBinding];

                    if (FuncConvertValue != null)
                    {   // Overwrite objColumnValue if a convert function is fill.
                        objColumnValue = FuncConvertValue(objColumnValue == null ? "" : objColumnValue.ToString());
                    }

                    if (objColumnValue == null || (objColumnValue.GetType() == FastProperty.Property.PropertyType))
                    {   // Ok, same type, nothing to do, can be directly assigned
                        // or null ? Try to assign but if don't work, don't know what to do
                        try
                        {
                            FastProperty.Set(_value, objColumnValue);
                        }
                        catch
                        {   // not work !
                            // Really strange if the assign not work... (exception into the setter of the class ?)
                        }
                    }
                    else
                    {   // Should convert the value
                        if (TypeConverter != null)
                        {   // Converter has been found
                            try
                            {
                                if (TypeConverter.CanConvertFrom(objColumnValue.GetType()))
                                {
                                    FastProperty.Set(_value, TypeConverter.ConvertFrom(objColumnValue));
                                }
                                else if (TypeConverter.CanConvertFrom(typeof(string)))
                                {
                                    FastProperty.Set(_value, TypeConverter.ConvertFrom(objColumnValue.ToString()));
                                }
                            }
                            catch
                            {   // not work !
                            }
                        }
                    }
                }
            }
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                 QueryableLinqToExcelAdapterDataBase                                                 |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// Base class for queryable.
        /// 
        /// Used to get data from source.
        /// </summary>
        public abstract class QueryableLinqToExcelAdapterDataBase
        {
            /// <summary>
            /// Get the source object.
            /// </summary>
            /// <returns></returns>
            public abstract IQueryable GetQueryableDatas();
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                   QueryableLinqToExcelAdapterData                                                   |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// The Queryable data.
        /// </summary>
        /// <typeparam name="TData"></typeparam>
        public class QueryableLinqToExcelAdapterData<TData> : QueryableLinqToExcelAdapterDataBase
                                                            , IQueryable<TData>
        {
            #region Properties IQueryable
            /// <summary>
            /// Get the provider.
            /// </summary>
            public IQueryProvider Provider                      { get; protected set; }

            /// <summary>
            /// Get the expression.
            /// </summary>
            public Expression Expression                        { get; protected set; }

            /// <summary>
            /// Get Type.
            /// </summary>
            public Type ElementType                             { get { return typeof(TData); } }
            #endregion

            #region Other Properties
            /// <summary>
            /// Get the options to use.
            /// 
            /// Range, mapping, etc.
            /// </summary>
            protected ReaderOptions Options                     { get; set; }
            #endregion

            #region Constructors
            /// <summary> 
            /// This constructor is called by Provider.CreateQuery(). 
            /// </summary> 
            /// <param name="_provider">The provider.</param>
            /// <param name="_expression">The expression.</param>
            public QueryableLinqToExcelAdapterData(LinqToExcelAdapterQueryProvider _provider, Expression _expression)
            {
                if (_provider == null)
                {
                    throw new ArgumentNullException("provider");
                }

                if (_expression == null)
                {
                    throw new ArgumentNullException("expression");
                }

                if (!typeof(IQueryable<TData>).IsAssignableFrom(_expression.Type))
                {
                    throw new ArgumentOutOfRangeException("expression");
                }

                Provider    = _provider;
                Expression  = _expression;
                Options     = _provider.Options;
            }

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="_options">Options to use.</param>
            public QueryableLinqToExcelAdapterData(ReaderOptions _options)
            {
                Options     = _options;

                Provider    = new LinqToExcelAdapterQueryProvider(_options);
                Expression  = Expression.Constant(this);

                if (Options.DataRows == null)
                {
                    throw new ArgumentNullException("QueryableLinqToExcelAdapterData: list of rows is null!");
                }
            }
            #endregion

            #region Enumerators
            /// <summary>
            /// Get the typed enumerator.
            /// </summary>
            /// <returns>The typed enumerator.</returns>
            public IEnumerator<TData> GetEnumerator()
            {
                return (Provider.Execute<IEnumerable<TData>>(Expression)).GetEnumerator();
            }

            /// <summary>
            /// Get the enumerator.
            /// </summary>
            /// <returns>The enumerator.</returns>
            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return (Provider.Execute<System.Collections.IEnumerable>(Expression)).GetEnumerator();
            }
            #endregion

            /// <summary>
            /// Read the row.
            /// 
            /// Create the data and fill it with mapping.
            /// </summary>
            /// <param name="_dataRow">Row data.</param>
            /// <returns>A TData item correctly filled with column mapping.</returns>
            private TData ReadRow(DataRow _dataRow)
            {
                TData val = (TData) Activator.CreateInstance(typeof(TData));

                if (Options.ListMapping != null)
                {   // Just to don't crash but usually ListMapping is not null
                    foreach (var item in Options.ListMapping)
                    {
                        item.Assign(_dataRow, val);
                    }
                }

                return val;
            }

            /// <summary>
            /// Is the row is in the range ?
            /// 
            /// Don't use directly the Range into the  GetQueryableDatas() function because when mapping object,
            /// the first column into range should be detected to manage mapping and this first column is ignored (no data into it).
            /// </summary>
            /// <param name="_row">Row to test.</param>
            /// <returns>true if the row sould be read, false else.</returns>
            public bool IsInRange(DataRow _row)
            {
                if (Options.ListMapping != null)
                {
                    long lRowID = Helper.GetRowID(_row);
                    bool bCheckBinding = false;

                    if (lRowID != 1 && Options.IsAutoDetectFirstRowForMapping)
                    {
                        bCheckBinding = (from item in Options.ListMapping where item.IsBinded select item).Count() == 0;
                    }

                    // Is the row is the first row that should be used as the header to fill datas ?
                    if (   (Options.Range == null && (lRowID == 1 || bCheckBinding))
                        || (Options.Range != null && Options.Range.IsFirstRowInRange(lRowID)))
                    {   // FirstRow for columns
                        foreach (var item in Options.ListMapping)
                        {   // Prepare all mapping to find correct columns index
                            item.ReadHeader(_row, Options);
                        }

                        // Check if all mandatory binding are found
                        if ((from item in Options.ListMapping where item.MandatoryBadBinding select item).Count() > 0)
                        {   // Else clear all
                            foreach(var item in (from item in Options.ListMapping where item.IsBinded select item))
                            {
                                item.ClearBinding();
                            }
                        }

                        return false; // Ignore first columns row (or if not found) : no data in it
                    }
                }

                return Options.Range == null || Options.Range.IsInRange(_row);
            }

            #region Implementation QueryableLinqToExcelAdapterDataBase
            /// <summary>
            /// Get the source objects.
            /// </summary>
            /// <returns></returns>
            public override IQueryable GetQueryableDatas()
            {
                return (from DataRow row
                        in Options.DataRows
                        where IsInRange(row)
                        select ReadRow(row)).AsQueryable();
            }
            #endregion
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                            ExcelQueryable                                                           |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// Class that used to record all rows and make rows queryable.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        public class ExcelQueryable<T> : QueryableLinqToExcelAdapterData<OneRowInfos<T>>
        {
            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="_options">Options to use.</param>
            public ExcelQueryable(ReaderOptions _options)
                : base(_options)
            {
                // Override Expression (very important to have correct type)
                Expression  = Expression.Constant(this);
            }

            #region Implementation QueryableLinqToExcelAdapterDataBase
            /// <summary>
            /// Get the source objects.
            /// </summary>
            /// <returns></returns>
            public override IQueryable GetQueryableDatas()
            {
                return (from DataRow row
                        in Options.DataRows
                        where Options.Range == null || Options.Range.IsInRange(row)
                        select new OneRowInfos<NoRowHeader>(row, Options.Range)).AsQueryable();
            }
            #endregion

        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                             Queryable<T>                                                            |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// Provider for Linq.
        /// </summary>
        public class LinqToExcelAdapterQueryProvider : IQueryProvider
        {
            #region Other Properties
            /// <summary>
            /// Get the options to use.
            /// 
            /// Range, mapping, etc.
            /// </summary>
            public ReaderOptions Options                    { get; set; }
            #endregion

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="_options">Options to use.</param>
            public LinqToExcelAdapterQueryProvider(ReaderOptions _options)
            {
                Options     = _options;
            }

            /// <summary>
            /// Create the query.
            /// </summary>
            /// <param name="_expression">Expression.</param>
            /// <returns>A queryable object.</returns>
            public IQueryable CreateQuery(Expression _expression)
            {
                Type elementType = TypeSystem.GetElementType(_expression.Type);

                try
                {
                    return (IQueryable)Activator.CreateInstance(typeof(QueryableLinqToExcelAdapterData<>).MakeGenericType(elementType), new object[] { this, _expression });
                }
                catch (System.Reflection.TargetInvocationException _tie)
                {
                    throw _tie.InnerException;
                }
            }

            /// <summary>
            /// Queryable's collection-returning standard query operators call this method. 
            /// </summary>
            /// <typeparam name="TResult">Result type.</typeparam>
            /// <param name="_expression">Expression.</param>
            /// <returns></returns>
            public IQueryable<TResult> CreateQuery<TResult>(Expression _expression)
            {
                return new QueryableLinqToExcelAdapterData<TResult>(this, _expression);
            }

            /// <summary>
            /// Execute the expression.
            /// </summary>
            /// <param name="_expression">Expression.</param>
            /// <returns>An object.</returns>
            public object Execute(Expression _expression)
            {
                return LinqToExcelAdapterQueryContext.Execute(_expression, false);
            }

            /// <summary>
            /// Queryable's "single value" standard query operators call this method.
            /// It is also called from QueryableTerraServerData.GetEnumerator().
            /// </summary>
            /// <param name="_expression">The expression.</param>
            /// <returns>The result.</returns>
            public TResult Execute<TResult>(Expression _expression)
            {
                bool IsEnumerable = (typeof(TResult).Name == "IEnumerable`1");

                return (TResult) LinqToExcelAdapterQueryContext.Execute(_expression, IsEnumerable);
            }
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                             Queryable<T>                                                            |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// Query context.
        /// </summary>
        internal class LinqToExcelAdapterQueryContext
        {
            // Executes the expression tree that is passed to it. 
            internal static object Execute(Expression _expression, bool _bIsEnumerable)
            {
                // The expression must represent a query over the data source. 
                if (!IsQueryOverDataSource(_expression))
                {
                    throw new InvalidProgramException("No query over the data source was specified.");
                }

                // Find the constant expression to make the request.
                ConstantExpression constantExpression = new InnerConstantFinder().GetInnerConstant(_expression, value => value.Value is QueryableLinqToExcelAdapterDataBase);
                if (constantExpression != null)
                {
                    QueryableLinqToExcelAdapterDataBase dataQuery = constantExpression.Value as QueryableLinqToExcelAdapterDataBase;
                    if (dataQuery != null)
                    {
                        // Copy the expression tree that was passed in, changing only the first 
                        // argument of the innermost MethodCallExpression.
                        IQueryable query = dataQuery.GetQueryableDatas();
                        
                        Expression newExpressionTree = new ExpressionTreeModifier(query, dataQuery.GetType()).Visit(_expression);

                        // This step creates an IQueryable that executes by replacing Queryable methods with Enumerable methods. 
                        if (_bIsEnumerable)
                        {
                            return query.Provider.CreateQuery(newExpressionTree);
                        }
                        else 
                        {
                            return query.Provider.Execute(newExpressionTree);
                        }
                    }
                }

                throw new NotImplementedException("Expression is not implemented in your query.");
            }

            /// <summary>
            /// ??
            /// </summary>
            /// <param name="expression"></param>
            /// <returns></returns>
            private static bool IsQueryOverDataSource(Expression expression)
            {
                // If expression represents an unqueried IQueryable data source instance, 
                // expression is of type ConstantExpression, not MethodCallExpression. 
                return (expression is MethodCallExpression);
            }
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                             Queryable<T>                                                            |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================

        /// <summary>
        /// Class queryable.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        public abstract class QueryableBase<T> : IQueryable<T>
        {
            #region Properties IQueryable
            /// <summary>
            /// Get the provider.
            /// </summary>
            public IQueryProvider Provider              { get; protected set; }

            /// <summary>
            /// Get the expression.
            /// </summary>
            public Expression Expression                { get; protected set; }

            /// <summary>
            /// Get Type.
            /// </summary>
            public Type ElementType                     { get { return typeof(T); } }
            #endregion

            #region Other Properties
            /// <summary>
            /// Get the options to use.
            /// 
            /// Range, mapping, etc.
            /// </summary>
            protected ReaderOptions Options                     { get; set; }
            #endregion

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="_options">Options to use.</param>
            public QueryableBase(ReaderOptions _options)
            {
                if (_options.DataRows == null)
                {
                    throw new ArgumentNullException("QueryableBase: list of rows is null!");
                }
                if (_options.Range == null)
                {
                    throw new ArgumentNullException("QueryableBase: range is null!");
                }
                Options = _options;
            }

            /// <summary>
            /// Create the enumerator.
            /// 
            /// Used by both function GetEnumerator.
            /// </summary>
            /// <returns>The enumerator.</returns>
            private IEnumerator<T> GetPrivateEnumerator()
            {
                return GetProtectedEnumerable().GetEnumerator();
            }

            /// <summary>
            /// Get the typed enumerator.
            /// </summary>
            /// <returns>The enumerator.</returns>
            public virtual IEnumerator<T> GetEnumerator()
            {
                return GetPrivateEnumerator();
            }

            /// <summary>
            /// Get the enumerator.
            /// </summary>
            /// <returns>The enumerator.</returns>
            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetPrivateEnumerator();
            }

            /// <summary>
            /// Get the enumerable.
            /// </summary>
            /// <returns>The enumerable</returns>
            protected abstract IEnumerable<T> GetProtectedEnumerable();
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                           ExcelCellAddress                                                          |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================

        /// <summary>
        /// Class used to record Row / Column address.
        /// </summary>
        public class ExcelCellAddress
        {
            /// <summary>
            /// Row index, by default (-1) all rows.
            /// </summary>
            public long                 Row { get; set; }   = -1;

            /// <summary>
            /// Columns index, by default (-1) all columns.
            /// </summary>
            public long                 Column { get; set; }   = -1;

            /// <summary>
            /// Constructor with excel address;
            /// </summary>
            /// <param name="_strCellName">Cell address</param>
            public ExcelCellAddress(String _strCellName)
            {
                var regex = new Regex("(?<Column>[A-Z]*)(?<Row>[0-9]*)");
                var match = regex.Match(_strCellName);
                if (match.Success && match.Groups["Column"].Success && match.Groups["Row"].Success)
                {
                    Column = string.IsNullOrEmpty(match.Groups["Column"].Value)
                        ? -1
                        : NumberFromExcelColumn(match.Groups["Column"].Value);
                    try
                    {
                        Row = string.IsNullOrEmpty(match.Groups["Row"].Value)
                            ? -1
                            : long.Parse(match.Groups["Row"].Value);
                    }
                    catch(Exception _ex)
                    {
                        throw new ArgumentException(String.Format("Bad format for cell {0}: {1}", _strCellName, _ex.Message));
                    }
                }
                else
                {
                    throw new ArgumentException(String.Format("Bad format for cell {0}",_strCellName));
                }
            }

            /// <summary>
            /// Convert column to column letter.
            /// </summary>
            /// <param name="_lColumn">Column index.</param>
            /// <returns>The string that represent the column letter.</returns>
            private static string ExcelColumnFromNumber(long _lColumn)
            {
                string strColumnString = "";
                decimal dColumnNumber = _lColumn;
                while (dColumnNumber > 0)
                {
                    decimal dCurrentLetterNumber = (dColumnNumber - 1) % 26;
                    char cCurrentLetter = (char)(dCurrentLetterNumber + 65);
                    strColumnString = cCurrentLetter + strColumnString;
                    dColumnNumber = (dColumnNumber - (dCurrentLetterNumber + 1)) / 26;
                }

                return strColumnString;
            }

            /// <summary>
            /// Convert column letter to column index.
            /// </summary>
            /// <param name="_strColumn">The string that represent the column letter.</param>
            /// <returns>The index of the column.</returns>
            private static long NumberFromExcelColumn(string _strColumn)
            {
                long lRetVal = 0;
                string strColumn = _strColumn.ToUpper();
                for (int iChar = strColumn.Length - 1; iChar >= 0; iChar--)
                {
                    char cColPiece = strColumn[iChar];
                    long lColNum = cColPiece - 64;
                    lRetVal = lRetVal + lColNum * (int) Math.Pow(26, strColumn.Length - (iChar + 1));
                }
                return lRetVal;
            }
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                           ExcelCellRange                                                            |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================

        /// <summary>
        /// Class used to record cells range.
        /// </summary>
        public class ExcelCellRange
        {
            /// <summary>
            /// Range begin.
            /// </summary>
            public ExcelCellAddress    RangeFrom                            { get; set; }   = null;

            /// <summary>
            /// Range end.
            /// </summary>
            public ExcelCellAddress    RangeTo                              { get; set; }   = null;

            /// <summary>
            /// Accept all.
            /// </summary>
            public ExcelCellRange(string _strRange = null)
            {
                if (_strRange != null)
                {
                    int iIndex = _strRange.IndexOf(':');
                    if (iIndex == -1)
                    {   // Ok, RangeTo is null
                        RangeFrom = new ExcelCellAddress(_strRange);
                    }
                    else
                    {
                        RangeFrom = new ExcelCellAddress(_strRange.Substring(0, iIndex));
                        RangeTo   = new ExcelCellAddress(_strRange.Substring(iIndex + 1));
                    }
                }
            }

            /// <summary>
            /// Create a range with 2 cell addresses.
            /// </summary>
            /// <param name="_cellAddressFrom">Range's begin.</param>
            /// <param name="_cellAddressTo">Range's end.</param>
            public ExcelCellRange(ExcelCellAddress _cellAddressFrom, ExcelCellAddress _cellAddressTo)
            {
                RangeFrom = _cellAddressFrom;
                RangeTo   = _cellAddressTo;
            }

            /// <summary>
            /// Is the row is in the range ?
            /// </summary>
            /// <param name="_row">Row to test.</param>
            /// <returns>true if the row index (1 based) is into the range, false else.</returns>
            public bool IsInRange(DataRow _row)
            {
                bool bRet = true;

                if (RangeFrom != null)
                {   // If a range is defined
                    long rowID = Helper.GetRowID(_row);
                    bRet = RangeFrom.Row <= rowID && RangeTo.Row >= rowID;
                }

                return bRet;
            }

            /// <summary>
            /// Is column index (1 based index) is into the range.
            /// </summary>
            /// <param name="_iColumnIndex">Column index to test.</param>
            /// <returns>true if the column's index is into the range, false else.</returns>
            public bool IsColumnInRange(int _iColumnIndex)
            {
                return    (RangeFrom == null && RangeTo == null)
                       || (RangeTo   == null && RangeFrom.Column <= _iColumnIndex)
                       || (RangeFrom == null && RangeTo.Column >= _iColumnIndex)
                       || (RangeFrom.Column <= _iColumnIndex && RangeTo.Column >= _iColumnIndex);
            }

            /// <summary>
            /// Is row is the first one of the range.
            /// </summary>
            /// <param name="_lRowIndex">Row index to test (1 based).</param>
            /// <returns>true if the row's index is the first one into the (or the range is null and row inde is 1), false else.</returns>
            public bool IsFirstRowInRange(long _lRowIndex)
            {
                return (RangeFrom == null && _lRowIndex == 1) || (RangeFrom.Row == _lRowIndex);
            }

            /// <summary>
            /// Get the first row of range, 1 based index.
            /// </summary>
            /// <returns>The row index (1 based) of the first row of the range, 1 if RangeFrom is null.</returns>
            public long GetFirstRowOfRange()
            {
                return RangeFrom == null ? 1 : RangeFrom.Row;
            }

            /// <summary>
            /// Get the last row of range, 1 based index.
            /// </summary>
            /// <returns>The row index (1 based) of the last row of the range, long.MaxValue if RangeTo is null.</returns>
            public long GetLastRowOfRange()
            {
                return RangeTo == null ? long.MaxValue : RangeTo.Row;
            }

            /// <summary>
            /// Get the first column of range, 1 based index.
            /// </summary>
            /// <returns>The column index (1 based) of the first column of the range, 1 if RangeFrom is null.</returns>
            public long GetFirstColumnOfRange()
            {
                return RangeFrom == null ? 1 : RangeFrom.Column;
            }

            /// <summary>
            /// Get the last column of range, 1 based index.
            /// </summary>
            /// <returns>The column index (1 based) of the last column of the range, long.MaxValue if RangeTo is null.</returns>
            public long GetLastColumnOfRange()
            {
                return RangeTo == null ? long.MaxValue : RangeTo.Column;
            }
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                              NoRowHeader                                                            |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================

        /// <summary>
        /// Class of excel data cells when not binded behind a class mapping.
        /// </summary>
        public class NoRowHeader
        {
            /// <summary>
            /// Current column's index.
            /// </summary>
            public long ColumnIndex                 { get; private set; }

            /// <summary>
            /// Current value.
            /// </summary>
            public object Value                     { get; private set; }

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="_lColumnIndex">Column(s index.</param>
            /// <param name="_oValue">Value.</param>
            public NoRowHeader(long _lColumnIndex, object _oValue)
            {
                ColumnIndex  = _lColumnIndex;
                Value        = _oValue;
            }
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                             OneRowInfos                                                             |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// This class record one row .
        /// </summary>
        public class OneRowInfos<T> : IEnumerable<T>
        {
            /// <summary>
            /// Current rows.
            /// </summary>
            protected DataRow DataRow                   { get; set; }

            /// <summary>
            /// Current cells range.
            /// </summary>
            protected ExcelCellRange Range              { get; set; }

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="_dataRow">One row, cannot be null.</param>
            /// <param name="_range">Range, cannot be null.</param>
            public OneRowInfos(DataRow _dataRow, ExcelCellRange _range)
                //: base(_dataRows, _range)
            {
                if (_dataRow == null)
                {
                    throw new ArgumentNullException("RowsList: row is null!");
                }
                if (_range == null)
                {
                    throw new ArgumentNullException("RowsList: range is null!");
                }

                DataRow = _dataRow;
                Range   = _range;
            }

            /// <summary>
            /// Get the enumerator.
            /// </summary>
            /// <returns>The enumerator.</returns>
            IEnumerator<T> IEnumerable<T>.GetEnumerator()
            {
                return GetPrivateEnumerator();
            }

            /// <summary>
            /// Get the enumerator.
            /// </summary>
            /// <returns>The enumerator.</returns>
            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetPrivateEnumerator();
            }

            /// <summary>
            /// Create the enumerator.
            /// 
            /// Used by both function GetEnumerator.
            /// </summary>
            /// <returns>The enumerator.</returns>
            protected IEnumerator<T> GetPrivateEnumerator()
            {
                if (typeof(T) == typeof(NoRowHeader))
                {
                    int iIndex = 0;
                    var items = DataRow.ItemArray;

                    // Create a range each time to always call ++iIndex on each iteration
                    // If put "Range == null || Range.IsColumnInRange(++iIndex)" condition, the
                    // index pass to NoRowHeader will not be correct.
                    ExcelCellRange range = Range ?? new ExcelCellRange();

                    IEnumerable<NoRowHeader> rows = from row
                                                    in DataRow.ItemArray
                                                    where range.IsColumnInRange(++iIndex)
                                                    select new NoRowHeader(iIndex, row);

                    return (IEnumerator<T>) rows.GetEnumerator();
                }

                // Only for no header, else it has no sense.
                throw new NotImplementedException();
            }
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                         InnerConstantFinder                                                         |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// Finder class used to find constant expression.
        /// </summary>
        internal class InnerConstantFinder : ExpressionVisitor
        {
            /// <summary>
            /// Expression found.
            /// </summary>
            private ConstantExpression              m_innerConstantExpression = null;

            private Func<ConstantExpression, bool>  m_condition;

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="_expression">Expression to explore.</param>
            /// <returns>The constant expression if found, null else.</returns>
            public ConstantExpression GetInnerConstant(Expression _expression, Func<ConstantExpression, bool> _condition)
            {
                m_condition = _condition;
                Visit(_expression);
                return m_innerConstantExpression;
            }

            /// <summary>
            /// Call by the visitor on each constant expression.
            /// </summary>
            /// <param name="_expression">constant expression.</param>
            /// <returns>The constant expression if found, null else.</returns>
            protected override Expression VisitConstant(ConstantExpression _expression)
            {
                if (_expression.NodeType == ExpressionType.Constant && (m_condition == null || m_condition(_expression)))
                {
                    m_innerConstantExpression = _expression;
                }

                return _expression;
            }
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                        ExpressionTreeModifier                                                       |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// Used to modify the expression tree by replacing constant value by the query result.
        /// 
        /// Create an instance of this class with correct parameters and call Visit on this class to replace expression.
        /// </summary>
        internal class ExpressionTreeModifier : ExpressionVisitor
        {
            /// <summary>
            /// Result to replace
            /// </summary>
            private IQueryable          m_queryable;

            /// <summary>
            /// Type to find and replace into constant expression.
            /// </summary>
            private Type                m_typeToFind;

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="_queryable">Queryable result to replace into constant expression.</param>
            /// <param name="_typeToFind">Expression type to find.</param>
            internal ExpressionTreeModifier(IQueryable _queryable, Type _typeToFind)
            {
                m_queryable  = _queryable;
                m_typeToFind = _typeToFind;
            }

            /// <summary>
            /// Called by Visitor when it find a constant.
            /// </summary>
            /// <param name="_expression">Input expression.</param>
            /// <returns>The same expression as input if the type of expression is not the same as input parameter,
            ///          the replace expression else.</returns>
            protected override Expression VisitConstant(ConstantExpression _expression)
            {
                // Replace the constant of type m_typeToFind (different type results) arg with the queryable result.
                return _expression.Type == m_typeToFind
                    ? Expression.Constant(m_queryable)
                    : _expression;
            }
        }

        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                              FastProperty                                                           |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// Class used to assign value with the mapping.
        /// 
        /// Assigning a value with a PropertyInfo is very slow. This class make delegate to be able to assign a value
        /// faster than using PropertyInfo class.
        /// 
        /// Class found here : http://geekswithblogs.net/Madman/archive/2008/06/27/faster-reflection-using-expression-trees.aspx
        /// </summary>
        public class FastProperty
        {
            /// <summary>
            /// Original property info.
            /// </summary>
            public PropertyInfo             Property                        { get; set; }
 
            /// <summary>
            /// Get delegate to read data.
            /// </summary>
            private Func<object, object>     GetDelegate                    { get; set; }

            /// <summary>
            /// Set delegate to write data.
            /// </summary>
            private Action<object, object>   SetDelegate                    { get; set; }
 
            /// <summary>
            /// Create the fast property.
            /// </summary>
            /// <param name="_property">Original property to bind to.</param>
            public FastProperty(PropertyInfo _property)
            {
                Property = _property;
                InitializeGet();
                InitializeSet();
            }
 
            /// <summary>
            /// Initialize the Setter.
            /// </summary>
            private void InitializeSet()
            {
                try
                {
                    var instance = Expression.Parameter(typeof(object), "instance");
                    var value = Expression.Parameter(typeof(object), "value");
 
                    // value as T is slightly faster than (T)value, so if it's not a value type, use that
                    UnaryExpression instanceCast = (!this.Property.DeclaringType.IsValueType) ? Expression.TypeAs(instance, this.Property.DeclaringType) : Expression.Convert(instance, this.Property.DeclaringType);
                    UnaryExpression valueCast = (!this.Property.PropertyType.IsValueType) ? Expression.TypeAs(value, this.Property.PropertyType) : Expression.Convert(value, this.Property.PropertyType);
                    SetDelegate = Expression.Lambda<Action<object, object>>(Expression.Call(instanceCast, this.Property.GetSetMethod(), valueCast), new ParameterExpression[] { instance, value }).Compile();
                }
                catch
                {
                }

            }
 
            /// <summary>
            /// Initialize the Getter.
            /// </summary>
            private void InitializeGet()
            {
                try
                {
                    var instance = Expression.Parameter(typeof(object), "instance");
                    UnaryExpression instanceCast = (!this.Property.DeclaringType.IsValueType) ? Expression.TypeAs(instance, this.Property.DeclaringType) : Expression.Convert(instance, this.Property.DeclaringType);
                    GetDelegate = Expression.Lambda<Func<object, object>>(Expression.TypeAs(Expression.Call(instanceCast, this.Property.GetGetMethod()), typeof(object)), instance).Compile();
                }
                catch
                {
                }
            }
 
            /// <summary>
            /// Get the value.
            /// </summary>
            /// <param name="_oInstance">Object instance.</param>
            /// <returns>The value as object.</returns>
            public object Get(object _oInstance)
            {
                return GetDelegate(_oInstance);
            }
 
            /// <summary>
            /// Set the value.
            /// </summary>
            /// <param name="_oInstance">Object instance.</param>
            /// <param name="_value">Value to set.</param>
            /// <returns></returns>
            public void Set(object _oInstance, object _value)
            {
                SetDelegate(_oInstance, _value);
            }
        }
        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                               Helpers                                                               |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// To add the helper class that is used by the System.Linq.IQueryProvider implementation.
        /// </summary>
        internal static class Helper
        {
            /// <summary>
            /// Field Info for the DataRow.
            /// </summary>
            private static FieldInfo m_RowIDFieldInfo = null;

            /// <summary>
            /// Used to get the rowid of the DataRow
            /// </summary>
            private static FieldInfo RowIDFieldInfo
            {
                get
                {
                    if (m_RowIDFieldInfo == null)
                    {
                        m_RowIDFieldInfo = typeof(DataRow).GetField("_rowID",BindingFlags.NonPublic | BindingFlags.Instance);
                    }

                    return m_RowIDFieldInfo;
                }
            }

            /// <summary>
            /// Get the Row ID for DataRow class.
            /// </summary>
            /// <param name="_row">Row.</param>
            /// <returns>The row index (1 based)</returns>
            public static long GetRowID(DataRow _row)
            {
                return (long) RowIDFieldInfo.GetValue(_row);
            }

            /// <summary>
            /// Get the column's name.
            /// 
            /// A column name is a cell value but used into the mapping to know the column index when to
            /// get the data to make binding.
            /// </summary>
            /// <param name="_row"></param>
            /// <param name="_bReplaceCariageReturnByUnderscore"></param>
            /// <returns></returns>
            public static String GetColumnName(object _row, bool _bReplaceCariageReturnByUnderscore)
            {
                String strRet;
                try
                {
                    strRet = _row == null
                        ? String.Empty
                        : (_bReplaceCariageReturnByUnderscore
                            ? _row.ToString().Replace("\r\n","\n").Replace("\r","\n").Replace("\n","_")
                            :  _row.ToString());
                }
                catch
                {
                    strRet = String.Empty;
                }

                return strRet;
            }
        }


        // =======================================================================================================================================
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // |                                                              TypeSystem                                                             |
        // |                                                                                                                                     |
        // |                                                                                                                                     |
        // =======================================================================================================================================
        /// <summary>
        /// To add the helper class that is used by the System.Linq.IQueryProvider implementation.
        /// </summary>
        internal static class TypeSystem
        {
            internal static Type GetElementType(Type _seqType)
            {
                Type ienum = FindIEnumerable(_seqType);
                if (ienum == null)
                {
                    return _seqType;
                }
                return ienum.GetGenericArguments()[0];
            }

            private static Type FindIEnumerable(Type _seqType)
            {
                if (_seqType == null || _seqType == typeof(string))
                {
                    return null;
                }

                if (_seqType.IsArray)
                {
                    return typeof(IEnumerable<>).MakeGenericType(_seqType.GetElementType());
                }

                if (_seqType.IsGenericType)
                {
                    foreach (Type arg in _seqType.GetGenericArguments())
                    {
                        Type ienum = typeof(IEnumerable<>).MakeGenericType(arg);
                        if (ienum.IsAssignableFrom(_seqType))
                        {
                            return ienum;
                        }
                    }
                }

                Type[] ifaces = _seqType.GetInterfaces();
                if (ifaces != null && ifaces.Length > 0)
                {
                    foreach (Type iface in ifaces)
                    {
                        Type ienum = FindIEnumerable(iface);
                        if (ienum != null)
                        {
                            return ienum;
                        }
                    }
                }

                if (_seqType.BaseType != null && _seqType.BaseType != typeof(object))
                {
                    return FindIEnumerable(_seqType.BaseType);
                }

                return null;
            }
        }
    }
}
