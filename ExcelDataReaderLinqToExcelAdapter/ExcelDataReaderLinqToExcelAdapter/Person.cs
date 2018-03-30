using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReaderLinqToExcelAdapter
{
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
        public String Comment       { set; get;}
    }
}
