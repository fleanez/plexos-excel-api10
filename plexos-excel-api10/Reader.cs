using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace plexos_excel_api_10
{
    internal class Reader
    {

        private String m_FileName;
        private Excel.Worksheet m_Configsheet;
        private static readonly int PARENT_CLASS_COLUMN = 1;
        private static readonly int CHILD_CLASS_COLUMN = 2;
        private static readonly int PARENT_NAME_COLUMN = 3;
        private static readonly int CHILD_NAME_COLUMN = 4;
        private static readonly int TYPE_COLUMN = 5;
        private static readonly int NAME_COLUMN = 6;
        private static readonly int VALUE_COLUMN = 7;
        private static readonly string OBJECT_KEYWORD = "object";
        private static readonly string MEMBER_KEYWORD = "collection";
        private static readonly string ATTRIBUTE_KEYWORD = "attribute";

        private List<PObject> m_objects = [];
        private List<PMember> m_members = [];
        private List<PAttribute> m_attributes = [];

        public Reader(String strFileName, string strConfigSheet)
        {
            this.m_FileName = strFileName;
            var xcel = new Excel.Application();
            
            var book = xcel.Workbooks.Open(strFileName,false,true);
            
            m_Configsheet = book.Worksheets[strConfigSheet];

            Excel.Range r = m_Configsheet.UsedRange;
            Console.WriteLine("Reading config file..");
            for (int i = 1; i < r.Rows.Count; i++)
            {

                string strType = r.Cells[i + 1, TYPE_COLUMN].Value;
                if (OBJECT_KEYWORD.Equals(strType,StringComparison.OrdinalIgnoreCase))
                {
                    string strName = r.Cells[i + 1, CHILD_NAME_COLUMN].Value;
                    string strClass = r.Cells[i + 1, CHILD_CLASS_COLUMN].Value;
                    if (string.IsNullOrEmpty(strName) || string.IsNullOrEmpty(strClass))
                    {
                        throw new Exception($"ERROR IN ROW {i + 1}: Child class (Column { CHILD_CLASS_COLUMN }) and child name (Column {CHILD_NAME_COLUMN}) are mandatory! Can't create empty class or object name");
                    }
                    PObject obj = new PObject(strClass, strName);
                    m_objects.Add(obj);
                } else if (MEMBER_KEYWORD.Equals(strType, StringComparison.OrdinalIgnoreCase))
                {
                    string strParentClass = r.Cells[i + 1, PARENT_CLASS_COLUMN].Value;
                    string strChildClass = r.Cells[i + 1, CHILD_CLASS_COLUMN].Value;
                    string strParentName = Convert.ToString(r.Cells[i + 1, PARENT_NAME_COLUMN].Value);
                    string strChildName = Convert.ToString(r.Cells[i + 1, CHILD_NAME_COLUMN].Value);
                    string strCollectionName = r.Cells[i + 1, NAME_COLUMN].Value;
                    if (string.IsNullOrEmpty(strParentClass) || string.IsNullOrEmpty(strChildClass) || string.IsNullOrEmpty(strParentName) || string.IsNullOrEmpty(strChildName) || string.IsNullOrEmpty(strCollectionName))
                    {
                        throw new Exception($"ERROR IN ROW {i + 1}: Columns { PARENT_CLASS_COLUMN } - {NAME_COLUMN} are mandatory for Memberships!");
                    }
                    PObject parent = new PObject(strParentClass, strParentName);
                    PObject child =  new PObject(strChildClass, strChildName);
                    PMember mem = new PMember(strCollectionName, parent, child);
                    m_members.Add(mem);
                }
                else if (ATTRIBUTE_KEYWORD.Equals(strType, StringComparison.OrdinalIgnoreCase))
                {
                    string strName = r.Cells[i + 1, NAME_COLUMN].Value;
                    string strClass = r.Cells[i + 1, CHILD_CLASS_COLUMN].Value;
                    string strChildName = r.Cells[i + 1, CHILD_NAME_COLUMN].Value;
                    double dValue = Convert.ToDouble(r.Cells[i + 1, VALUE_COLUMN].Value);
                    if (string.IsNullOrEmpty(strName) || string.IsNullOrEmpty(strClass) || string.IsNullOrEmpty(strChildName))
                    {
                        throw new Exception($"ERROR IN ROW {i + 1}: Column { CHILD_CLASS_COLUMN }, {CHILD_NAME_COLUMN}, {NAME_COLUMN} and {VALUE_COLUMN} are mandatory for Attributes!");
                    }
                    PAttribute attr = new PAttribute(strClass, strChildName, strName, dValue);
                    m_attributes.Add(attr);
                }
                else
                {
                    throw new Exception($"ERROR IN ROW {i+1}: Unsupported parameter type '{strType}' (Column {TYPE_COLUMN})");
                }
            }
            book.Close();

        }

        public List<PObject> GetObjects() {
            return m_objects;
        }

        public List<PMember> GetMembers()
        {
            return m_members;
        }
        public List<PAttribute> GetAttributes()
        {
            return m_attributes;
        }

    }
}
