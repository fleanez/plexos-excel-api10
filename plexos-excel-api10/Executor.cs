using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PLEXOS_NET.Core;
using EnergyExemplar.PLEXOS.Energy;
using Path = System.IO.Path;


namespace plexos_excel_api_10
{
    internal class Executor
    {

        private static readonly String DEFAULT_SUFFIXE = "(MOD)";

        public static bool AddObject(DatabaseCore db, PObject obj)
        {
            int ret = 0;
            bool bExists = false;
            int classID = GetClassId(db, obj.ClassName);
            string[] objectlist = db.GetObjects(classID);
            if (objectlist?.Length > 0)
            {
                bExists = Array.IndexOf(objectlist, obj.Name) >= 0;
            }
            if (!bExists)
            {
                ret = db.AddObject(obj.Name, (int)GetClassEnum(obj.ClassName), true);
            }
            return (ret == 1);
        }

        public static bool AddMembership (DatabaseCore db, PMember member)
        {
            EEUTILITY.Enums.CollectionEnum collection = GetCollectionEnum(member);
            string[] children = db.GetChildMembers((int)collection, member.Parent.Name);
            if (children?.Length>0)
            {
                if (collection == EEUTILITY.Enums.CollectionEnum.ModelScenarios)
                {
                    db.RemoveMembership(((int)collection), member.Parent.Name, children[0]);
                }
                else
                {
                    int nRem = 0;
                    foreach (var s in children)
                    {
                        nRem += db.RemoveMembership((int)collection, member.Parent.Name, s);
                    }
                }
            }

            int ret = db.AddMembership((int)collection, member.Parent.Name, member.Child.Name);
            return (ret == 1);
        }

        public static int GetMembershipID(DatabaseCore db, PMember member)
        {
            EEUTILITY.Enums.CollectionEnum collection = GetCollectionEnum(member);
            return db.GetMembershipID((int)collection, member.Parent.Name, member.Child.Name);
        }

        public static bool SetAttributeValue (DatabaseCore db, PAttribute attribute)
        {
            int nAttributeId = GetAttributeEnum(db, attribute.ClassName, attribute.Name);
            return db.SetAttributeValue((int)GetClassEnum(attribute.ClassName), attribute.ChildName, nAttributeId, attribute.Value);
        }

        private static EEUTILITY.Enums.ClassEnum GetClassEnum (String strClassName)
        {
            switch(strClassName)
            {
                case "Generator":
                    return EEUTILITY.Enums.ClassEnum.Generator;
                case "Model":
                    return EEUTILITY.Enums.ClassEnum.Model;
                case "Scenario":
                    return EEUTILITY.Enums.ClassEnum.Scenario;
                case "Horizon":
                    return EEUTILITY.Enums.ClassEnum.Horizon;
                case "Stochastic":
                    return EEUTILITY.Enums.ClassEnum.Stochastic;
                case "Performance":
                    return EEUTILITY.Enums.ClassEnum.Performance;
                case "System":
                    return EEUTILITY.Enums.ClassEnum.System;
                case "Transmission":
                    return EEUTILITY.Enums.ClassEnum.Transmission;
                case "MTSchedule":
                    return EEUTILITY.Enums.ClassEnum.MTSchedule;
                case "STSchedule":
                    return EEUTILITY.Enums.ClassEnum.STSchedule;
                case "Report":
                    return EEUTILITY.Enums.ClassEnum.Report;
                case "Diagnostic":
                    return EEUTILITY.Enums.ClassEnum.Diagnostic;
                case "Production":
                    return EEUTILITY.Enums.ClassEnum.Production;
                default:
                    throw new Exception("Unsupported class " + strClassName);
            }
        }

        public static bool IsSupportedClass (String strClassName)
        {
            try
            {
                EEUTILITY.Enums.ClassEnum classenum = GetClassEnum(strClassName);
                return true;
            } catch { return false; }
        }

        public static EEUTILITY.Enums.CollectionEnum GetCollectionEnum(PMember member)
        {
            return GetCollectionEnum(member.Parent.ClassName, member.Child.ClassName, member.Collection);
        }

        public static EEUTILITY.Enums.CollectionEnum GetCollectionEnum(String strParentClass, String strChildClass, String strCollectionName)
        {
            if (strParentClass == "System" && strChildClass=="Model" && strCollectionName== "Models")
            {
                return EEUTILITY.Enums.CollectionEnum.SystemModels;
            } 
            else if (strParentClass == "Model" && strChildClass == "Scenario" && strCollectionName == "Scenarios")
            {
                return EEUTILITY.Enums.CollectionEnum.ModelScenarios;
            }
            else if (strParentClass == "Model" && strChildClass == "Horizon" && strCollectionName == "Horizon")
            {
                return EEUTILITY.Enums.CollectionEnum.ModelHorizon;
            }
            else if (strParentClass == "Model" && strChildClass == "Report" && strCollectionName == "Report")
            {
                return EEUTILITY.Enums.CollectionEnum.ModelReport;
            }
            else if (strParentClass == "Model" && strChildClass == "MT Schedule" && strCollectionName == "MT Schedule")
            {
                return EEUTILITY.Enums.CollectionEnum.ModelMTSchedule;
            }
            else if (strParentClass == "Model" && strChildClass == "Stochastic" && strCollectionName == "Stochastic")
            {
                return EEUTILITY.Enums.CollectionEnum.ModelStochastic;
            }
            else if (strParentClass == "Model" && strChildClass == "Transmission" && strCollectionName == "Transmission")
            {
                return EEUTILITY.Enums.CollectionEnum.ModelTransmission;
            }
            else if (strParentClass == "Model" && strChildClass == "Production" && strCollectionName == "Production")
            {
                return EEUTILITY.Enums.CollectionEnum.ModelProduction;
            }
            else if (strParentClass == "Model" && strChildClass == "Performance" && strCollectionName == "Performance")
            {
                return EEUTILITY.Enums.CollectionEnum.ModelPerformance;
            }
            else if (strParentClass == "Model" && strChildClass == "Diagnostic" && strCollectionName == "Diagnostic")
            {
                return EEUTILITY.Enums.CollectionEnum.ModelDiagnostic;
            }
            else if (strParentClass == "Model" && strChildClass == "LT Plan" && strCollectionName == "LT Plan")
            {
                return EEUTILITY.Enums.CollectionEnum.ModelLTPlan;
            }
            throw new Exception("Unsupported collection " + strCollectionName);
        }

        public static bool IsSupportedCollection(String strParentClass, String strChildClass, String strCollectionName) 
        {
            try
            {
                EEUTILITY.Enums.CollectionEnum collectenum = GetCollectionEnum(strParentClass, strChildClass, strCollectionName);
                return true;
            }
            catch { return false; }
        }

        private static int GetAttributeEnum(DatabaseCore db, String strClassName, String strAttributeName)
        {
            //String[] sClassFields = new String[] { "class_id", "name", "enum_id" };
            String[] sClassFields = new String[] { }; //Workaround! this seems to work the opposite! it returns all fields except the selected!
            int nClassId = GetClassId(db, strClassName);
            ADODB.Recordset rec = db.GetData("t_attribute", ref sClassFields);

            while (!rec.EOF)
            {
                int nId = (int)rec.Fields["class_id"].Value;
                var strName = rec.Fields["name"].Value.ToString();
                if (nId == nClassId && strName == strAttributeName)
                {
                    int nRet = (int)rec.Fields["enum_id"].Value;
                    rec.Close();
                    return nRet;
                }
                rec.MoveNext();
            }
            rec.Close();
            throw new Exception("Couldn't find the id for class " + strClassName);
        }

        private static Dictionary<String, int> ClassIds =[];
        private static int GetClassId(DatabaseCore db, String strClassName)
        {
            if (ClassIds.Count == 0)
            {
                InitClassId(db);
            }
            return ClassIds[strClassName];
        }

        private static void InitClassId(DatabaseCore db)
        {
            //String[] sClassFields = new String[] {"class_id", "name"};
            String[] sClassFields = []; //Workaround! this seems to work the opposite! it returns all fields except the selected!
            ADODB.Recordset rec = db.GetData("t_class", ref sClassFields);
            while (!rec.EOF)
            {
                if (rec.Fields["name"] != null && rec.Fields["class_id"] != null)
                {
                    var strName = rec.Fields["name"].Value.ToString();
                    int nRet = (int)rec.Fields["class_id"].Value;
                    ClassIds.Add(strName, nRet);
                }
                rec.MoveNext();
            }
            rec.Close();
        }

        public static String CreateBackUpFile(String strFile, String strSuffix)
        {
            String strInputFileBackUp = Path.GetDirectoryName(strFile) + Path.DirectorySeparatorChar + Path.GetFileNameWithoutExtension(strFile) + strSuffix + Path.GetExtension(strFile);
            File.Copy(strFile, strInputFileBackUp, true);
            Console.WriteLine("Backing up file " + strFile);
            return strInputFileBackUp;
        }

        private static void PrintHelp()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("Plexos-Excel-api.exe [CONFIG_EXCEL_FILE] [CONFIG_EXCEL_SHEET] [PLEXOS_XML_FILE] [SUFFIXE]");
            Console.WriteLine("   [CONFIG_EXCEL_FILE]     Path to Excel config file");
            Console.WriteLine("   [CONFIG_EXCEL_SHEET]    Excel's sheet name (case sensitive)");
            Console.WriteLine("   [PLEXOS_XML_FILE]       PLEXOS input database (xml) to be modified");
            Console.WriteLine($"   [SUFFIXE]               (optional) Suffixe to be appended to the resulting modified PLEXOS input database. (Default={DEFAULT_SUFFIXE})");
        }

        static void Main(string[] args)
        {
            String strConfigFile;
            String strSheetName;
            String strPLEXOSFile;
            String strSuffix;
            if (args.Length < 3)
            {
                Console.WriteLine("Parameters are not optional");
                PrintHelp();
                Console.ReadKey();
                System.Environment.Exit(1);
            }
            strConfigFile = args[0];
            strSheetName = args[1];
            strPLEXOSFile = args[2];
            strSuffix = DEFAULT_SUFFIXE;
            if (args.Length >= 4)
            {
                strSuffix = args[3];
            }
            //Read config:
            Reader r = new Reader(strConfigFile, strSheetName);

            //Create Connection:
            string strInputFile = CreateBackUpFile(strPLEXOSFile, strSuffix);
            DatabaseCore db = new();
            db.Connection(strInputFile);

            //Write Modifications:
            int nAddObj = 0;
            foreach (PObject obj in r.GetObjects())
            {
                if (AddObject(db, obj))
                {
                    nAddObj++;
                }
            }
            Console.WriteLine($"Writen {nAddObj} objects to database. Failed {r.GetObjects().Count - nAddObj}");
            int nAddMembers = 0;
            foreach (PMember mem in r.GetMembers())
            {
                if (AddMembership(db, mem))
                {
                    nAddMembers++;
                }
            }
            Console.WriteLine($"Writen {nAddMembers} memberships to database. Failed {r.GetMembers().Count - nAddMembers}");
            int nAddAttribute = 0;
            foreach (PAttribute att in r.GetAttributes())
            {
                if (SetAttributeValue(db, att))
                {
                    nAddAttribute++;
                }
            }
            Console.WriteLine($"Writen {nAddAttribute} attributes to database. Skipped {r.GetAttributes().Count - nAddAttribute}");
            Console.WriteLine($"Saving changes to file {strInputFile}");
            db.Close();
            Console.WriteLine($"Press any key to finish...");
            Console.ReadKey();
        }
    }
}
