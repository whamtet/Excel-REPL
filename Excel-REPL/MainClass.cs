using clojure.lang;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Reflection;
using Ionic.Zip;
using System.Collections.Concurrent;
using System.Windows.Forms;
using System.Threading;

namespace ClojureExcel
{
    public class MainClass : ExcelDna.Integration.CustomUI.ExcelRibbon, IExcelAddIn
    {
        public void AutoClose() { }
        public void AutoOpen()
        {
            Init();
            ExcelIntegration.RegisterUnhandledExceptionHandler(
                ex => "!!! EXCEPTION: " + ex.ToString());
        }
        
        public static void InsertNewWorksheet(String name)
        {
            NetOffice.ExcelApi.Application app = NetOffice.ExcelApi.Application.GetActiveInstance();
            var book = app.ActiveWorkbook;
            var sheets = book.Worksheets;

        }

        private static Object GetFirst(Object o)
        {
            return ((Object[,])o)[0, 0];
        }

        public static IFn export_udfs, invoke_anonymous_macros;

        public static void ExportUdfs()
        {
            try
            {
                export_udfs.invoke();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        [ExcelCommand]
        public static object InvokeAnonymousMacros()
        {
            try
            {
                return invoke_anonymous_macros.invoke();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return null;
            }
        }

        public static Object AssemblyPaths()
        {
            var assemblies = AppDomain.CurrentDomain
                            .GetAssemblies()
                            .Where(a => !a.IsDynamic)
                            .Select(a => a.Location);
            return assemblies;
        }
        //referenced by excel-repl.udf
        public static void RegisterMethods(MethodInfo[] methods)
        {
            List<MethodInfo> l = new List<MethodInfo>();
            foreach (MethodInfo info in methods)
            {
                l.Add(info);
            }
            Integration.RegisterMethods(l);
        }

//        [ExcelCommand(MenuName = "f", MenuText = "f")]
//        public static void f()
//        {
//            var app = ExcelDnaUtil.Application as Microsoft.Office.Interop.Excel.Application;
//            app.Selection.Insert();
//        }

        
        private static void Init()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();

                var resourceName = "Excel_REPL.nrepl.zip";
                Stream stream = assembly.GetManifestResourceStream(resourceName);
                ZipFile f = ZipFile.Read(stream);

                string tempPath = System.IO.Path.GetTempPath();
                string tempFolder = tempPath + "Excel_REPL\\";
                f.ExtractAll(tempFolder, ExtractExistingFileAction.OverwriteSilently);
                stream.Close();
                tempFolder += "nrepl\\";
                appendLoadPath(tempFolder);

                String clojureSrc = ResourceSlurp("excel-repl.clj");
                msg = (String)GetFirst(my_eval(clojureSrc, "clojure.core"));
                //try loading main within main
                String main_load = @"
(require '[excel-repl.interop :as interop])
(interop/require-sheet '[main A])
(main/main)
";
            }
            catch (Exception e)
            {
                msg = e.ToString();
            }
        }
        public static String ResourceSlurp(String resource)
        {
            return (String)slurp.invoke(Assembly.GetExecutingAssembly().GetManifestResourceStream("Excel_REPL." + resource));
        }
        //used by drawbridge-client
        public static BlockingCollection<Object> GetCollection()
        {
            return new BlockingCollection<Object>();
        }
        public static String appendLoadPath(String newPath)
        {
            String loadPath = Environment.GetEnvironmentVariable("CLOJURE_LOAD_PATH");
            if (loadPath == null)
            {
                loadPath = newPath;
            }
            else
            {
                loadPath += ";" + newPath;
            }
            Environment.SetEnvironmentVariable("CLOJURE_LOAD_PATH", loadPath);
            return loadPath;
        }
        private static IFn load_string = clojure.clr.api.Clojure.var("clojure.core", "load-string");
 //       private static IFn is_nil = clojure.clr.api.Clojure.var("clojure.core", "nil?");
        private static IFn slurp = clojure.clr.api.Clojure.var("clojure.core", "slurp");
//        private static IFn number = clojure.clr.api.Clojure.var("clojure.core", "number?");
        public static Dictionary<String, String> d = new Dictionary<String, String>();
        private static string msg = "nothing";


        public static String GetMsg()
        {
            return msg;
        }
        
        public static Object my_eval(String input)
        {
            return my_eval(input, getSheetName());
        }
        public static Object my_eval(String input, String sheetName)
        {
            Object o;
            try
            {
                if (!input.StartsWith("(ns"))
                {
                    input = String.Format("(ns {0})\n", sheetName) + input;
                }
                o = load_string.invoke(input);
            }
            catch (Exception e)
            {
                return pack(Regex.Split(e.ToString(), "\n"));
            }
            return process_output(o);
        }

        private static Object cleanValue(object o)
        {
            if (o == null)
            {
                return "";
            }
            if (o is bool)
            {
                return o;
            }
            if (o is Ratio)
            {
                return ((Ratio)o).ToDouble(null);
            }
            if (o is sbyte
            || o is byte
            || o is short
            || o is ushort
            || o is int
            || o is uint
            || o is long
            || o is ulong
            || o is float
            || o is double
            || o is decimal)
            {
                return o;
            }
            else
            {
                return o.ToString();
            }
        }

        private static Object process_output(Object o)
        {
            return process_output(o, 1);
        }
        private static Object process_output(Object o, int level)
        {
            if (o is IPersistentCollection)
            {
                o = ((IPersistentCollection)o).seq();
            }
            if (o is ISeq)
            {
                ISeq r2 = (ISeq)o;
                object[] outArr = new object[r2.count()];
                int i = 0;
                while (r2 != null)
                {
                    outArr[i] = level == 1 ? process_output(r2.first(), 0) : cleanValue(r2.first());
                    i++;
                    r2 = r2.next();
                }
                return level == 1 ? pack(outArr) : outArr;
            }
            else
            {
                return level == 1 ? pack(cleanValue(o)) : cleanValue(o);
            }
        }
        private static object pack(object o)
        {
            if (o is object[])
            {
                object[] o2 = (object[])o;
                int m = o2.Length;
                m = m < 2 ? 2 : m;
                int n = 2;
                foreach (object o3 in o2)
                {
                    if (o3 is object[])
                    {
                        object[] o4 = (object[])o3;
                        if (o4.Length > n)
                        {
                            n = o4.Length;
                        }
                    }
                }
                var oot = new object[m, n];
                for (int i = 0; i < m; i++)
                {
                    if (i >= o2.Length)
                    {
                        for (int j = 0; j < n; j++)
                        {
                            oot[i, j] = "";
                        }
                    }
                    else
                    {
                        object o3 = o2[i];
                        if (o3 is object[])
                        {
                            object[] o4 = (object[])o3;
                            for (int j = 0; j < o4.Length; j++)
                            {
                                oot[i, j] = o4[j];
                            }
                            for (int j = o4.Length; j < n; j++)
                            {
                                oot[i, j] = "";
                            }
                        }
                        else
                        {
                            oot[i, 0] = o3;
                            for (int j = 1; j < n; j++)
                            {
                                oot[i, j] = "";
                            }
                        }
                    }
                }
                return oot;
            }
            else
            {
                var oot = new object[2, 2];
                oot[0, 0] = o;
                oot[0, 1] = "";
                oot[1, 0] = "";
                oot[1, 1] = "";
                return oot;
            }
        }
        
        public static String GetVersion()
        {
            return "0.0.1";
        }
        public static Object RaggedArray(Object arrayCandidate)
        {
            var input = arrayCandidate as Object[,];
            if (input == null)
            {
                return arrayCandidate;
            }
            int m = input.GetUpperBound(0) + 1;
            int n = input.GetUpperBound(1) + 1;
            Object[][] output = new Object[m][];
            for (int i = 0; i < m; i++ )
            {
                Object[] row = new Object[n];
                for (int j = 0; j < n; j++)
                {
                    row[j] = input[i, j];
                }
                output[i] = row;
            }
            return output;
        }

        public static Object[,] RectangularArray(Object[] input)
        {
            int m = input.Length;
            Object[] firstRow = (Object[])input[0];
            int n = firstRow.Length;
            var output = new Object[m, n];
            for (var i = 0; i < m; i++)
            {
                Object[] row = (Object[])input[i];
                for (var j = 0; j < n; j++)
                {
                    output[i, j] = row[j];
                }
            }

            return output;
        }
        
        public static String getSheetName()
        {
            ExcelReference reference = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            string sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
            sheetName = Regex.Split(sheetName, "\\]")[1];
            sheetName = sheetName.Replace(" ", "");
            return sheetName;
        }

        //referenced by clojure.data.drawbridge-client
        public static Object TakeItem(BlockingCollection<Object> c)
        {
            Object outObj;
            c.TryTake(out outObj, 0);
            return outObj;
        }
        public static object Load(Object[] name)
        {

            StringBuilder input = new StringBuilder();
            foreach (Object s in name)
            {
                if (s.GetType() != typeof(ExcelEmpty))
                {
                    input.Append(s + "\n");
                }
            }

            return my_eval(input.ToString());

        }
    }
}
