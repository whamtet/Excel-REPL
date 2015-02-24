using clojure.lang;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Reflection;
using Ionic.Zip;
using NetOffice.ExcelApi;
using System.Collections.Concurrent;
using System.Web;
using System.IO.Compression;

namespace ClojureExcel
{
    public static class MainClass
    {

        static MainClass()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();

                var resourceName = "Excel_REPL.nrepl.zip";
                Stream stream = assembly.GetManifestResourceStream(resourceName);
                ZipFile f = ZipFile.Read(stream);

                string tempPath = System.IO.Path.GetTempPath();
                string tempFolder = tempPath + "\\Excel_REPL";
                f.ExtractAll(tempFolder, ExtractExistingFileAction.OverwriteSilently);
                stream.Close();
                appendLoadPath(tempFolder);
                string clojureSrc = @"
(import System.Environment)
(import System.Text.RegularExpressions.Regex)
(import System.IO.Directory)
(require '[clojure.repl :as r])
(require 'clojure.pprint)

(defn get-cd []
    (Directory/GetCurrentDirectory))
(defn set-cd [new-d]
    (Directory/SetCurrentDirectory new-d))

(defn get-load-path []
    (set (Regex/Split (Environment/GetEnvironmentVariable ""CLOJURE_LOAD_PATH"") "";"")))

(defn set-load-path [s]
    (let [
        new-path (apply str (interpose "";"" s))
        ]
    (Environment/SetEnvironmentVariable ""CLOJURE_LOAD_PATH"" new-path)
        new-path))

(defn append-load-path [new-path]
    (set-load-path (conj (get-load-path) new-path)))

(defmacro with-out-strs [x]
    `(System.Text.RegularExpressions.Regex/Split (with-out-str ~x) ""\n""))

(defmacro source [x]
    `(with-out-strs (r/source ~x)))

(defmacro doc [x]
    `(with-out-strs (r/doc ~x)))

(defmacro pprint [x]
    `(with-out-strs (clojure.pprint/pprint ~x)))

(defmacro time-str [x]
    `(with-out-strs (time ~x)))
";
                Object[,] o = (Object[,]) my_eval(clojureSrc, "clojure.core");
                
                msg = (String)o[0, 0];
            }
            catch (Exception e)
            {
                msg = e.ToString();
            }

        }
        public static BlockingCollection<Object> getCollection()
        {
            return new BlockingCollection<Object>();
        }
        public static HttpUtility getUtility()
        {
            return new HttpUtility();
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
        private static IFn is_nil = clojure.clr.api.Clojure.var("clojure.core", "nil?");
        private static string msg;

        //this section is dummy stuff to make sure things load correctly
        public static Application a = Application.GetActiveInstance();

        private static Object doublize(object o)
        {
            if ((bool)is_nil.invoke(o))
            {
                return "nil";
            }
            if (o is Ratio)
            {
                return ((Ratio)o).ToDouble(null);
            }
            if (o is Var || o.GetType().ToString() == "System.RuntimeType" || o is IFn)
            {
                return o.ToString();
            }
            else
            {
                return o;
            }
        }

        private static object seqize(Object o)
        {
            if (o is IPersistentCollection) {
                o = ((IPersistentCollection)o).seq();
            }
            if (o is ISeq)
            {
                ISeq r2 = (ISeq)o;
                object[] outArr = new object[r2.count()];
                int i = 0;
                while (r2 != null)
                {
                    outArr[i] = doublize(r2.first());
                    i++;
                    r2 = r2.next();
                }
                return outArr;
            }
            else
            {
                return doublize(o);
            }
        }
        [ExcelFunction]
        public static String GetMsg()
        {
            return msg;
        }
        [ExcelFunction(Description = "define client")]
        public static Object DefineClient(Object connect_info)
        {
            string input = @"
(require '[clojure.tools.nrepl :as nrepl])

(def conn (nrepl/url-connect ""{0}""))
(def client (nrepl/client conn 10000))

(defn remote-eval [code]
(->
client
(nrepl/message (hash-map :op ""eval"" :code code))
nrepl/response-values))

(defmacro eval2 [& body]
`(first (remote-eval (nrepl/code ~@body))))
";
            try
            {
                String connect_str;
                if (connect_info is String)
                {
                  connect_str = (String) connect_info;
                }
                else
                {
                  connect_str = String.Format("nrepl://localhost:{0}", connect_info);
                }
                my_eval(String.Format(input, connect_str), "client");
                return connect_str;
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        private static Object remote_eval(String input)
        {
            input = String.Format("(eval2 {0})", input);
            return my_eval(input, "client");
        }

        private static Object my_eval(String input)
        {
            return my_eval(input, getSheetName());
        }
        private static Object my_eval(String input, String sheetName)
        {
            Object o;
            try
            {
                input = String.Format("(ns {0})\n", sheetName) + input;
                o = load_string.invoke(input);
            }
            catch (Exception e)
            {
                return pack(Regex.Split(e.ToString(), "\n"));
            }
            return process_output(o);
        }

        private static Object process_output(Object o)
        {
            if ((bool) (is_nil.invoke(o)))
            {
                return pack("nil");
            }
            
            if (o is IPersistentCollection) {
                o = ((IPersistentCollection)o).seq();
            }
            if (o is ISeq)
            {
                ISeq r2 = (ISeq)o;
                object[] outArr = new object[r2.count()];
                int i = 0;
                while (r2 != null) {
                    outArr[i] = seqize(r2.first());
                    i++;
                    r2 = r2.next();
                }
                return pack(outArr);
            }
            else
            {
                return pack(doublize(o));
            }
        }

        private static String stringify(object o)
        {
            String s = o.GetType().Name;
            switch (s)
            {
                case "String":
                    return "\"" + o + "\"";
                case "ExcelEmpty":
                    return "nil";
                case "ExcelMissing":
                    return "nil";
                case "ExcelError":
                    return "nil";
                default:
                    return o.ToString();
            }
        }
        private static String stringifies(object[] o)
        {
            if (o.Length == 1)
            {
                return stringify(o[0]) + " ";
            }
            else
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("[");
                foreach (object o2 in o)
                {
                    sb.Append(stringify(o2) + " ");
                }
                sb.Append("] ");
                return sb.ToString();
            }
        }
        private static String stringifies2(params object[][] o)
        {
            StringBuilder sb = new StringBuilder();
            foreach (object[] o2 in o)
            {
                sb.Append(stringifies(o2));
            }
            return sb.ToString();
        }

        [ExcelFunction(Description = "Get Version")]
        public static String GetVersion()
        {
            return "0.0.1";
        }
        private static object pack(object o) {
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
                            oot[i, j] = "nil";
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
                                oot[i, j] = "nil";
                            }
                        }
                        else
                        {
                            oot[i, 0] = o3;
                            for (int j = 1; j < n; j++)
                            {
                                oot[i, j] = "nil";
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
                oot[0, 1] = "nil";
                oot[1, 0] = "nil";
                oot[1, 1] = "nil";
                return oot;
            }
        }
        private static String getSheetName()
        {
            ExcelReference reference = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            string sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
            sheetName = Regex.Split(sheetName, "\\]")[1];
            sheetName = sheetName.Replace(" ", "");
            return sheetName;
        }
        [ExcelFunction(Description="hihi")]
        public static Object Test()
        {
            try
            {
                ExcelExamplesCS45.Example01 e = new ExcelExamplesCS45.Example01();
                e.RunExample();
                return "done";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }


        [ExcelFunction(Description = "My first .NET function")]
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
        [ExcelFunction(Description = "")]
        public static Object Invoke1(String f, Object[] a0)
        {
            String s = "(" + f + " " + stringifies2(a0) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object Invoke2(String f, Object[] a0, Object[] a1)
        {
            String s = "(" + f + " " + stringifies2(a0, a1) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object Invoke3(String f, Object[] a0, Object[] a1, Object[] a2)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object Invoke4(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object Invoke5(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object Invoke6(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object Invoke7(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5, Object[] a6)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5, a6) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object Invoke8(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5, Object[] a6, Object[] a7)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5, a6, a7) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object Invoke9(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5, Object[] a6, Object[] a7, Object[] a8)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5, a6, a7, a8) + ")";
            return my_eval(s);
        }

        [ExcelFunction(Description = "")]
        public static Object RInvoke1(String f, Object[] a0)
        {
            String s = "(" + f + " " + stringifies2(a0) + ")";
            return remote_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object RInvoke2(String f, Object[] a0, Object[] a1)
        {
            String s = "(" + f + " " + stringifies2(a0, a1) + ")";
            return remote_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object RInvoke3(String f, Object[] a0, Object[] a1, Object[] a2)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2) + ")";
            return remote_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object RInvoke4(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3) + ")";
            return remote_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object RInvoke5(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4) + ")";
            return remote_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object RInvoke6(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5) + ")";
            return remote_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object RInvoke7(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5, Object[] a6)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5, a6) + ")";
            return remote_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object RInvoke8(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5, Object[] a6, Object[] a7)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5, a6, a7) + ")";
            return remote_eval(s);
        }
        [ExcelFunction(Description = "")]
        public static Object RInvoke9(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5, Object[] a6, Object[] a7, Object[] a8)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5, a6, a7, a8) + ")";
            return remote_eval(s);
        }
    }
}
