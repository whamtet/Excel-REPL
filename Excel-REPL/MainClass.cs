using clojure.lang;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace ClojureExcel
{
    public static class MainClass
    {
        private static IFn load_string = clojure.clr.api.Clojure.var("clojure.core", "load-string");
        private static string msg;

        private static Object doublize(object o)
        {
            if (o is Ratio)
            {
                return ((Ratio)o).ToDouble(null);
            }
            if (o is Var || o.GetType().ToString() == "System.RuntimeType" || o is IFn)
            {
                return o.ToString();
            }
            if (o == null)
            {
                return "nil";
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
        
        
        private static Object my_eval(String input)
        {
            //return pack(Regex.Split(input, "\n"));
            Object o;
            try
            {
                input = String.Format("(ns {0})\n", getSheetName()) + input;
                o = load_string.invoke(input);
            }
            catch (Exception e)
            {
                return pack(Regex.Split(e.ToString(), "\n"));
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
        [ExcelFunction(Description = "Evaluate in Octant")]
        public static string Require(string s)
        {
            return my_eval(String.Format("(require '{0})", s)).ToString();
        }
        [ExcelFunction(Description = "Evaluate in Octant")]
        public static Object Invoke1(String f, Object[] a0)
        {
            String s = "(" + f + " " + stringifies2(a0) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "Evaluate in Octant")]
        public static Object Invoke2(String f, Object[] a0, Object[] a1)
        {
            String s = "(" + f + " " + stringifies2(a0, a1) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "Evaluate in Octant")]
        public static Object Invoke3(String f, Object[] a0, Object[] a1, Object[] a2)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "Evaluate in Octant")]
        public static Object Invoke4(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "Evaluate in Octant")]
        public static Object Invoke5(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "Evaluate in Octant")]
        public static Object Invoke6(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "Evaluate in Octant")]
        public static Object Invoke7(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5, Object[] a6)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5, a6) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "Evaluate in Octant")]
        public static Object Invoke8(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5, Object[] a6, Object[] a7)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5, a6, a7) + ")";
            return my_eval(s);
        }
        [ExcelFunction(Description = "Evaluate in Octant")]
        public static Object Invoke9(String f, Object[] a0, Object[] a1, Object[] a2, Object[] a3, Object[] a4, Object[] a5, Object[] a6, Object[] a7, Object[] a8)
        {
            String s = "(" + f + " " + stringifies2(a0, a1, a2, a3, a4, a5, a6, a7, a8) + ")";
            return my_eval(s);
        }
    }
}
