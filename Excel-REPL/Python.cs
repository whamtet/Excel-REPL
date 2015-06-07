using Microsoft.Scripting.Hosting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace clojureexcel
{
    public class Python
    {
        private static ScriptEngine engine;
        private static ScriptScope scope;

        static Python()
        {
        }
        private static void Init()
        {
            engine = IronPython.Hosting.Python.CreateEngine();
            scope = engine.CreateScope();
            var paths = engine.GetSearchPaths();
            paths.Add("C:\\Program Files (x86)\\Excel-REPL\\Excel-REPL\\Lib");
            paths.Add("C:\\Anaconda\\Lib");
            paths.Add("C:\\Program Files (x86)\\Excel-REPL\\Excel-REPL\\python");
            paths.Add("C:\\Users\\xuehuit\\Documents\\Visual Studio 2012\\Projects\\Excel-REPL\\Excel-REPL\\python");
            engine.SetSearchPaths(paths);

            String initCode = @"
from pygments import highlight
from pygments.lexers import ClojureLexer
from pygments.formatters import HtmlFormatter

clojureLexer = ClojureLexer()
htmlFormatter = HtmlFormatter()

def raw(code):
  return highlight(code, clojureLexer, htmlFormatter)";

            engine.Execute(initCode, scope);
        }


        public static dynamic Eval(String input)
        {
            return engine.Execute(input, scope);
        }

        public static String Html(String code)
        {
            code = String.Format("raw(\"\"\"{0}\"\"\")", code);
            return Eval(code);
        }

        public static String Css()
        {
            return Eval("htmlFormatter.get_style_defs()");
        }
    }
}