using clojure.lang;
using ExcelDna.Integration;
using System;
using System.Windows.Forms;

public class Class1
{
    private static IFn ifn_list;
    private static dynamic MainClassInstance;

    public Class1(dynamic MainClass, IFn foo)
    {
        Class1.MainClassInstance = MainClass;
    }

    private void Poo() { }



}
