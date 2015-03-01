#Excel REPL

![WFT](WFT.png)
```clojure
;This is a macro
(defmacro kick-some-ass [how-much?]
  `(do ~@(repeat how-much? '(kick-ass!))))
```

As much as we all love VBA there are better languages available such as ... Clojure.  Excel REPL makes it easy to start a ClojureCLR Repl from within Excel.  Simply install it as an Excel Add-In to provide a few additional Excel Functions

##Load

    =Load(A:A)

Concatenates the contents of selected cells and evaluates them in namespace SheetName.  SheetName is the name of the current worksheet.

##Invoke

    =Invoke1("myFunction", A1)
    =Invoke2("myFunction", A2:A5, C6)
    ...

Invokes myFunction with the supplied arguments and returns the result.  Range arguments are provided as a Clojure vector unless the range is of length one, in which case they are the same as a single cell argument.

Arity cannot be overloaded so you must use Invoke1, Invoke2 etc.  For functions with zero arguments simply use Load.

##RInvoke and RLoad

    =RInvoke1("myFunction", A1)
    =RInvoke2("myFunction", A2:A5, C6)
    =RLoad("(+ 1 2)")
    ...

The same as Invoke and Load but evaluated on a remote nrepl instance.  To define the remote address first invoke

    =DefineClient("nrepl://host:port")

Nrepl connections are also supported over http with the help of [Drawbridge](https://github.com/cemerick/drawbridge).  Add the drawbridge server handler from this project and Excel REPL will take care of the client side.

    =DefineClient("http://some.server/drawbridge-nrepl-handler")


##Returning 1D and 2D arrays

If your function returns a 1 or 2 dimensional collection you may paste it into a range of Excel Cells.  To do so

1) Drag from the top left hand corner the number of cells for your output

2) Click in the formula bar and enter your formula

3) Press Control + Shift + Enter instead of simply enter

##Error Messages

Errors are caught and returned as text within the output cells.  The stacktrace is split down the column so select multiple cells for output as mentioned above.

##Auxilliary Methods

Excel REPL adds useful functions and macros to clojure.core that are useful when interacting with a worksheet.  Please see [excel-repl.clj](https://github.com/whamtet/Excel-REPL/blob/master/Excel-REPL/excel-repl.clj) for details.

If you wish to pull stuff off the net straight into your worksheet [clr-http-lite](https://github.com/whamtet/clr-http-lite) is included

```clojure

(require '[clr-http.lite.client :as client])

(client/get "http://google.com")
=> {:status 200
    :headers {"date" "Sun, 01 Aug 2010 07:03:49 GMT"
              "cache-control" "private, max-age=0"
              "content-type" "text/html; charset=ISO-8859-1"
              ...}
    :body "<!doctype html>..."}

```

##Build and Installation

The build process is a bit of a manual hack.  Please contact the author if you want help with this.

To install on a machine copy the contents of Excel-REPL/bin/Debug into ~/AppData/Roaming/Microsoft/AddIns and then add via the Excel add-ins menu.  Excel-REPL will then auto-load every time you start Excel.

##System Requirements

Excel Repl works with Microsoft Excel 97+ (that's quite old) and Microsoft .NET 4.0 or 4.5.

##Good Riddance VBA

What a crap language.  Lisp existed then, so why did Microsoft create VBA?  Because they're Microsoft.
