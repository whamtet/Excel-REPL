(ns excel-repl.formatting)

(import ClojureExcel.MainClass)
(import System.Windows.Forms.MessageBox)

(defn f [x]
  ;(future
    (MyRibbon/SetOutput (map #(.ToUpper %) x)))
  ;)

(set! MainClass/format_code f)
