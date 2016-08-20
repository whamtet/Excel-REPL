;no ns.  this will be evaluated in clojure.core for simplicity

(import System.Environment)
(import System.IO.Directory)
(import System.Windows.Forms.MessageBox)

(import ExcelDna.Integration.ExcelReference)
(import ExcelDna.Integration.XlCall)
(import ClojureExcel.MainClass)

;(import NetOffice.ExcelApi.Application)

(require '[clojure.repl :as r])
(require 'clojure.pprint)
(require '[clojure.string :as string])
(require 'clojure.walk)

(defn show
  "Show MessageBox"
  [x]
  (MessageBox/Show x))

(defn get-cd
  "returns current directory as a string"
  []
  (Directory/GetCurrentDirectory))

(defn set-cd
  "sets current directory as a string"
  [new-d]
  (Directory/SetCurrentDirectory new-d))

(defn get-load-path []
  (set (string/split (Environment/GetEnvironmentVariable "CLOJURE_LOAD_PATH") #";")))

(defn set-load-path! [s]
  (let [
        new-path (apply str (interpose ";" s))
        ]
    (Environment/SetEnvironmentVariable "CLOJURE_LOAD_PATH" new-path)
    new-path))

(defn append-load-path!
  "appends file string to clojure load path"
  [new-path]
  (set-load-path! (conj (get-load-path) new-path)))

(defn split-lines [s]
  (string/split s #"\n"))

(defmacro with-out-strs
  "evaluates expression and returns list of lines printed"
  [x]
  `(split-lines (with-out-str ~x)))

(defmacro source
  "function source returned as string"
  [x]
  `(with-out-strs (r/source ~x)))

(defmacro doc
  "function docstring"
  [x]
  `(with-out-strs (r/doc ~x)))

(defmacro pprint
  "pprint to string"
  [x]
  `(with-out-strs (clojure.pprint/pprint ~x)))

(defmacro time-str
  "times evaluation of expression x"
  [x]
  `(with-out-strs (time ~x)))


(def letters "ABCDEFGHIJKLMNOPQRSTUVWXYZ")
(def letter->val (into {} (map-indexed (fn [i s] [s i]) letters)))

(defn letter->val2
  "column number of excel coumn A, AZ etc"
  [[s t :as ss]]
  (if t
    (apply + 26
           (map *
                (map letter->val (reverse ss))
                (map #(Math/Pow 26 %) (range))))
    (letter->val s)))

(defn col-num
  "column number of reference in form A4 etc"
  [s]
  (if (string? s)
    (letter->val2 (re-find #"[A-Z]+" s))
    (second s)))

(defn row-num
  [s]
  (if (string? s)
    (dec (int (re-find #"[0-9]+" s)))
    (first s)))

(defn get-values
  "Returns values at ref which is of the form A1 or A1:B6.
  Single cell selections are returned as a value, 2D selections as an Object[][] array"
  [sheet ref]
  (let [
        refs (if (.Contains ref ":") (string/split ref #":") [ref ref])
        [i id] (map row-num refs)
        [j jd] (map col-num refs)
        value (.GetValue (ExcelReference. i id j jd sheet))
        ]
    (if (.Contains ref ":")
      (MainClass/RaggedArray value)
      value)))

;;otha stuff
