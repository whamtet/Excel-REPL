;no ns.  this will be evaluated in clojure.core for simplicity

(import System.Environment)
(import System.Text.RegularExpressions.Regex)
(import System.IO.Directory)
(import NetOffice.ExcelApi.Application)

(require '[clojure.repl :as r])
(require 'clojure.pprint)
(require 'clojure.walk)
(require '[clojure.string :as string])

(defn get-cd
  "returns current directory as a string"
  []
  (Directory/GetCurrentDirectory))

(defn set-cd
  "sets current directory as a string"
  [new-d]
  (Directory/SetCurrentDirectory new-d))

(defn get-load-path []
  (set (Regex/Split (Environment/GetEnvironmentVariable "CLOJURE_LOAD_PATH") ";")))

(defn set-load-path [s]
  (let [
        new-path (apply str (interpose ";" s))
        ]
    (Environment/SetEnvironmentVariable "CLOJURE_LOAD_PATH" new-path)
    new-path))

(defn append-load-path
  "appends file string to clojure load path"
  [new-path]
  (set-load-path (conj (get-load-path) new-path)))

(defmacro with-out-strs
  "evaluates expression and returns list of lines printed"
  [x]
  `(System.Text.RegularExpressions.Regex/Split (with-out-str ~x) "\n"))

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

;;process output
(def letters "ABCDEFGHIJKLMNOPQRSTUVWXYZ")
(def letter->val (util/reduce-map (map-indexed (fn [i s] [s i]) letters)))

(defn letter->val2 [[s t :as ss]]
  (if t
    (apply + 26
           (map *
                (map letter->val (reverse ss))
                (map #(Math/Pow 26 %) (range))))
    (letter->val s)))

(defn letter->val3 [s]
  (letter->val2 (re-find #"[A-Z]+" s)))

(defn selection-width [select-str]
  (inc (apply - (map letter->val3 (string/split select-str ":")))))

(defn transpose [arr]
  (let [n (count (first arr))]
    (for [j (range n)]
      (map #(nth % j) arr))))

(defn partition-range [values select-str]
  (let [
        selection-width (selection-width select-str)
        ]
    (if (= 1 selection-width)
      values
      (let [
            m (partition selection-width values)
            ]
        (if (.EndsWith select-str "'")
          (transpose m)
          m)))))

(defn range-values
  "returns row-wise seq of values selected by select-str"
  [app select-str]
  (let [
        select-str2 (.Replace select-str "'" "")
        range (.Range app select-str2)
        ]
    (if (.Contains select-str ":")
      (mapv #(.Value %) range)
      (.Value range))))

(defn excel-reference? [ref]
  (if (symbol? ref)
    (re-find #"[A-Z]+[0-9]+" (str ref))))

(defmacro with-excel-refs
  "Expands excel references of the form A1 A2:B6 or Sheet3!A3 etc.
  References must be symbols.  External references not supported."
  [x]
  (let [
        app (Application/GetActiveInstance)
        f #(if (excel-reference? %)
             (range-values app (str %))
             %)
        ]
    (clojure.walk/prewalk f x)))
