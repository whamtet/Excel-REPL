;no ns.  this will be evaluated in clojure.core for simplicity

(import System.Environment)
(import System.IO.Directory)
;(import NetOffice.ExcelApi.Application)

(require '[clojure.repl :as r])
(require 'clojure.pprint)
(require '[clojure.string :as string])
(require 'clojure.walk)

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

;;otha stuff

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

(defn letter->val3
  "column number of reference in form A4 etc"
  [s]
  (letter->val2 (re-find #"[A-Z]+" s)))

(defn selection-width
  "selection width of reference in form A1:B2"
  [select-str]
  (let [
        [a b] (string/split select-str #"!")
        select-str (or b a)
        ]
    (inc (Math/Abs (apply - (map letter->val3 (string/split select-str #":")))))))

(defn transpose
  "transpose 2d array"
  [arr]
  (let [n (count (first arr))]
    (vec
     (for [j (range n)]
       (mapv #(nth % j) arr)))))

(defn partition-range
  "processes values into 2d array if necessary"
  [values select-str]
  (let [
        selection-width (selection-width select-str)
        ]
    (if (= 1 selection-width)
      (vec values)
      (let [
            m (partition selection-width values)
            ]
        (if (.EndsWith select-str "'")
          (transpose m)
          (mapv vec m))))))

(defn range-values
  "Returns array of values selected by select-str."
  [app select-str]
  (let [
        select-str2 (.Replace select-str "'" "")
        range (.Range app select-str2)
        ]
    (if (.Contains select-str ":")
      (partition-range (map #(.Value %) range) select-str)
      ;(mapv #(.Value %) range)
      (.Value range))))

(defn excel-reference? [ref]
  (if (symbol? ref)
    (re-find #"[A-Z]+[0-9]+" (str ref))))

#_(defmacro with-excel-refs
  "Expands excel references of the form A1 A2:B6 or Sheet3!A3 etc.
  References must be symbols.  External references are not supported.

  Two dimensional references are returned as row wise arrays.  To transpose
  append with a dash e.g. A2:B4'
  "
  [x]
  (let [
        app (Application/GetActiveInstance)
        f #(if (excel-reference? %)
             (range-values app (str %))
             %)
        ]
    (clojure.walk/prewalk f x)))
