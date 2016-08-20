(ns excel-repl.interop)

(import ExcelDna.Integration.ExcelReference)
(import ExcelDna.Integration.XlCall)
(import ClojureExcel.MainClass)

(assembly-load "ExcelApi")
(import NetOffice.ExcelApi.Application)

(require '[clojure.string :as str])

(defn comma-interpose [s] (apply str (interpose ", " s)))
(defn line-interpose [s] (apply str (interpose "\r\n" s)))


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

;;DEPRECATED!!
;;WRITING TO THE WORKBOOK RANDOMLY CRASHES IT.
#_(defn insert-value
  "Inserts val at ref."
  [sheet ref val]
  (let [
        i (row-num ref)
        j (col-num ref)
        ref (ExcelReference. i i j j sheet)
        ]
    (.SetValue ref val)))

(defn split-str [s]
  (map #(str "\"" (apply str %) "\"") (partition-all 250 s)))

(defn concatenated-str [s]
  (format "CONCATENATE(%s)" (comma-interpose (split-str s))))

(defn excel-pr-str [s]
  (if (string? s) (concatenated-str (.Replace s "\"" "\"\"")) s))

(defn formula-str
  "Generates Excel formula string =f(arg1, \"arg2\"...).
  Long strings a split via =CONCATENATE to conform to the Excel 255 character limit"
  [f & args]
  (format "=%s(%s)" f (comma-interpose (map excel-pr-str args))))


(defn regularize-array
  "ensures array is rectangular"
  [arr]
  (let [
        n (apply max (map count arr))
        extend #(take n (concat % (repeat nil)))
        ]
    (map extend arr)))


;;DEPRECATED!!
;;WRITING TO THE WORKBOOK RANDOMLY CRASHES IT.
#_(defn insert-values
  "Inserts 2d array of values at ref."
  [sheet ref values]
  (let [
        values (regularize-array values)
        m (count values)
        n (count (first values))
        values (-> values to-array-2d MainClass/RectangularArray)
        i (row-num ref)
        j (col-num ref)
        id (+ i m -1)
        jd (+ j n -1)
        ref (ExcelReference. i id j jd sheet)
        ]
    (-> ref (.SetValue values))))

(defn get-values
  "Returns values at ref which is of the form A1 or A1:B6.
  Single cell selections are returned as a value, 2D selections as an Object[][] array"
  [sheet ref]
  (let [
        refs (if (.Contains ref ":") (str/split ref #":") [ref ref])
        [i id] (map row-num refs)
        [j jd] (map col-num refs)
        value (.GetValue (ExcelReference. i id j jd sheet))
        ]
    (if (.Contains ref ":")
      (MainClass/RaggedArray value)
      value)))

;;DEPRECATED!!
;;WRITING TO THE WORKBOOK RANDOMLY CRASHES IT.
#_(defn insert-formula
  "Takes a single formula and inserts it into one or many cells.
  Use this instead of insert-values when you have a formula.
  Because Excel-REPL abuses threads the formulas may be stale when first inserted.
  "
  [sheet ref formula]
  (let [
        refs (if (.Contains ref ":") (str/split ref #":") [ref ref])
        [i id] (map row-num refs)
        [j jd] (map col-num refs)
        ref (ExcelReference. i id j jd sheet)
        ]
    (XlCall/Excel XlCall/xlcFormulaFill (object-array [formula ref]))))

;;DEPRECATED!!
;;WRITING TO THE WORKBOOK RANDOMLY CRASHES IT.
#_(defn add-sheet
  "Adds new sheet to current workbook."
  [name]
  (let [
        sheets (-> (Application/GetActiveInstance) .ActiveWorkbook .Worksheets)
        existing-names (set (map #(.Name %) sheets))
        name (if (existing-names name)
               (loop [i 1]
                 (let [new-name (format "%s (%s)" name i)]
                   (if (existing-names new-name)
                     (recur (inc i))
                     new-name))) name)
        sheet (.Add sheets)
        ]
    (set! (.Name sheet) name)))

(defn require-sheet [v]
  "Require excel spreadsheet.  V is of form
  [sheet A C D]
  [sheet A C D :as alias]
  A C D are the columns containing source code"
  (let [
        sheet-name (first v)
        [_as alias-name] (take-last 2 v)
        [alias-name cols]
        (if (= :as _as)
          [alias-name (drop 1 (drop-last 2 v))]
          [sheet-name (drop 1 v)])
        ns-aliases (-> *ns* ns-aliases keys set)
        ]
    (if-not (ns-aliases alias-name)
      (let [
            source
            (apply str
                   (flatten
                    (for [col cols]
                      (line-interpose
                       (filter string?
                               (map first
                                    (get-values (str sheet-name) (format "%s1:%s200" col col))))))))
            ]
        (MainClass/my_eval source (str sheet-name))
        (require (vector sheet-name :as alias-name))))))

;;DEPRECATED!!
;;WRITING TO THE WORKBOOK RANDOMLY CRASHES IT.
#_(defmacro clear-contents
  "Clears an m by n grid at sheet, ref.
  Must be called inside udf/in-macro-context"
  [sheet ref m n]
  `(interop/insert-values ~sheet ~ref
                          (let [
                                row# (repeat ~n nil)
                                rows# (repeat ~m row#)
                                ]
                            rows#)))

