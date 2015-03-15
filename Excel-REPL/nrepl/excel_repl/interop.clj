(ns excel-repl.interop)

(import ExcelDna.Integration.ExcelReference)
(import ClojureExcel.MainClass)

(require '[clojure.string :as str])
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
  (letter->val2 (re-find #"[A-Z]+" s)))

(defn row-num [s]
  (dec (int (re-find #"[0-9]+" s))))

(defn insert-value
  "Inserts val at ref.  Ref may be of the form A1 or SheetName!B6"
  [ref val]
  (let [
        [sheet ref] (if (.Contains ref "!") (str/split ref #"!") [nil ref])
        i (row-num ref)
        j (col-num ref)
        ref (if sheet (ExcelReference. i i j j sheet) (ExcelReference. i j))
        ]
    (.SetValue ref val)))

(defn insert-values
  "Inserts 2d array of values at ref.  Ref may of the form A1 or SheetName!B6"
  [ref values]
  (let [
        m (count values)
        n (count (first values))
        values (MainClass/RectangularArray (to-array-2d values))
        [sheet ref] (if (.Contains ref "!") (str/split ref #"!") [nil ref])
        i (row-num ref)
        j (col-num ref)
        id (+ i m -1)
        jd (+ j n -1)
        ref (if sheet (ExcelReference. i id j jd sheet) (ExcelReference. i id j jd))
        ]
    (-> ref (.SetValue values))))

(defn get-values
  "Returns values at ref which is of the form A1, SheetName!B6 or A1:B6.
   Single cell selections are returned as a value, 2D selections as an Object[][] array"
  [ref]
  (let [
        [sheet ref] (if (.Contains ref "!") (str/split ref #"!") [nil ref])
        refs (if (.Contains ref ":") (str/split ref #":") [ref ref])
        [i id] (map row-num refs)
        [j jd] (map col-num refs)
        ref (if sheet (ExcelReference. i id j jd sheet) (ExcelReference. i id j jd))
        ]
    (-> ref .GetValue MainClass/RaggedArray)))

