(ns excel-repl.udf)

(import System.CodeDom.Compiler.CompilerParameters)
(import Microsoft.CSharp.CSharpCodeProvider)
(import System.Reflection.BindingFlags)
(import ClojureExcel.MainClass)
(import System.Windows.Forms.MessageBox)

(assembly-load "ExcelApi")
(import NetOffice.ExcelApi.Application)

(require '[excel-repl.schedule-udf :as schedule-udf])
(require '[excel-repl.util :as util])

(def loaded-classes (MainClass/AssemblyPaths))
(defn load-path [s] (some #(if (.Contains % s) %) loaded-classes))

(defn my-compile [code]
  (let [
        cp (CompilerParameters.)
        ]
    (set! (.GenerateExecutable cp) false)
    (set! (.GenerateInMemory cp) true)

    (-> cp .ReferencedAssemblies (.Add (load-path "System.Windows.Forms.dll")))
    (-> cp .ReferencedAssemblies (.Add (load-path "Clojure.dll")))
    (-> cp .ReferencedAssemblies (.Add (load-path "ExcelDna.Integration.dll")))
    (-> cp .ReferencedAssemblies (.Add (load-path "System.Core.dll")))
    (.CompileAssemblyFromSource
     (CSharpCodeProvider.)
     cp (into-array [code]))))

(def to-clean {"?" "_QMARK_" "!" "_BANG_" ">" "_GT_" "<" "_LT_" "-" "_"})

(defn clean-str [s]
  (reduce (fn [s [old new]] (.Replace s old new)) (.ToLower (str s)) to-clean))


(defn dirty-arglist? [l]
  (some #(or (= '& %) (map? %)) l))

(defn filter-arglists [v]
  (let [
        {:keys [name arglists doc export]} (meta v)
        arglists (remove dirty-arglist? arglists)
        ]
    (if (and (not-empty arglists) export) [(clean-str name) arglists doc (var-get v)])))

(defn filter-ns-interns [ns]
  (filter identity (map filter-arglists (vals (ns-interns ns)))))

(defn filter-all-interns []
  (mapcat filter-ns-interns (schedule-udf/get-ns)))

(defn append-replace [s a b]
  (.Replace s a (str a b)))

(defn emit-static-method [method-name name arglist doc]
  (let [
        doc (format "[ExcelFunction(Description=@\"%s\")]" (or doc ""))
        arg-types (map #(if (vector? %) "Object[] " "Object ") arglist)
        clean-args (map #(if (vector? %) (gensym) (clean-str %)) arglist)
        arglist1 (util/comma-interpose (map str arg-types clean-args))
        arglist2 (util/comma-interpose clean-args)
        ]
    (format "%s
            public static object %s(%s)
            {
            //MessageBox.Show(\"invoking\");
            try { return %s.invoke(%s); } catch (Exception e) {return e.ToString();}
            }" doc method-name arglist1 name arglist2)))

(defn emit-static-methods [[name arglists doc]]
  (let [
        method-names (if (= 1 (count arglists))
                       [(.ToUpper name)]
                       (map #(str (.ToUpper name) (count %)) arglists))]
    (util/line-interpose (map #(emit-static-method %1 name %2 doc) method-names arglists))))

(defn class-str [d]
  (let [
        fns (map first d)
        fn-str (util/comma-interpose fns)
        construct-fn-str (util/comma-interpose (map #(str "IFn " %) fns))
        construct-body-str (apply str (map #(format "        Class1.%s = %s;\r\n" % %) fns))
        static-methods (util/line-interpose (map emit-static-methods d))

        s (MainClass/ResourceSlurp "Class1.cs")
        s (.Replace s "ifn_list" fn-str)
        s (.Replace s "IFn foo" construct-fn-str)
        s (append-replace s "Class1.MainClassInstance = MainClass;\r\n" construct-body-str)
        s (.Replace s "    private void Poo() { }" static-methods)
        ]
    s))

(defn get-methods [t]
  (.GetMethods t (enum-or BindingFlags/Public BindingFlags/Static)))

(defn export-udfs []
  (let [d (filter-all-interns)]
    (if (not-empty d)
      (let [
            t (-> d class-str my-compile .CompiledAssembly .GetTypes first)
            types (map last d)
            mci (MainClass.)
            constructor-args (object-array (conj types mci))
            ]
        (Activator/CreateInstance t constructor-args)
        (MainClass/RegisterMethods (get-methods t))))))

(defn export-fns []
  (schedule-udf/add-curr-ns)
  (.Run (Application/GetActiveInstance) "ExportUdfs"))

(defmacro in-macro-context
  "Evaluates body within an Excel macro context so that cell values can be set without throwing an exception."
  [& body]
  `(do
     (schedule-udf/add-fn
      (fn [] ~@body))
     (.Run (NetOffice.ExcelApi.Application/GetActiveInstance) "InvokeAnonymousMacros")))

(defn invoke-anonymous-macros []
  (or (last (map #(%) (schedule-udf/get-fns))) "Result Empty"))

(set! MainClass/export_udfs export-udfs)
(set! MainClass/invoke_anonymous_macros invoke-anonymous-macros)

(defn split-words [n s]
  (loop [
         todo s
         sb (StringBuilder.)
         i 0
         done []]
    (if-let [c (first todo)]
      (do
        (.Append sb c)
        (if (and (> i n) (= \space c))
          (recur (rest s) (StringBuilder.) 0 (conj done (str sb)))
          (recur (rest s) sb (inc i) done)))
      (let [
            last-line (str sb)
            ]
        (if (= "" last-line)
          done
          (conj done last-line))))))


