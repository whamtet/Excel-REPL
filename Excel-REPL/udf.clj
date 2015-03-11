(import System.CodeDom.Compiler.CompilerParameters)
(import Microsoft.CSharp.CSharpCodeProvider)
(import System.Reflection.BindingFlags)
(import ClojureExcel.MainClass)

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

(defn comma-interpose [s] (apply str (interpose ", " s)))
(defn line-interpose [s] (apply str (interpose "\r\n" s)))

(defn dirty-arglist? [l]
  (some #(or (= '& %) (map? %)) l))

(defn filter-arglists [v]
  (let [
        {:keys [name arglists doc export]} (meta v)
        arglists (remove dirty-arglist? arglists)
;        clean-str2 #(if (vector? %) % (clean-str %))
        ]
    (if (and (not-empty arglists) export) [(clean-str name) arglists doc (var-get v)])))

(defn filter-ns-interns [ns]
  (filter identity (map filter-arglists (vals (ns-interns ns)))))

(defn append-replace [s a b]
  (.Replace s a (str a b)))

(defn emit-static-method [method-name name arglist doc]
  (let [

        doc (format "[ExcelFunction(Description=@\"%s\")]" (or doc ""))
        arg-types (map #(if (vector? %) "Object[] " "Object ") arglist)
        clean-args (map #(if (vector? %) (gensym) (clean-str %)) arglist)
        arglist1 (comma-interpose (map str arg-types clean-args))
        arglist2 (comma-interpose clean-args)
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
    (line-interpose (map #(emit-static-method %1 name %2 doc) method-names arglists))))

(defn class-str [d]
  (let [

        fns (map first d)
        fn-str (comma-interpose fns)
        construct-fn-str (comma-interpose (map #(str "IFn " %) fns))
        construct-body-str (apply str (map #(format "        Class1.%s = %s;\r\n" % %) fns))
        static-methods (line-interpose (map emit-static-methods d))

        s (MainClass/ResourceSlurp "Class1.cs")
        s (.Replace s "ifn_list" fn-str)
        s (.Replace s "IFn foo" construct-fn-str)
        s (append-replace s "Class1.MainClassInstance = MainClass;\r\n" construct-body-str)
        s (.Replace s "    private void Poo() { }" static-methods)
        ]
    s))

(defn get-methods [t]
  (.GetMethods t (enum-or BindingFlags/Public BindingFlags/Static)))

(defn compile-ns [ns]
  (let [
        d (filter-ns-interns ns)
        t (-> d class-str my-compile .CompiledAssembly .GetTypes first)
        types (map last d)
        mci (MainClass.)
        constructor-args (object-array (conj types mci))
        ]
    (Activator/CreateInstance t constructor-args)
    (MainClass/RegisterMethods (get-methods t))))
