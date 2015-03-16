(ns excel-repl.schedule-udf)

;;ok, two strategies
;;for ns we will concatenate a list of namespaces to compile
;;ExportUDFS will invoke a fixed function that compiles them

;;for in-macro-context we shall concatenate a list of fuctions to execute
;;INVOKEMacroContext shall execute them all and return the last result


;;need to define a macro that defines the properties that we want

(def nss (ref []))
(def fns (ref []))

(defn add-ns [ns]
  (dosync
   (alter nss conj ns)))

(defn add-curr-ns []
  (add-ns *ns*))

(defn add-fn [fn]
  (dosync
   (alter fns conj fn)))

(defn get-ns []
  (dosync
   (let [a @nss]
     (ref-set nss [])
     a)))

(defn get-fns []
  (dosync
   (let [a @fns]
     (ref-set fns [])
     a)))
