(ns excel-repl.util)

(defn comma-interpose [s] (apply str (interpose ", " s)))
(defn line-interpose [s] (apply str (interpose "\r\n" s)))
