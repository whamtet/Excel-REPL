(ns excel-repl.util)
(import System.Windows.Forms.Clipboard)

(defn comma-interpose [s] (apply str (interpose ", " s)))
(defn line-interpose [s] (apply str (interpose "\r\n" s)))

(defn to-clipboard [s]
  (Clipboard/SetText s))
