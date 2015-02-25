;no ns.  this will be evaluated in clojure.core for simplicity

(import System.Environment)
(import System.Text.RegularExpressions.Regex)
(import System.IO.Directory)
(require '[clojure.repl :as r])
(require 'clojure.pprint)

(defn get-cd []
    (Directory/GetCurrentDirectory))
(defn set-cd [new-d]
    (Directory/SetCurrentDirectory new-d))

(defn get-load-path []
    (set (Regex/Split (Environment/GetEnvironmentVariable "CLOJURE_LOAD_PATH") ";")))

(defn set-load-path [s]
    (let [
        new-path (apply str (interpose ";" s))
        ]
    (Environment/SetEnvironmentVariable "CLOJURE_LOAD_PATH" new-path)
        new-path))

(defn append-load-path [new-path]
    (set-load-path (conj (get-load-path) new-path)))

(defmacro with-out-strs [x]
    `(System.Text.RegularExpressions.Regex/Split (with-out-str ~x) "\n"))

(defmacro source [x]
    `(with-out-strs (r/source ~x)))

(defmacro doc [x]
    `(with-out-strs (r/doc ~x)))

(defmacro pprint [x]
    `(with-out-strs (clojure.pprint/pprint ~x)))

(defmacro time-str [x]
    `(with-out-strs (time ~x)))