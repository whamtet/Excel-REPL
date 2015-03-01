;; Copyright (c) Stuart Sierra, 2012. All rights reserved.  The use
;; and distribution terms for this software are covered by the Eclipse
;; Public License 1.0 (http://opensource.org/licenses/eclipse-1.0.php)
;; which can be found in the file epl-v10.html at the root of this
;; distribution.  By using this software in any fashion, you are
;; agreeing to be bound by the terms of this license.  You must not
;; remove this notice, or any other, from this software.

;;  Modified to run under ClojureCLR by David Miller
;;  Changes are
;; Copyright (c) David Miller, 2013. All rights reserved.  The use
;; and distribution terms for this software are covered by the Eclipse
;; Public License 1.0 (http://opensource.org/licenses/eclipse-1.0.php)
;; which can be found in the file epl-v10.html at the root of this
;; distribution.  By using this software in any fashion, you are
;; agreeing to be bound by the terms of this license.  You must not
;; remove this notice, or any other, from this software.

;; Changes to Stuart Sierra's code are clearly marked.
;; An end-of-line comment start with  ;DM:  indicates a change on that line.
;; The commented material indicates what was in the original code or indicates a new line inserted.
;; For example, a comment such as the following
;;    (defn- read-array [^PushbackTextReader stream]  ;DM: ^PushbackReader
;; indicates that the type hint needed to be replaced
;; More substantial changes are commented more substantially.


(ns ^{:author "Stuart Sierra, modifed for ClojureCLR by David Miller"
      :doc "JavaScript Object Notation (JSON) parser/generator.
  See http://www.json.org/"}
  clojure.data.json
  (:refer-clojure :exclude (read))
  (:require [clojure.pprint :as pprint])
  ;DM: (:import (java.io PrintWriter PushbackReader StringWriter
  ;DM:                   StringReader Reader EOFException)))
  (:import (System.IO EndOfStreamException StreamWriter StringReader         ;DM: Added
                      StringWriter TextWriter)                    ;DM: Added
           (clojure.lang PushbackTextReader)                                 ;DM: Added
		   (System.Globalization NumberStyles StringInfo)                    ;DM: Added
		   ))                                                                ;DM: Added
 
(set! *warn-on-reflection* true)

;;; JSON READER

(def ^:dynamic ^:private *bigdec*)
(def ^:dynamic ^:private *key-fn*)
(def ^:dynamic ^:private *value-fn*)

(defn- default-write-key-fn
  [x]
  (cond (instance? clojure.lang.Named x)
        (name x)
        (nil? x)
        (throw (Exception. "JSON object properties may not be nil"))
        :else (str x)))

(defn- default-value-fn [k v] v)

(declare -read)

(defmacro ^:private codepoint [c]
  (int c))

(defn- codepoint-clause [[test result]]
  (cond (list? test)
        [(map int test) result]
        (= test :whitespace)
        ['(9 10 13 32) result]
        (= test :simple-ascii)
        [(remove #{(codepoint \") (codepoint \\) (codepoint \/)}
                 (range 32 127))
         result]
        :else
        [(int test) result]))

(defmacro ^:private codepoint-case [e & clauses]
  `(case ~e
     ~@(mapcat codepoint-clause (partition 2 clauses))
     ~@(when (odd? (count clauses))
         [(last clauses)])))

(defn- read-array [^PushbackTextReader stream]                                               ;DM: ^PushbackReader
  ;; Expects to be called with the head of the stream AFTER the
  ;; opening bracket.
   (loop [result (transient [])]
    (let [c (.Read stream)]                                                                   ;DM: .read
      (when (neg? c)
        (throw (EndOfStreamException. "JSON error (end-of-file inside array)")))              ;DM: EOFException.
      (codepoint-case c
        :whitespace (recur result)
        \, (recur result)
        \] (persistent! result)
        (do (.Unread stream c)                                                                ;DM: .unread
            (let [element (-read stream true nil)]
              (recur (conj! result element))))))))
	
(defn- read-object [^PushbackTextReader stream]                                              ;DM: ^PushbackReader
  ;; Expects to be called with the head of the stream AFTER the
  ;; opening bracket.
  (loop [key nil, result (transient {})] 
    (let [c (.Read stream)]                                                                   ;DM: .read 
      (when (neg? c)
        (throw (EndOfStreamException. "JSON error (end-of-file inside object)")))             ;DM: EOFException.
      (codepoint-case c
        :whitespace (recur key result)  

        \, (recur nil result) 

        \: (recur key result) 

        \} (if (nil? key)
              (persistent! result)
              (throw (Exception. "JSON error (key missing value in object)")))

       (do (.Unread stream c)                                                                 ;DM: .unread 
           (let [element (-read stream true nil)]
             (if (nil? key)
               (if (string? element)
                 (recur element result)                                                       ;DM: .read 
                 (throw (Exception. "JSON error (non-string key in object)")))
               (recur nil 
                       (let [out-key (*key-fn* key)
                             out-value (*value-fn* out-key element)]
                         (if (= *value-fn* out-value)
                           result
                           (assoc! result out-key out-value)))))))))))
							
(defn- read-hex-char [^PushbackTextReader stream]                                   ;DM: ^PushbackReader
  ;; Expects to be called with the head of the stream AFTER the
  ;; initial "\u".  Reads the next four characters from the stream.
  (let [a (.Read stream)                                                             ;DM: .read
        b (.Read stream)                                                             ;DM: .read
        c (.Read stream)                                                             ;DM: .read
        d (.Read stream)]                                                            ;DM: .read
    (when (or (neg? a) (neg? b) (neg? c) (neg? d))
      (throw (EndOfStreamException.                                                  ;DM: EOFException.
	          "JSON error (end-of-file inside Unicode character escape)")))
    (let [s (str (char a) (char b) (char c) (char d))]
      (char (Int32/Parse s NumberStyles/HexNumber)))))                               ;DM: (Integer/parseInt s 16)

(defn- read-escaped-char [^PushbackTextReader stream]                               ;DM: ^PushbackReader
  ;; Expects to be called with the head of the stream AFTER the
  ;; initial backslash.
  (let [c (.Read stream)]                                                            ;DM: .read
    (codepoint-case c
      (\" \\ \/) (char c)
      \b \backspace
      \f \formfeed
      \n \newline
      \r \return
      \t \tab
      \u (read-hex-char stream))))  

(defn- read-quoted-string [^PushbackTextReader stream]                             ;DM: ^PushbackReader
  ;; Expects to be called with the head of the stream AFTER the
  ;; opening quotation mark.
  (let [buffer (StringBuilder.)]
    (loop []
      (let [c (.Read stream)]                                                       ;DM: .read
        (when (neg? c)
          (throw (EndOfStreamException. "JSON error (end-of-file inside string)")))  ;DM: EOFException.
        (codepoint-case c  
          \" (str buffer)
          \\ (do (.Append buffer (read-escaped-char stream))                        ;DM: .append
                 (recur))
          (do (.Append buffer (char c))                                             ;DM: .append
              (recur)))))))

(defn- read-integer [^String string]
  (if (< (count string) 18)  ; definitely fits in a Long
    (Int64/Parse string)                                                          ;DM: Long/valueOf
    (or (try (Int64/Parse string)                                                 ;DM: Long/valueOf
	         (catch OverflowException e nil)                                      ;DM: Added
             (catch FormatException e nil))                                       ;DM: NumberFormatException
        (clojure.lang.BigInteger/Parse string))))                                 ;DM: (bigint string) TODO: Fix when we have a BigInteger c-tor that takes a string

(defn- read-decimal [^String string]
  (if *bigdec*
    (clojure.lang.BigDecimal/Parse string)                               ;DM: (bigdec string)  -- TODO: we can change this back when we fix BigDecimal
    (Double/Parse string)))                                              ;DM: Double/valueOf

(defn- read-number [^PushbackTextReader stream]                         ;DM: ^PushbackReader
  (let [buffer (StringBuilder.)
        decimal? (loop [decimal? false]
                   (let [c (.Read stream)]                               ;DM: .read
                     (codepoint-case c
                       (\- \+ \0 \1 \2 \3 \4 \5 \6 \7 \8 \9)
                       (do (.Append buffer (char c))                     ;DM: .append
                           (recur decimal?))
                       (\e \E \.)
                       (do (.Append buffer (char c))                     ;DM: .append
                           (recur true))
                       (do (.Unread stream c)                            ;DM: .unread
                           decimal?))))]
    (if decimal?
      (read-decimal (str buffer))
      (read-integer (str buffer)))))  

(defn- -read
  [^PushbackTextReader stream eof-error? eof-value]                            ;DM: ^PushbackReader
  (loop []
    (let [c (.Read stream)]                  ;DM: .read
      (if (neg? c) ;; Handle end-of-stream
        (if eof-error?
          (throw (EndOfStreamException. "JSON error (end-of-file)"))           ;DM: EOFException.
          eof-value)
        (codepoint-case
          c
          :whitespace (recur)

          ;; Read numbers
          (\- \0 \1 \2 \3 \4 \5 \6 \7 \8 \9)
          (do (.Unread stream c)                                               ;DM: .unread
              (read-number stream))

          ;; Read strings
          \" (read-quoted-string stream)

          ;; Read null as nil
          \n (if (and (= (codepoint \u) (.Read stream))                       ;DM: .read
                      (= (codepoint \l) (.Read stream))                       ;DM: .read
                      (= (codepoint \l) (.Read stream)))                      ;DM: .read
               nil
               (throw (Exception. "JSON error (expected null)")))

          ;; Read true
          \t (if (and (= (codepoint \r) (.Read stream))                       ;DM: .read
                      (= (codepoint \u) (.Read stream))                       ;DM: .read
                      (= (codepoint \e) (.Read stream)))                      ;DM: .read
               true
               (throw (Exception. "JSON error (expected true)")))

          ;; Read false
          \f (if (and (= (codepoint \a) (.Read stream))                       ;DM: .read
                      (= (codepoint \l) (.Read stream))                       ;DM: .read
                      (= (codepoint \s) (.Read stream))                       ;DM: .read
                      (= (codepoint \e) (.Read stream)))                      ;DM: .read
               false
               (throw (Exception. "JSON error (expected false)")))

          ;; Read JSON objects
          \{ (read-object stream)

          ;; Read JSON arrays
          \[ (read-array stream)

          (throw (Exception. 
		          (str "JSON error (unexpected character): " (char c)))))))))

(defn read
  "Reads a single item of JSON data from a java.io.Reader. Options are
  key-value pairs, valid options are:

     :eof-error? boolean

        If true (default) will throw exception if the stream is empty.

     :eof-value Object

        Object to return if the stream is empty and eof-error? is
        false. Default is nil.

     :bigdec boolean

        If true use BigDecimal for decimal numbers instead of Double.
        Default is false.

     :key-fn function

        Single-argument function called on JSON property names; return
        value will replace the property names in the output. Default
        is clojure.core/identity, use clojure.core/keyword to get
        keyword properties.

     :value-fn function

        Function to transform values in the output. For each JSON
        property, value-fn is called with two arguments: the property
        name (transformed by key-fn) and the value. The return value
        of value-fn will replace the value in the output. If value-fn
        returns itself, the property will be omitted from the output.

        The default value-fn returns the value unchanged."
  [reader & options]
  (let [{:keys [eof-error? eof-value bigdec key-fn value-fn]
         :or {bigdec false
              eof-error? true
              key-fn identity
              value-fn default-value-fn}} options]
    (binding [*bigdec* bigdec
              *key-fn* key-fn
              *value-fn* value-fn]
      (-read (PushbackTextReader. reader) eof-error? eof-value))))                    ;DM: PushbackReader.

(defn read-str
  "Reads one JSON value from input String. Options are the same as for
  read."
  [string & options]
  (apply read (StringReader. string) options))

;;; JSON WRITER

(def ^:dynamic ^:private *escape-unicode*)
(def ^:dynamic ^:private *escape-slash*)

(defprotocol JSONWriter
  (-write [object out]
    "Print object to PrintWriter out as JSON"))

(defn- write-string [^String s ^TextWriter out]                                        ;DM: ^CharSequence  ^PrintWriter
  (let [sb (StringBuilder. (count s))]
    (.Append sb \")                                                                    ;DM: .append
    (dotimes [i (count s)]
      (let [cp (int (.get_Chars s i))]                                                 ;DM: (Character/codePointAt s i)
        (codepoint-case cp
          ;; Printable JSON escapes
          \" (.Append sb "\\\"")                                                       ;DM: .append
          \\ (.Append sb "\\\\")                                                       ;DM: .append
          \/ (.Append sb (if *escape-slash* "\\/" "/"))                                ;DM: .append
          ;; Simple ASCII characters
          :simple-ascii (.Append sb (.get_Chars s i))                                  ;DM: .append  .charAt
          ;; JSON escapes
          \backspace (.Append sb "\\b")                                                ;DM: .append
          \formfeed  (.Append sb "\\f")                                                ;DM: .append
          \newline   (.Append sb "\\n")                                                ;DM: .append
          \return    (.Append sb "\\r")                                                ;DM: .append
          \tab       (.Append sb "\\t")                                                ;DM: .append
          ;; Any other character is Unicode
          (if *escape-unicode*
            (.Append sb (format "\\u%04x" cp)) ; Hexadecimal-escaped                   ;DM: .append
            (.Append sb (.get_Chars s i))))))                                          ;DM: (.appendCodePoint sb cp)
    (.Append sb \")                                                                    ;DM: .append
    (.Write out (str sb))))                                               ;DM: .print

(defn- write-object [m ^TextWriter out]                                 ;DM: ^PrintWriter
  (.Write out \{)                                                       ;DM: .print
  (loop [x m]
    (when (seq m)
      (let [[k v] (first x)
            out-key (*key-fn* k)
            out-value (*value-fn* k v)]
        (when-not (string? out-key)
          (throw (Exception. "JSON object keys must be strings")))
        (when-not (= *value-fn* out-value)
          (write-string out-key out)
          (.Write out \:)                                               ;DM: .print
          (-write out-value out)))
      (let [nxt (next x)]
        (when (seq nxt)
          (.Write out \,)                                                             ;DM: .print
          (recur nxt)))))
  (.Write out \}))                                                                    ;DM: .print

(defn- write-array [s ^TextWriter out]                                                ;DM: ^PrintWriter
  (.Write out \[)                                                                     ;DM: .print
  (loop [x s]
    (when (seq x)
      (let [fst (first x)
            nxt (next x)]
        (-write fst out)
        (when (seq nxt)
          (.Write out \,)                                                             ;DM: .print
          (recur nxt)))))
  (.Write out \]))                                                                    ;DM: .print

(defn- write-bignum [x ^TextWriter out]                                               ;DM: ^PrintWriter
  (.Write out (str x)))                                                  ;DM: .print

(defn- write-plain [x ^TextWriter out]                                                ;DM: ^PrintWriter
  (.Write out x))                                                                     ;DM: .print

(defn- write-null [x ^TextWriter out]                                                 ;DM: ^PrintWriter
  (.Write out "null"))                                                                ;DM: .print

(defn- write-named [x out]
  (write-string (name x) out))

(defn- write-generic [x out]
  (if (.IsArray (class x))                                                            ;DM: isArray
    (-write (seq x) out)
    (throw (Exception. (str "Don't know how to write JSON of " (class x))))))

(defn- write-ratio [x out]
  (-write (double x) out))

;;DM: Added write-float
(defn- write-float [x ^TextWriter out] 
  (.Write out (fp-str x)))                                  

  
;DM: ;; nil, true, false
;DM: (extend nil                    JSONWriter {:-write write-null})
;DM: (extend java.lang.Boolean      JSONWriter {:-write write-plain})
;DM: 
;DM: ;; Numbers
;DM: (extend java.lang.Number       JSONWriter {:-write write-plain})
;DM: (extend clojure.lang.Ratio     JSONWriter {:-write write-ratio})
;DM: (extend clojure.lang.BigInt    JSONWriter {:-write write-bignum})
;DM: (extend java.math.BigInteger   JSONWriter {:-write write-bignum})
;DM: (extend java.math.BigDecimal   JSONWriter {:-write write-bignum})
;DM: 
;DM: ;; Symbols, Keywords, and Strings
;DM: (extend clojure.lang.Named     JSONWriter {:-write write-named})
;DM: (extend java.lang.CharSequence JSONWriter {:-write write-string})
;DM: 
;DM: ;; Collections
;DM: (extend java.util.Map          JSONWriter {:-write write-object})
;DM: (extend java.util.Collection   JSONWriter {:-write write-array})
;DM: 
;DM: ;; Maybe a Java array, otherwise fail
;DM: (extend java.lang.Object       JSONWriter {:-write write-generic})

;;DM: Following added
;; nil, true, false
(extend nil                             JSONWriter {:-write write-null})
(extend clojure.lang.Named              JSONWriter {:-write write-named})
(extend System.Boolean                  JSONWriter {:-write write-plain})

;; Numbers
;; no equivalent to java.lang.Number.  Sigh.
(extend System.Byte                     JSONWriter {:-write write-plain})
(extend System.SByte                    JSONWriter {:-write write-plain})
(extend System.Int16                    JSONWriter {:-write write-plain})
(extend System.Int32                    JSONWriter {:-write write-plain})
(extend System.Int64                    JSONWriter {:-write write-plain})
(extend System.UInt16                   JSONWriter {:-write write-plain})
(extend System.UInt32                   JSONWriter {:-write write-plain})
(extend System.UInt64                   JSONWriter {:-write write-plain})
(extend System.Double                   JSONWriter {:-write write-float})
(extend System.Single                   JSONWriter {:-write write-float})
(extend System.Decimal                  JSONWriter {:-write write-plain})
(extend clojure.lang.Ratio              JSONWriter {:-write write-ratio})
(extend clojure.lang.BigInt             JSONWriter {:-write write-bignum})
(extend clojure.lang.BigInteger         JSONWriter {:-write write-bignum})
(extend clojure.lang.BigDecimal         JSONWriter {:-write write-bignum})

;; Symbols, Keywords, and Strings
(extend clojure.lang.Named              JSONWriter {:-write write-named})
(extend System.String                   JSONWriter {:-write write-string})

;; Collections
(extend clojure.lang.IPersistentMap     JSONWriter {:-write write-object})
(extend System.Collections.IDictionary  JSONWriter {:-write write-object})
;; Cannot handle generic types!!!! 
(extend System.Collections.ICollection  JSONWriter {:-write write-array})
(extend clojure.lang.ISeq               JSONWriter {:-write write-array})

;; Maybe a Java array, otherwise fail
(extend System.Object                   JSONWriter {:-write write-generic})
;;DM: End addition

(defn write
  "Write JSON-formatted output to a java.io.Writer.    Options are 
  key-value pairs, valid options are:

    :escape-unicode boolean

       If true (default) non-ASCII characters are escaped as \\uXXXX

    :escape-slash boolean
       If true (default) the slash / is escaped as \\/

    :key-fn function

        Single-argument function called on map keys; return value will
        replace the property names in the output. Must return a
        string. Default calls clojure.core/name on symbols and
        keywords and clojure.core/str on everything else.

    :value-fn function

        Function to transform values before writing. For each
        key-value pair in the input, called with two arguments: the
        key (BEFORE transformation by key-fn) and the value. The
        return value of value-fn will replace the value in the output.
        If the return value is a number, boolean, string, or nil it
        will be included literally in the output. If the return value
        is a non-map collection, it will be processed recursively. If
        the return value is a map, it will be processed recursively,
        calling value-fn again on its key-value pairs. If value-fn
        returns itself, the key-value pair will be omitted from the
        output."
  [x writer & options]                                                                       ; ^Writer  -- can't do. we might get a TextWriter or a Stream
  (let [{:keys [escape-unicode escape-slash key-fn value-fn]
         :or {escape-unicode true
              escape-slash true
              key-fn default-write-key-fn
              value-fn default-value-fn}} options]
    (binding [*escape-unicode* escape-unicode
              *escape-slash* escape-slash
              *key-fn* key-fn
              *value-fn* value-fn]
     (-write x (if (instance? TextWriter writer) writer (StreamWriter. writer))))))    ;DM: (-write x(PrintWriter. writer))
	
(defn write-str
  "Converts x to a JSON-formatted string. Options are the same as
  write."
  [x & options]
  (let [sw (StringWriter.)]
    (apply write x sw options)
    (.ToString sw)))                                                  ;DM: .toString

;;; JSON PRETTY-PRINTER

;; Based on code by Tom Faulhaber

(defn- pprint-array [s] 
  ((pprint/formatter-out "~<[~;~@{~w~^, ~:_~}~;]~:>") s))

(defn- pprint-object [m]
  ((pprint/formatter-out "~<{~;~@{~<~w:~_~w~:>~^, ~_~}~;}~:>") 
   (for [[k v] m] [(*key-fn* k) v])))

(defn- pprint-generic [x]
  (if (.IsArray (class x))                                                   ;DM: isArray
    (pprint-array (seq x))
    ;; pprint proxies Writer, so we can't just wrap it	
    (print (with-out-str (-write x (if (instance? TextWriter *out*) *out* (StreamWriter. *out*)))))))           ; DM: PrintWriter

(defn- pprint-dispatch [x]
  (cond (nil? x) (print "null")
        (true? x) (print "true")                                                ;DM: Added
		(false? x) (print "false")                                              ;DM: Added
        (instance? System.Collections.IDictionary x) (pprint-object x)     ;DM: java.util.Map
        (instance? System.Collections.ICollection x) (pprint-array x)      ;DM: java.util.Collection
        (instance? clojure.lang.ISeq x) (pprint-array x)
        :else (pprint-generic x)))

(defn pprint
  "Pretty-prints JSON representation of x to *out*. Options are the
  same as for write except :value-fn, which is not supported."
  [x & options]
  (let [{:keys [escape-unicode escape-slash key-fn]
         :or {escape-unicode true
              escape-slash true
              key-fn default-write-key-fn}} options]
    (binding [*escape-unicode* escape-unicode
              *escape-slash* escape-slash
              *key-fn* key-fn]
      (pprint/write x :dispatch pprint-dispatch))))
	
;; Local Variables:
;; mode: clojure
;; eval: (define-clojure-indent (codepoint-case (quote defun)))
;; End:	