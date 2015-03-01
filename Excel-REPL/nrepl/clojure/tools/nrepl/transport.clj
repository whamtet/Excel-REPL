(ns ^{:author "Chas Emerick"}
     clojure.tools.nrepl.transport
  (:require [clojure.tools.nrepl.bencode :as be]
            [clojure.clr.io :as io]                                       ;DM: clojure.java.io
			[clojure.tools.nrepl.debug :as debug]
			[clojure.tools.nrepl.sync-channel :as sc]                    ;DM: Added
            clojure.main                                                 ;Matt: Added
            (clojure walk set))
  (:use [clojure.tools.nrepl.misc :only (returning uuid)])
  (:refer-clojure :exclude (send))
  (:import (System.IO  Stream  EndOfStreamException)                      ;DM: (java.io InputStream OutputStream PushbackInputStream
           (clojure.lang PushbackInputStream PushbackTextReader)          ;DM:  PushbackReader IOException EOFException)
           (System.Net.Sockets Socket SocketException)                    ;DM: (java.net Socket SocketException)
           (System.Collections.Concurrent                                 ;DM: (java.util.concurrent SynchronousQueue LinkedBlockingQueue
               |BlockingCollection`1[System.Object]|)                     ;DM: BlockingQueue TimeUnit)
;		   clojure.tools.nrepl.transport.Transport                        ;DM: Added
           clojure.lang.RT ))

(defprotocol Transport
  "Defines the interface for a wire protocol implementation for use
   with nREPL."
  (recv [this] [this timeout]
    "Reads and returns the next message received.  Will block.
     Should return nil the a message is not available after `timeout`
     ms or if the underlying channel has been closed.")
  (send [this msg] "Sends msg. Implementations should return the transport."))

(deftype FnTransport [recv-fn send-fn close]
  Transport
  ;; TODO this keywordization/stringification has no business being in FnTransport
  (send [this msg] #_(debug/prn-thread "FnTransport:: send " msg) (-> msg clojure.walk/stringify-keys send-fn) this)
  (recv [this] #_(debug/prn-thread "FnTransprot:: recv ") (.recv this Int32/MaxValue))                                      ;DM: Long/MAX_VALUE
  (recv [this timeout] #_(debug/prn-thread "FnTransport:: recv [" timeout "]") (clojure.walk/keywordize-keys (recv-fn timeout)))
  System.IDisposable                                                             ;DM: java.io.Closeable
  (Dispose [this] #_(debug/prn-thread "FnTranpsort:: Dispose " (.GetHashCode this)) (close)))                                                      ;DM: (close [this] (close)))  TODO: This violates good IDisposable practice

(defn fn-transport
  "Returns a Transport implementation that delegates its functionality
   to the 2 or 3 functions provided."
  ([read write] (fn-transport read write nil))
  ([read write close]
    (let [read-queue (sc/make-simple-sync-channel)                                 ;DM: (SynchronousQueue.)
	      msg-pump (future (try
                             (while true
				               #_(debug/prn-thread "fn-transport:: ready to read")
  			                   (sc/put read-queue (read))                                       ;DM: .put
				               #_(debug/prn-thread "fn-transport:: put to queue"))   ;DEBUG
                             (catch Exception t                                                ;DM: Throwable
				               #_(debug/prn-thread "fn-transport:: caught exception!!!!")
                              (sc/put read-queue t))))]                                        ;DM: .put
      (FnTransport.
        (let [failure (atom nil)]
          #(if @failure
             (throw @failure)
             (let [msg (sc/poll read-queue % )]                                   ;DM:  .poll, remove TimeUnit/MILLISECONDS
			   #_(debug/prn-thread "fn-transport:: read from queue: " (let [mstr (str msg)] (if (< (count mstr) 75) mstr (subs mstr 0 75)))) ;DEBUG
               (if (instance? Exception msg)                                      ;DM: Throwable
                 (do #_(debug/prn-thread "fn-transport:: read Exception: " (let [mstr (str msg)] (if (< (count mstr) 75) mstr (subs mstr 0 75)))) (reset! failure msg) (throw msg))
                 msg))))
        write
        (fn [] (close) #_(future-cancel msg-pump))))))

(defmulti #^{:private true} <bytes class)

(defmethod <bytes :default
  [input]
  input)
                                                                           ;DM:Added
(def #^{:private true} utf8 (System.Text.UTF8Encoding.))          ;DM:Added
                                                                           ;DM:Added
(defmethod <bytes |System.Byte[]|                                          ;DM: (RT/classForName "[B")
  [#^|System.Byte[]| input]                                                ;DM: #^"[B"
  (.GetString utf8 input))                                                 ;DM: (String. input "UTF-8"))

(defmethod <bytes clojure.lang.IPersistentVector
  [input]
  (vec (map <bytes input)))

(defmethod <bytes clojure.lang.IPersistentMap
  [input]
  (->> input
    (map (fn [[k v]] [k (<bytes v)]))
    (into {})))

(defmacro ^{:private true} rethrow-on-disconnection
  [^Socket s & body]
  `(try
     #_(debug/prn-thread "rethrow-on-disconnection: begin body, socket = " (.GetHashCode ~s))
     ~@body
     (catch EndOfStreamException e#                                        ;DM: EOFException
	   #_(debug/prn-thread "rethrow-on-disconnection: EndOfStreamException, socket = " (.GetHashCode ~s))
       (throw (ObjectDisposedException. "The transport's socket appears to have lost its connection to the nREPL server" e#)))
     (catch Exception e#                                                   ;DM: Throwable
       #_(debug/prn-thread "rethrow-on-disconnection: Exception: socket = " (.GetHashCode ~s) ", connected = "(and ~s (not (.Connected ~s))))
       (if (or (instance? SocketException (#'clojure.main/root-cause e#))
	           (and ~s (not (.Connected ~s))))                                  ;DM: .isConnected
         (throw (ObjectDisposedException. "The transport's socket appears to have lost its connection to the nREPL server" e#))
         (throw e#)))))

(def bencode-sockets (atom ())) ;DEBUG

(defn bencode
  "Returns a Transport implementation that serializes messages
   over the given Socket or InputStream/OutputStream using bencode."
  ([^Socket s] (bencode s s s))
  ([in out & [^Socket s]]
    #_(debug/prn-thread "Creating bencode")  ;DEBUG
	(swap! bencode-sockets conj s)
    (let [in (PushbackInputStream. (io/input-stream in))
          out (io/output-stream out)]
      (fn-transport
        #(let [payload (rethrow-on-disconnection s (be/read-bencode in))
               unencoded (<bytes (payload "-unencoded"))
               to-decode (apply dissoc payload "-unencoded" unencoded)]
           (merge (dissoc payload "-unencoded")
                  (when unencoded {"-unencoded" unencoded})
                  (<bytes to-decode)))
        #(rethrow-on-disconnection s
           (locking out
             (doto out
               (be/write-bencode %)
               .Flush)))                                                  ;DM: .flush
        (fn []
          (if s
            (.Close s)                                                    ;DM: .close
            (do
              (.Close in)                                                 ;DM: .close
              (.Close out))))))))                                         ;DM: .close

(defn tty
  "Returns a Transport implementation suitable for serving an nREPL backend
   via simple in/out readers, as with a tty or telnet connection."
  ([^Socket s] (tty s s s))
  ([in out & [^Socket s]]
    (let [r (PushbackTextReader. (io/text-reader in))                     ;DM: PushbackReader. io/reader
          w (io/text-writer out)                                          ;DM: io/writer
          cns (atom "user")
          prompt (fn [newline?]
                   (when newline? (.Write w (int \newline)))              ;DM: .write
                   (.Write w (str @cns "=> ")))                           ;DM: .write
          session-id (atom nil)
          read-msg #(let [code (read r)]
                      (merge {:op "eval" :code [code] :ns @cns :id (str "eval" (uuid))}
                             (when @session-id {:session @session-id})))
          read-seq (atom (cons {:op "clone"} (repeatedly read-msg)))
          write (fn [{:strs [out err value status ns new-session id] :as msg}]
                  (when new-session (reset! session-id new-session))
                  (when ns (reset! cns ns))
                  (doseq [^String x [out err value] :when x]
                    (.Write w x))                                                    ;DM: .write
                  (when (and (= status #{:done}) id (.startsWith ^String id "eval"))
                    (prompt true))
                  (.Flush w))                                                        ;DM: .flush
          read #(let [head (promise)]
                  (swap! read-seq (fn [s]
                                     (deliver head (first s))
                                     (rest s)))
                  @head)]
      (fn-transport read write
        (when s
          (swap! read-seq (partial cons {:session @session-id :op "close"}))
          #(.Close s))))))                                                         ;DM: .close

(defn tty-greeting
  "A greeting fn usable with clojure.tools.nrepl.server/start-server,
   meant to be used in conjunction with Transports returned by the
   `tty` function.

   Usually, Clojure-aware client-side tooling would provide this upon connecting
   to the server, but telnet et al. isn't that."
  [transport]
  (send transport {:out (str ";; Clojure " (clojure-version)
                             \newline "user=> ")}))

(deftype QueueTransport [^|System.Collections.Concurrent.BlockingCollection`1[System.Object]| in
                         ^|System.Collections.Concurrent.BlockingCollection`1[System.Object]| out]           ;DM: ^BlockingQueue
  clojure.tools.nrepl.transport.Transport
  (send [this msg] (.Add out msg) this)                                            ;DM: .put
  (recv [this] (.Take in))                                                         ;DM: .take
  (recv [this timeout] (let [x nil] (.TryTake in (by-ref x) (int timeout)) x)))    ;DM: .poll, removed TimeUnit/MILLISECONDS, added (int .), let, ref

(defn piped-transports
  "Returns a pair of Transports that read from and write to each other."
  []
  (let [a (|System.Collections.Concurrent.BlockingCollection`1[System.Object]|.)                                                    ;DM: LinkedBlockingQueue
        b (|System.Collections.Concurrent.BlockingCollection`1[System.Object]|.)]                                                   ;DM: LinkedBlockingQueue
    [(QueueTransport. a b) (QueueTransport. b a)]))
