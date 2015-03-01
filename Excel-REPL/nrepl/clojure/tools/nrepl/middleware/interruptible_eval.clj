(ns ^{:author "Chas Emerick"}
     clojure.tools.nrepl.middleware.interruptible-eval
  (:require [clojure.tools.nrepl.transport :as t]
            clojure.tools.nrepl.middleware.pr-values
  			[clojure.tools.nrepl.debug :as debug]			
            clojure.main)
  (:use [clojure.tools.nrepl.misc :only (response-for returning)]
        [clojure.tools.nrepl.middleware :only (set-descriptor!)])
  (:import clojure.lang.LineNumberingTextReader                                            ;DM: clojure.lang.LineNumberingPushbackReader
           (System.IO StringReader TextWriter)                                             ;DM: (java.io StringReader Writer)
           clojure.lang.AtomicLong                                                         ;DM: java.util.concurrent.atomic.AtomicLong
           (System.Threading Thread ThreadStart WaitCallback ThreadAbortException)         ;DM: (java.util.concurrent Executor LinkedBlockingQueue ThreadFactory
            ))                                                                             ;DM:                      SynchronousQueue TimeUnit ThreadPoolExecutor)

(def ^{:dynamic true
       :doc "The message currently being evaluated."}
      *msg* nil)

(def ^{:dynamic true
       :doc "Function returning the evaluation of its argument."}
       *eval* nil)
	  
(defn- capture-thread-bindings
  "Capture thread bindings, excluding nrepl implementation vars."
  []
  (dissoc (get-thread-bindings) #'*msg* #'*eval*))

(defn evaluate
  "Evaluates some code within the dynamic context defined by a map of `bindings`,
   as per `clojure.core/get-thread-bindings`.

   Uses `clojure.main/repl` to drive the evaluation of :code in a second
   map argument (either a string or a seq of forms to be evaluated), which may
   also optionally specify a :ns (resolved via `find-ns`).  The map MUST
   contain a Transport implementation in :transport; expression results and errors
   will be sent via that Transport.

   Returns the dynamic scope that remains after evaluating all expressions
   in :code.

   It is assumed that `bindings` already contains useful/appropriate entries
   for all vars indicated by `clojure.main/with-bindings`."
  [bindings {:keys [code ns transport session eval] :as msg}]
  (let [explicit-ns-binding (when-let [ns (and ns (-> ns symbol find-ns))]
                              {#'*ns* ns})
        bindings (atom (merge bindings explicit-ns-binding))
        session (or session (atom nil))		
        out (@bindings #'*out*)
        err (@bindings #'*err*)]
    (if (and ns (not explicit-ns-binding))
      (t/send transport (response-for msg {:status #{:error :namespace-not-found :done}}))
      (with-bindings @bindings
        (try
		  #_(debug/prn-thread "Evaluating " code " in " (.ManagedThreadId (Thread/CurrentThread))) ;DEBUG
          (clojure.main/repl
            :eval (if eval (find-var (symbol eval)) clojure.core/eval)
		  ;; clojure.main/repl paves over certain vars even if they're already thread-bound
            :init #(do (set! *compile-path* (@bindings #'*compile-path*))
                     (set! *1 (@bindings #'*1))
                     (set! *2 (@bindings #'*2))
                     (set! *3 (@bindings #'*3))
                     (set! *e (@bindings #'*e)))   
            :read (if (string? code)
                    (let [reader (LineNumberingTextReader. (StringReader. code))]               ;DM: LineNumberingPushbackReader
                      #(read reader false %2))
                    (let [^System.Collections.IEnumerator code (.GetEnumerator code)]           ;DM:  .iterator
                      #(or (and (.MoveNext code) (.Current code)) %2)))                         ;DM: .hasNext .next
            :prompt (fn [])
            :need-prompt (constantly false)
            ; TODO pretty-print?
            :print (fn [v]
                     (reset! bindings (assoc (capture-thread-bindings)
                                             #'*3 *2
                                             #'*2 *1
                                             #'*1 v))
                     (.Flush ^TextWriter err)                                                   ;DM: .flush ^Writer
                     (.Flush ^TextWriter out)                                                   ;DM: .flush ^Writer
                     (reset! session @bindings)					 
					 #_(debug/prn-thread "Evaluating " code " yields " v) ;DEBUG
                     (t/send transport (response-for msg
                                                     {:value v
                                                      :ns (-> *ns* ns-name str)})))
            ; TODO customizable exception prints
            :caught (fn [e]
                      (let [root-ex (#'clojure.main/root-cause e)]
                        (when-not (instance? ThreadAbortException root-ex)                      ;DM: ThreadDeath
                          (reset! bindings (assoc (capture-thread-bindings) #'*e e))
                          (t/send transport (response-for msg {:status :eval-error
                                                               :ex (-> e class str)
                                                               :root-ex (-> root-ex class str)}))
                          (clojure.main/repl-caught e)))))
          (finally
            (.Flush ^TextWriter out)                                                            ;DM: .flush ^Writer
            (.Flush ^TextWriter err)))))                                                        ;DM: .flush ^Writer
    @bindings))

;(defn- configure-thread-factory
;  "Returns a new ThreadFactory for the given session.  This implementation
;   generates daemon threads, with names that include the session id."
;  []
;  (let [session-thread-counter (AtomicLong. 0)]
;    (reify ThreadFactory
;      (newThread [_ runnable]
;        (doto (Thread. runnable
;                (format "nREPL-worker-%s" (.getAndIncrement session-thread-counter)))
;          (.setDaemon true))))))
;
;(def ^{:private true} jdk6? (try
;                              (Class/forName "java.util.ServiceLoader")
;                              true
;                              (catch ClassNotFoundException e false)))
;
;; this is essentially the same as Executors.newCachedThreadPool, except
;; for the JDK 5/6 fix described below
;(defn- configure-executor
;  "Returns a ThreadPoolExecutor, configured (by default) to
;   have no core threads, use an unbounded queue, create only daemon threads,
;   and allow unused threads to expire after 30s."
;  [& {:keys [keep-alive queue thread-factory]
;      :or {keep-alive 30000
;           queue (SynchronousQueue.)}}]
;  (let [^ThreadFactory thread-factory (or thread-factory (configure-thread-factory))]
;    ; ThreadPoolExecutor in JDK5 *will not run* submitted jobs if the core pool size is zero and
;    ; the queue has not yet rejected a job (see http://kirkwylie.blogspot.com/2008/10/java5-vs-java6-threadpoolexecutor.html)
;    (ThreadPoolExecutor. (if jdk6? 0 1) Integer/MAX_VALUE
;                         (long 30000) TimeUnit/MILLISECONDS
;                         ^BlockingQueue queue
;                         thread-factory)))

;DM:Added
(def ^{:private true} session-thread-counter (AtomicLong. 0))

#_(defn- exec-eval [f]
  (let [tstart (gen-delegate ThreadStart []
                 (try 
				  #_(debug/prn-thread "exec-eval: Starting in thread " (.ManagedThreadId (Thread/CurrentThread)))
				  (f) 
				  #_(debug/prn-thread "exec-eval: Exiting thread " (.ManagedThreadId (Thread/CurrentThread)))
				  (catch ThreadAbortException e 
				    #_(debug/prn-thread "exec-eval: Aborting thread " (.ManagedThreadId (Thread/CurrentThread)))
					#_(Thread/ResetAbort)
					nil)))
        thread (doto (Thread. tstart)
                 (.set_Name (format "nREPL-worker-%s" (.getAndIncrement session-thread-counter)))
	             (.set_IsBackground true)
	             (.Start))]
	 #_(debug/prn-thread "exec-eval: Started thread " (.ManagedThreadId thread))
	 nil))
	
(defn- exec-eval [f interrupt-handle]
  (let [done-handle (System.Threading.AutoResetEvent. false)
        handles (make-array System.Threading.WaitHandle 2)
        tstart (gen-delegate ThreadStart []
                 (try 
				  #_(debug/prn-thread "exec-eval: Starting in thread " (.ManagedThreadId (Thread/CurrentThread)))
				  (f) 
				  #_(debug/prn-thread "exec-eval: Exiting thread " (.ManagedThreadId (Thread/CurrentThread)))
				  (catch ThreadAbortException e 
				    #_(debug/prn-thread "exec-eval: Aborting thread " (.ManagedThreadId (Thread/CurrentThread)))
					(Thread/ResetAbort)
					nil)
				  (finally (.Set done-handle))))
        thread (doto (Thread. tstart)
                 (.set_Name (format "nREPL-worker-%s" (.getAndIncrement session-thread-counter)))
	             (.set_IsBackground true)
	             (.Start))]
	 #_(debug/prn-thread "exec-eval: Started thread " (.ManagedThreadId thread))
	 #_(debug/prn-thread "exec-eval: Starting wait")
	 (aset handles 0 interrupt-handle)
	 (aset handles 1 done-handle)
	 (let [i (System.Threading.WaitHandle/WaitAny handles)]
	   #_(debug/prn-thread "exec-eval: done waiting, handle = " i)
	   (when (= i 0)
	     #_(debug/prn-thread "exec-eval: interrupted, aborting thread")
	     (.Abort thread))
	   #_(when (= i 1)
	     (debug/prn-thread "exec.eval: normal exit")))
	 nil))
		
;DM:end Added

; A little mini-agent implementation. Needed because agents cannot be used to host REPL
; evaluation: http://dev.clojure.org/jira/browse/NREPL-17
(defn- prep-session
  [session]
  (locking session
    (returning session
      (when-not (-> session meta :queue)
        (alter-meta! session assoc :queue (atom clojure.lang.PersistentQueue/EMPTY))))))

(declare run-next)
(defn- run-next*
  [session executor ihandle]                               ;DM: removed ^Executor
  #_(debug/prn-thread "run-next* on session ") ;DEBUG
  (let [qa (-> session meta :queue)]
    (loop []
      (let [q @qa
            qn (pop q)]
        (if-not (compare-and-set! qa q qn)
          (recur)
          (when (seq qn)
		    (let [fnext (run-next session executor ihandle (peek qn))]
            (exec-eval fnext ihandle))))))))
			;DM: (.execute executor (run-next session executor (peek qn)))

(defn- run-next
  [session executor ihandle f]
  #(try
     #_(debug/prn-thread "run-next: ready to run f, thread = " (.ManagedThreadId (Thread/CurrentThread)))
     (f)
	 #_(debug/prn-thread "run-next: after running f, thread = " (.ManagedThreadId (Thread/CurrentThread)))
     (finally
       #_(debug/prn-thread "run-next: looping, thread = " (.ManagedThreadId (Thread/CurrentThread)))
       (run-next* session executor ihandle))))

(defn- queue-eval
  "Queues the function for the given session."
  [session executor ihandle f]                                                                   ;DM: removed ^Executor
  (let [qa (-> session prep-session meta :queue)]
    (loop []
      (let [q @qa]
        (if-not (compare-and-set! qa q (conj q f))
          (recur)
		  (when (empty? q)
			(let [fnext (run-next session executor ihandle f)]
              (exec-eval fnext ihandle))))))))
			;DM: (.execute executor (run-next session executor f))

(defn interruptible-eval
  "Evaluation middleware that supports interrupts.  Returns a handler that supports
   \"eval\" and \"interrupt\" :op-erations that delegates to the given handler
   otherwise."
  [h & {:keys [executor] :or {executor nil}}]                                    ;DM: (configure-executor) replaced with nil
  (let [interrupt-handle (System.Threading.AutoResetEvent. false)]
  (fn [{:keys [op session interrupt-id id transport] :as msg}]
    (case op
      "eval"
      (if-not (:code msg)
        (do #_(debug/prn-thread "IEval: no code: " msg) (t/send transport (response-for msg :status #{:error :no-code})))
        (queue-eval session executor interrupt-handle
          (fn []
            (alter-meta! session assoc
                         :thread (Thread/CurrentThread)                                  ;DM: Thread/currentThread
                         :eval-msg msg)
            (binding [*msg* msg]
  		      #_(debug/prn-thread "IEval: getting ready to call evaluate, thread = " (.ManagedThreadId (Thread/CurrentThread)))
              (evaluate @session msg)
			  #_(debug/prn-thread "IEval: sending status done")
			  (t/send transport (response-for msg :status :done))
              (alter-meta! session dissoc :thread :eval-msg)))))			  
    
      "interrupt"
      ; interrupts are inherently racy; we'll check the agent's :eval-msg's :id and
      ; bail if it's different than the one provided, but it's possible for
      ; that message's eval to finish and another to start before we send
      ; the interrupt / .stop.
      (let [{:keys [id eval-msg ihandle]} (meta session)]  ;;; ^Thread thread
	    #_(debug/prn-thread "IEval: interrupt received")
		#_(debug/prn-thread "IEval: interrupt-id = " interrupt-id ", id = " (:id eval-msg))
		#_(debug/prn-thread "IEval: ihandle = " ihandle)
		#_(debug/prn-thread "IEval: interrupt thread = " (and thread (.ManagedThreadId thread)))
		#_(if (or (not interrupt-id)
                (= interrupt-id (:id eval-msg)))
	      (if-not thread
		     (debug/prn-thread "IEval: interrupt: Sending status :done :session-idle")
			 (debug/prn-thread "IEval: interrupt: aborting thread, sending status :interrupted"))
		  (debug/prn-thread "IEval: interrupt: sending interrupt-id-mismatch"))
        (if (or (not interrupt-id)
                (= interrupt-id (:id eval-msg)))
          (if-not ihandle                                                              ;;; thread
            (t/send transport (response-for msg :status #{:done :session-idle}))
            (do
              ; notify of the interrupted status before we .stop the thread so
              ; it is received before the standard :done status (thereby ensuring
              ; that is stays within the scope of a clojure.tools.nrepl/message seq)
			  #_(debug/prn-thread "IEval: interrupt: sending :interrupted status message")
              (t/send transport {:status #{:interrupted}
                                 :id (:id eval-msg)
                                 :session id})
			  #_(debug/prn-thread "IEval: interrupt: preparing to abort thread " #_(.ManagedThreadId thread))			  
              #_(.Abort thread)                                                   ;DM: .stop
			  (.Set ihandle)
			  #_(debug/prn-thread "IEval: interrupt: thread .Abort called")
			  #_(debug/prn-thread "IEval: interrupt: preparing to send :done status")
              (t/send transport (response-for msg :status #{:done}))
			  #_(debug/prn-thread "IEval: interrupt: preparing to send :done status AGAIN")
              (t/send transport (response-for msg :status #{:done}))
			  
			  ))
          (t/send transport (response-for msg :status #{:error :interrupt-id-mismatch :done}))))
      
      (h msg)))))

(set-descriptor! #'interruptible-eval
  {:requires #{"clone" "close" #'clojure.tools.nrepl.middleware.pr-values/pr-values}
   :expects #{}
   :handles {"eval"
             {:doc "Evaluates code."
              :requires {"code" "The code to be evaluated."
                         "session" "The ID of the session within which to evaluate the code."}
              :optional {"id" "An opaque message ID that will be included in responses related to the evaluation, and which may be used to restrict the scope of a later \"interrupt\" operation."}
              :returns {}}
             "interrupt"
             {:doc "Attempts to interrupt some code evaluation."
              :requires {"session" "The ID of the session used to start the evaluation to be interrupted."}
              :optional {"id" "An opaque message ID that will be included in responses related to the evaluation, and which may be used to restrict the scope of a later \"interrupt\" operation."
                         "eval" "A fully-qualified symbol naming a var whose function value will be used to evaluate [code], instead of `clojure.core/eval` (the default)."}
              :returns {"status" "'interrupted' if an evaluation was identified and interruption will be attempted
'session-idle' if the session is not currently evaluating any code
'interrupt-id-mismatch' if the session is currently evaluating code sent using a different ID than specified by the \"interrupt-id\" value "}}}})