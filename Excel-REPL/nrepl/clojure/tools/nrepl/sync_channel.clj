;-
;   Copyright (c) David Miller. All rights reserved.
;   The use and distribution terms for this software are covered by the
;   Eclipse Public License 1.0 (http://opensource.org/licenses/eclipse-1.0.php)
;   which can be found in the file epl-v10.html at the root of this distribution.
;   By using this software in any fashion, you are agreeing to be bound by
;   the terms of this license.
;   You must not remove this notice, or any other, from this software.


(ns #^{:author "David Miller"
       :doc "A simple synchronous channel"}
  clojure.tools.nrepl.sync-channel
  (:refer-clojure :exclude (take)))
  
 ;; Reason for existence
 ;;
 ;; The original (ClojureJVM) FnTransport code uses a java.util.concurrent.SynchronousQueue.
 ;; However, in that use there is a single producer and a single consumer.
 ;; The CLR does not supply such a construct.  
 ;; (The closest equivalent would be a System.Collections.Concurrent.BlockingCollection<T> with zero capacity, 
 ;; but that class only allows capacity greater than zero.)
 ;;
 ;; Rather then do a full-blown implementation of a synchronous queue, 
 ;; something along the lines of Doug Lea's C# implementation
 ;; http://code.google.com/p/netconcurrent/source/browse/trunk/src/Spring/Spring.Threading/Threading/Collections/SynchronousQueue.cs
 ;; we can go with a much simpler construct - a synchronous channel between producer and consumer
;;  that blocks either one if the other is not waiting.
 
(defprotocol SyncChannel
  "A synchronous channel (single-threaded on producer and consumer)"
  (put   [this value]   "Put a value to this channel (Producer)")
  (take  [this]         "Get a value from this channel (Consumer)")
  (poll  [this] [this timeout]    "Get a value from this channel if one is available (within the designated timeout period)"))
  

(definterface ITakeValue 
   (takeValue []))
  
 ;; SimpleSyncChannel assumes there is a single producer thread and a single consumer thread.
 (deftype SimpleSyncChannel [#^:volatile-mutable value 
                             #^:volatile-mutable c-waiting?
							 #^:volatile-mutable p-waiting?
							 lock]
  SyncChannel
  (put [this v] 
    (when (nil? v)
	  (throw (NullReferenceException. "Cannot put nil on SyncChannel")))
    (locking lock
	  (when p-waiting?
	     (throw (Exception. "Producer not single-threaded")))
	  (set! value v)
      (System.Threading.Monitor/Pulse lock)
	  (set! p-waiting? true)
      (System.Threading.Monitor/Wait lock)
      (set! p-waiting? false)))
	     
  (poll [this] 
    (locking lock
	  (when c-waiting?
	    (throw (Exception. "Consumer not single-threaded")))
	  (when-not (nil? value)
	    (.takeValue ^ITakeValue this))))
		  
  (poll [this timeout]
    (locking lock
	  (when c-waiting?
	    (throw (Exception. "Consumer not single-threaded")))
      (if (nil? value)
        (do 
		  (set! c-waiting? true)
		  (let [result 
		        (when (System.Threading.Monitor/Wait lock (int timeout))
		           (.takeValue ^ITakeValue this))]
		    (set! c-waiting? false)
			result))
	    (.takeValue ^ITakeValue this))))
		
  (take [this]
    (locking lock
	   (when c-waiting?
	   	     (throw (Exception. "Consumer not single-threaded")))
      (when (nil? value)
        (set! c-waiting? true)
        (System.Threading.Monitor/Wait lock)
		(set! c-waiting? false))
	  (.takeValue ^ITakeValue this)))

  ITakeValue
  (takeValue [this]
    (let [curval value]
      (set! value nil)
      (System.Threading.Monitor/Pulse lock)
	  curval)))

(defn make-simple-sync-channel []
  (SimpleSyncChannel. nil false false (Object.)))
  
  
(comment
   
(def prn-agent (agent nil))
(defn sprn [& strings] (send-off prn-agent (fn [v] (apply prn strings))))  
(defn f [n]
   (let [sc (make-simple-sync-channel)
	      p (agent nil)
		  c (agent nil)]
	  (send c (fn [v] (dotimes [i n] (sprn (str "Consumer " i)) (sprn (str "====> "(take sc))))))
	  (send p (fn [v] (dotimes [i n] (sprn (str "Producer " i)) (put sc i))))
	  [p c sc]))
)	  
	  
   