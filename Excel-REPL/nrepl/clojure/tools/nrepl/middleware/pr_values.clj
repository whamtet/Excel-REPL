
(ns ^{:author "Chas Emerick"}
     clojure.tools.nrepl.middleware.pr-values
  (:require [clojure.tools.nrepl.transport :as t]
  			[clojure.tools.nrepl.debug :as debug])                ;DM: Added
  (:use [clojure.tools.nrepl.middleware :only (set-descriptor!)])
  (:import clojure.tools.nrepl.transport.Transport))

(defn pr-values
  "Middleware that returns a handler which transforms any :value slots
   in messages sent via the request's Transport to strings via `pr`,
   delegating all actual message handling to the provided handler.

   Requires that results of eval operations are sent in messages in a
   :value slot."
  [h]
  (fn [{:keys [op ^Transport transport] :as msg}]
    (h (assoc msg :transport (let [wt (reify Transport
                               (recv [this] (.recv transport))
                               (recv [this timeout] (.recv transport timeout))
                               (send [this resp]
							     #_(debug/prn-thread "pr-values - sending on to " (.GetHashCode transport))
                                 (.send transport
                                   (if-let [[_ v] (find resp :value)]
                                     (let [repr (System.IO.StringWriter.)]                        ;;; java.io.StringWriter.
                                       (assoc resp :value (do (if *print-dup*
                                                                (print-dup v repr)
                                                                (print-method v repr))
                                                              (str repr))))
                                     resp))
                                 this))]
								 #_(debug/prn-thread "pr-values - reify, wrapping " (.GetHashCode wt) " around " (.GetHashCode transport)) 
								  wt)
								 ))))

(set-descriptor! #'pr-values
  {:requires #{}
   :expects #{}
   :handles {}})