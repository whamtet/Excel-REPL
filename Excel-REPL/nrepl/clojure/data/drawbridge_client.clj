(ns clojure.data.drawbridge-client
  (:require [clojure.data.json :as json]
            [clojure.tools.nrepl :as nrepl]
            [clr-http.lite.client :as http]
            )
  (:import ClojureExcel.MainClass
           System.IO.StreamReader
           ))

;DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond
(defn time-millis []
  (/ (.Ticks DateTime/Now) TimeSpan/TicksPerMillisecond))

(defn ring-client-transport
  "Returns an nREPL client-side transport to connect to HTTP nREPL
  endpoints implemented by `ring-handler`.

  This fn is implicitly registered as the implementation of
  clojure.tools.nrepl/url-connect for `http` and `https` schemes;
  so, once this namespace is loaded, any tool that uses url-connect
  will use this implementation for connecting to HTTP and HTTPS
  nREPL endpoints."
  [url]
  (let [incoming (MainClass/GetCollection); a bit of a hack
        fill #(when-let [responses (->> %
                                        StreamReader.
                                        line-seq
                                        rest
                                        drop-last
                                        (remove empty?)
                                        (map json/read-str)
                                        (remove nil?)
                                        seq)]
                (doseq [response responses]
                  (.Add incoming response)))

        session-cookies (atom nil)

        http (fn [& [msg]]
               (let [
                     req-map (merge {:as :stream
                                     :cookies @session-cookies}
                                    (when msg {:form-params msg}))
                     {:keys [cookies body] :as resp} ((if msg http/post http/get)
                                                      url
                                                      req-map)]
                 (println "cookies" cookies)
                 (swap! session-cookies merge cookies)
                 (fill body)))
        poll #(MainClass/TakeItem incoming)
        ]
    (clojure.tools.nrepl.transport.FnTransport.
     (fn read [timeout]
       (let [t (time-millis)]
         (or (poll)
             (when (pos? timeout)
               (http)
               (recur (- timeout (- (time-millis) t)))))))
     http
     (fn close []))))


;(.removeMethod nrepl/url-connect "http")
;(.removeMethod nrepl/url-connect "https")

(.addMethod nrepl/url-connect "http" #'ring-client-transport)
(.addMethod nrepl/url-connect "https" #'ring-client-transport)

