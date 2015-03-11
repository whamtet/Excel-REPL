;ns client

(require '[clojure.tools.nrepl :as nrepl])
(require '[clojure.data.drawbridge-client :as drawbridge-client])

(def client (atom nil))

(defn set-client! [s]
  (let [
        s (if (string? s) s (str "nrepl://localhost:" s))
        ]
    (reset! client (nrepl/client (nrepl/url-connect s) 10000))))

(defn remote-eval [code]
(->
@client
(nrepl/message (hash-map :op "eval" :code code))
nrepl/response-values))

(defmacro eval2 [& body]
`(first (remote-eval (nrepl/code ~@body))))
