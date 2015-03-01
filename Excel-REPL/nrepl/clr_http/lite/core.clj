(ns clr-http.lite.core
  "Core HTTP request/response implementation."
  (:require [clojure.clr.io :as io]
            [clr-http.lite.util :as util]
            [clr-http.lite.cookies :as cookies]
            )
  (:import
   System.Net.WebRequest
   System.Net.CookieContainer))

(defn safe-conj [a b]
  (if (vector? a)
    (conj a b)
    [a b]))

(defn parse-headers
  "Takes a URLConnection and returns a map of names to values.

  If a name appears more than once (like `set-cookie`) then the value
  will be a vector containing the values in the order they appeared
  in the headers."
  [conn]
  (let [headers (.Headers conn)]
    (apply merge-with
           safe-conj
           (for [header (.Headers conn)]
             {header (.Get headers header)}))))

(defn- coerce-body-entity
  "Coerce the http-entity from an HttpResponse to either a byte-array, or a
  stream that closes itself and the connection manager when closed."
  [{:keys [as]} conn]
  (let [ins (.GetResponseStream conn)]
    (if (or (= :stream as) (nil? ins))
      ins
      (util/to-byte-array ins))))

(defn request
  "Executes the HTTP request corresponding to the given Ring request map and
  returns the Ring response map corresponding to the resulting HTTP response.
  Note that where Ring uses InputStreams for the request and response bodies,
  the clj-http uses ByteArrays for the bodies."
  [{:keys [request-method scheme server-name server-port uri query-string
           headers content-type character-encoding body socket-timeout
           cookies save-request? follow-redirects] :as req}]
  (let [http-url (str (name scheme) "://" server-name
                      (when server-port (str ":" server-port))
                      uri
                      (when query-string (str "?" query-string)))
        request (WebRequest/Create http-url)
        Headers (.Headers request)
        ;^CookieContainer cookie-container (.CookieContainer request)
        cookie-container (CookieContainer.)
        ]
    (when (and content-type character-encoding)
      (set! (.ContentType request) (str content-type
                                        "; charset="
                                        character-encoding)))
    (when (and content-type (not character-encoding))
      (set! (.ContentType request) content-type))
    (doseq [[h v] headers]
      (.Add Headers h v))
    (when (false? follow-redirects)
      (set! (.AllowAutoRedirect request) false))
    (set! (.Method request) (.ToUpper (name request-method)))
    (when socket-timeout
      (set! (.ReadWriteTimeout request) socket-timeout))
    (doseq [cookie (map cookies/map->cookie cookies)]
      (if (empty? (.Domain cookie))
        (set! (.Domain cookie) server-name))
      (.Add cookie-container cookie))
    (set! (.CookieContainer request) cookie-container)
    (when body
      (with-open [out (.GetRequestStream request)]
        (io/copy body out)))
    (try
      (let [
            response (.GetResponse request)
            ]
        (merge {:headers (parse-headers response)
                :status (-> response .StatusCode int)
                :body (when-not (= request-method :head)
                        (coerce-body-entity req response))
                :cookies (into {} (map cookies/cookie->map (.Cookies response)))
                }
               (when save-request?
                 {:request (-> req
                               (dissoc :save-request?)
                               (assoc :http-url http-url))})))
      (catch Exception e
        {:status 500
         :body (str e)}))))
