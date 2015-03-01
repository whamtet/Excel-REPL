(ns clr-http.lite.cookies
  (:require
   [clojure.string :as string]
   [clr-http.lite.util :as util]
   )
  (:import
   System.Net.Cookie
   System.Net.CookieCollection
   ))

(defn compact-map
  "Removes all map entries where value is nil."
  [m]
  (reduce #(if (get m %2) (assoc %1 %2 (get m %2)) %1)
          (sorted-map) (sort (keys m))))

(defn cookie->map
  "Converts a ClientCookie object into a tuple where the first item is
  the name of the cookie and the second item the content of the
  cookie."
  [cookie]
  [(.Name cookie)
   (compact-map
    {:comment (.Comment cookie)
     :comment-url (if (.CommentUri cookie) (str (.CommentUri cookie)))
     :discard (.Discard cookie)
     :domain (.Domain cookie)
     :expires (let [expires (.Expires cookie)]
                (if-not (= (DateTime. 0) expires) expires))
     :path (.Path cookie)
     :ports (let [ports
                  (filter identity (map #(try (Int32/Parse %) (catch Exception e))
                                        (-> cookie .Port (string/replace "\"" "") (string/split #","))))]
              (if-not (empty? ports) ports))
     :secure (.Secure cookie)
     :value (try
              (util/url-decode (.Value cookie))
              (catch Exception _ (.Value cookie)))
     :version (.Version cookie)})])

(defn map->cookie
  [[cookie-name value]]
  (if (map? value)
    (let [
          {:keys [value comment comment-url discard domain expires path ports secure version]} value
          cookie
          (util/doto-set
           (Cookie. (name cookie-name) (-> value name util/url-encode))
           (.Comment comment)
           (.Discard (if (nil? discard) true discard))
           (.Domain domain)
           (.Path path)
           (.Secure (boolean secure))
           (.Version (or version 0))
           )]
      (if comment-url (set! (.CommentUri cookie) (Uri. comment-url)))
      (if ports (set! (.Port cookie) (->> ports (interpose ",") (apply str) pr-str)))
      (if expires (set! (.Discard cookie) expires))
      cookie)
    (Cookie. (name cookie-name) (-> value name util/url-encode))))

#_(defn decode-cookie
    "Decode the Set-Cookie string into a cookie seq."
    [set-cookie-str]
    (if-not (string/blank? set-cookie-str)
      ;; I just want to parse a cookie without providing origin. How?
      (let [domain (string/lower-case (str (gensym)))
            origin (CookieOrigin. domain 80 "/" false)
            [cookie-name cookie-content] (-> (cookie-spec)
                                             (.parse (BasicHeader.
                                                      "set-cookie"
                                                      set-cookie-str)
                                                     origin)
                                             first
                                             to-cookie)]
        [cookie-name
         (if (= domain (:domain cookie-content))
           (dissoc cookie-content :domain) cookie-content)])))

#_(defn decode-cookies
    "Converts a cookie string or seq of strings into a cookie map."
    [cookies]
    (reduce #(assoc %1 (first %2) (second %2)) {}
            (map decode-cookie (if (sequential? cookies) cookies [cookies]))))

#_(defn decode-cookie-header
    "Decode the Set-Cookie header into the cookies key."
    [response]
    (if-let [cookies (get (:headers response) "set-cookie")]
      (assoc response
        :cookies (decode-cookies cookies)
        :headers (dissoc (:headers response) "set-cookie"))
      response))

#_(defn cookie-response
    "Adds cookies map to server response"
    [response]
    (if-let [^CookieCollection cookies (.Cookies response)]
      (assoc response :cookies (seq cookies))
      response))

#_(defn encode-cookie
    "Encode the cookie into a string used by the Cookie header."
    [cookie]
    (when-let [header (-> (cookie-spec)
                          (.formatCookies [(to-basic-client-cookie cookie)])
                          first)]
      (.getValue ^org.apache.http.Header header)))

#_(defn encode-cookies
    "Encode the cookie map into a string."
    [cookie-map] (string/join ";" (map encode-cookie (seq cookie-map))))

#_(defn encode-cookie-header
    "Encode the :cookies key of the request into a Cookie header."
    [request]
    (if (:cookies request)
      (-> request
          (assoc-in [:headers "Cookie"] (encode-cookies (:cookies request)))
          (dissoc :cookies))
      request))

#_(defn cookie-request
    "Adds cookies to request based on request map"
    [request]
    ;have to look into this one
    )

#_(defn wrap-cookies
    [client]
    #_(fn [request]
        (let [response (client (encode-cookie-header request))]
          (decode-cookie-header response))))
