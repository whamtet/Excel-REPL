(ns excel-repl.coerce-db
  (:use [clojure.data.json :only [write-str read-str]]
        )
  (:import [clojure.lang IPersistentMap Keyword IPersistentCollection Ratio Symbol]
           [System.Collections IEnumerable IDictionary IList]
           ;  [java.util Map List Set]
           ;  [com.mongodb DBObject BasicDBObject BasicDBList]
           ;  [com.mongodb.gridfs GridFSFile]
           ;  [com.mongodb.util JSON]
           ))

(assembly-load "MongoDB.Bson")
(import '[MongoDB.Bson BsonArray BsonInt32 BsonInt64 BsonBoolean
          BsonDateTime BsonDouble BsonNull BsonRegularExpression
          BsonString BsonSymbol BsonDocument BsonElement
          BsonValue BsonExtensionMethods
          ])



(defprotocol ConvertibleFromMongo
  (mongo->clojure [o]))

(extend-protocol ConvertibleFromMongo
  BsonDocument
  (mongo->clojure [^BsonDocument m]
                  (into {}
                        (map #(vector (mongo->clojure (.Name %)) (mongo->clojure (.Value %))) (seq m))))
  BsonArray
  (mongo->clojure [^IEnumerable l]
                  (mapv mongo->clojure l))
  Object
  (mongo->clojure [o] o)
  nil
  (mongo->clojure [o] o)
  BsonInt32
  (mongo->clojure [o] (int o))
  BsonInt64
  (mongo->clojure [o] (long o))
  BsonBoolean
  (mongo->clojure [o] (.Value o))
  BsonDateTime
  (mongo->clojure [o] (str o))
  BsonDouble
  (mongo->clojure [o] (.Value o))
  BsonNull
  (mongo->clojure [o])
  BsonRegularExpression
  (mongo->clojure [o] (-> o .Pattern re-pattern))
  String
  (mongo->clojure [o]
                  (let [s (str o)]
                    (if (.StartsWith s ":")
                      (keyword (.Substring s 1))
                      s)))
  BsonString
  (mongo->clojure [o]
                  (let [s (str o)]
                    (if (.StartsWith s ":")
                      (keyword (.Substring s 1))
                      s)))
  BsonSymbol
  (mongo->clojure [o] (-> o str symbol))
  )


;; ;;; Converting data from Clojure into data objects suitable for Mongo

(defprotocol ConvertibleToMongo
  (clojure->mongo [o]))

(extend-protocol ConvertibleToMongo
  IPersistentMap
  (clojure->mongo [m]
                  (let [out (BsonDocument.)]
                    (doseq [[k v] m]
                      (.Add out (str k) (clojure->mongo v)))
                    out))
  IPersistentCollection
  (clojure->mongo [m] (BsonArray. (map clojure->mongo m)))
  Keyword
  (clojure->mongo [^Keyword o]
                  (BsonString. (str o)))
  nil
  (clojure->mongo [o] BsonNull/Value)
  Object
  (clojure->mongo [o] (BsonValue/Create o))
  Ratio
  (clojure->mongo [o] (BsonValue/Create (.ToDouble o nil)))
  Symbol
  (clojure->mongo [o] (-> o str BsonValue/Create))
  Int64
  (clojure->mongo [o] (BsonInt64. o))
  )
