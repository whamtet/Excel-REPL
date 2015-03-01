(ns clr-http.lite.util
  "Helper functions for the HTTP client."
  (:require [clojure.clr.io :as io])
  (:import
   System.Text.Encoding
   System.IO.MemoryStream
   System.IO.Stream
   System.IO.Compression.GZipStream
   System.IO.Compression.CompressionMode
   System.IO.Compression.DeflateStream
   ))

(assembly-load-with-partial-name "System.Web")
(import System.Web.HttpUtility)

(defn utf8-bytes
  "Returns the UTF-8 bytes corresponding to the given string."
  [^String s]
  (.GetBytes Encoding/UTF8 s))

(defn utf8-string
  "Returns the String corresponding to the UTF-8 decoding of the given bytes."
  [b]
  (if (string? b)
    b
    (.GetString Encoding/UTF8 b)))

(defn url-decode
  "Returns the form-url-decoded version of the given string, using either a
  specified encoding or UTF-8 by default."
  [encoded & [encoding]]
  (HttpUtility/UrlDecode encoded (or encoding Encoding/UTF8)))

(defn url-encode
  "Returns an UTF-8 URL encoded version of the given string."
  [unencoded]
  (HttpUtility/UrlEncode unencoded Encoding/UTF8))

(defn base64-encode
  "Encode an array of bytes into a base64 encoded string."
  [unencoded]
  (Convert/ToBase64String unencoded))

(defn to-byte-array
  "Returns a byte array for the InputStream provided."
  [is]
  (with-open [os (MemoryStream.)]
    (io/copy is os)
    (.ToArray os)))

(defn gunzip
  "Returns a gunzip'd version of the given byte array."
  [b]
  (when b
    (if (instance? Stream b)
      (GZipStream. b CompressionMode/Decompress)
      (with-open [
                  is (GZipStream. (MemoryStream. b) CompressionMode/Decompress)
                  ]
        (to-byte-array is)))))

(defn gzip
  "gzips binary array"
  [b]
  (when b
  (with-open [compressIntoMs (MemoryStream.)]
    (with-open [gzs (GZipStream. compressIntoMs CompressionMode/Compress)]
      (.Write gzs b 0 (.Length b)))
    (.ToArray compressIntoMs))))

(defn inflate
  "Returns a zlip inflated version of a the given byte array."
  [b]
  (when b
    (with-open [
                is (DeflateStream. (MemoryStream. b) CompressionMode/Decompress)
                ]
      (to-byte-array is))))

(defn deflate
  "Deflates binary array"
  [b]
  (when b
    (with-open [
                mem-stream (MemoryStream.)
                ]
      (with-open [deflate-stream (DeflateStream. mem-stream CompressionMode/Compress)]
        (.Write deflate-stream b 0 (.Length b)))
      (.ToArray mem-stream))))

(defmacro doto-set
  "Similar to doto however sets Csharp properties instead"
  [new & rest]
  (let [x (gensym)]
    `(let [~x ~new]
       ~@(for [[a b] rest]
           `(set! (~a ~x) ~b))
       ~x)))
