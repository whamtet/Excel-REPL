using clojure.lang;
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Driver.Builders;
using System;
using System.Text;

public static class DB
{
    private static IFn read_string = clojure.clr.api.Clojure.var("clojure.core", "read-string");
    private static IFn pr_str = clojure.clr.api.Clojure.var("clojure.core", "pr-str");
    public static MongoDB.Driver.MongoClient Connect(String s)
    {
        var client = new MongoDB.Driver.MongoClient(s);
        return client;
    }
    //mongodb://[username:password@]hostname[:port][/[database][?options]]
    public static MongoDB.Driver.MongoClient Connect(String host, int port)
    {
        return new MongoDB.Driver.MongoClient(String.Format("mongodb://{0}:{1}", host, port));
    }
    public static MongoDB.Driver.MongoClient Connect()
    {
        return new MongoDB.Driver.MongoClient();
    }

    public static void Set(MongoDB.Driver.MongoClient c, String k, BsonDocument v)
    {
        var db = c.GetServer().GetDatabase("db");
        var coll = db.GetCollection(k);
        Entity entity = new Entity { data = v, _id = k};
        coll.Save(entity);
    }
    public static Object Get(MongoDB.Driver.MongoClient c, String k)
    {
        var db = c.GetServer().GetDatabase("db");
        var coll = db.GetCollection(k);
        var query = Query<Entity>.EQ(e => e._id, k);
        var entity = coll.FindOne(query);
        var v = entity.GetElement("data");
        //return read_string.invoke(v.Value.ToString());
        return v.Value;
    }
    
    public class Entity
    {
        [BsonId]
        public string _id { get; set; }

        public BsonDocument data { get; set; }
    }
}
