using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace ConrepOutlookAddin
{
    public class MailItemHeader
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public class MailItemHeaderConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            List<MailItemHeader> headers = (List<MailItemHeader>) value;

            writer.WriteStartArray();
            foreach (var header in headers)
            {
                writer.WriteStartObject();
                writer.WritePropertyName(header.Key);
                writer.WriteValue(header.Value);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }

        public override bool CanRead { get; } = false;

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(List<MailItemHeader>);
        }
    }
}
