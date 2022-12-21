using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WSCalendarProvider.Models
{
    public class PairValue
    {
        [JsonProperty("Key")]
        public string Key { get; set; }

        [JsonProperty("Value")]
        public string Value { get; set; }
    }
}