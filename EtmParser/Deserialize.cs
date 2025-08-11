using System.Text.Json.Serialization;

namespace EtmParser
{
    public class AuthorizationResponse
    {
        [JsonPropertyName("data")]
        public AuthorizationData Data { get; set; }
    }

    public class AuthorizationData
    {
        [JsonPropertyName("session")]
        public string Session { get; set; }
    }

    public class Good
    {
        [JsonPropertyName("data")]
        public GoodsData Data { get; set; }
    }

    public class GoodsData
    {
        [JsonPropertyName("gdsPacks")]
        public List<PackData> PackData { get; set; }

        [JsonPropertyName("gdsCode")]
        public string Code { get; set; }

        [JsonPropertyName("gdsNameTitle")]
        public string GoodName { get; set; }
    }

    public class PackData
    {
        [JsonPropertyName("gdsPackWeightVal")]
        public string Weight { get; set; }

        [JsonPropertyName("gdsPackLengthVal")]
        public string Length { get; set; }

        [JsonPropertyName("gdsPackWidthVal")]
        public string Width { get; set; }

        [JsonPropertyName("gdsPackHeigthVal")]
        public string Height { get; set; }
    }


    public class Remain
    {
        [JsonPropertyName("data")]
        public RemainData Data { get; set; }
    }

    public class RemainData
    {
        [JsonPropertyName("InfoStores")]
        public List<InfoStore> InfoStores { get; set; }
    }

    public class InfoStore
    {
        [JsonPropertyName("StoreCode")]
        public int StoreCode { get; set; }

        [JsonPropertyName("StoreQuantRem")]
        public int Remain { get; set; }
    }
}
