using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavigionStockUpdate
{

  public class Responce
  {
    [JsonProperty("__invalid_name__odata.metadata")]
    public string metadata { get; set; }
    public List<Value> value { get; set; }
    [JsonProperty("__invalid_name__odata.nextLink")]
    public string nextLink { get; set; }
  }
  public class Value
  {
    public string Item_No { get; set; }
    public string Location_Code { get; set; }
    public string Company_Name { get; set; }
    public string Old_Item_No { get; set; }
    public string Description { get; set; }
    public string Description2 { get; set; }
    public string Vendor_No { get; set; }
    public string Vendor_Item_No { get; set; }
    public string Available_Stock { get; set; }
    public DateTime Stock_Date { get; set; }
    public string Division { get; set; }
    public string Item_Category { get; set; }
    public string Product_Group { get; set; }
    public string Item_Class_Code { get; set; }
    public string Old_Division { get; set; }
    public string Old_Item_Category { get; set; }
    public string Old_Product_Group { get; set; }
    public string Qty_on_Open_Sales_Order { get; set; }
    public string Qty_on_Paid_Sales_Order { get; set; }
    public string Net_Inventory { get; set; }
    public string Brand { get; set; }
    public string ETag { get; set; }
  }


}
