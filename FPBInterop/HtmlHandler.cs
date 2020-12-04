using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;


namespace FPBInterop {
    [Flags]
    internal enum OrderMetadata {
        None = 0,
        DoNotProcess = 1,
        FileToDateFolder = 2,
        DateChanged = 4,
        DetailsChanged = 8,
        StoreChanged = 16,
        Cancelled = 32,
    }

    internal static class HtmlHandler {

        private static readonly TraceSource Tracer = new TraceSource("FPBInterop.HTMLHandling");

        internal static class Wufoo {
            private static void _WipeHTMLNodes(HtmlNode baseNode) {
                baseNode.DescendantsAndSelf().ToList()
                    .ForEach(n => {
                        switch (n.Name) {
                            case ("span"):
                                if (n.InnerText.Length > 0) { n.ParentNode.RemoveChild(n, false); }
                                else { n.ParentNode.RemoveChild(n, true); };
                                break;
                            case "div":
                            case "abbr":
                            case "ul":
                            case "li":
                            case "a":
                            case "br":
                                n.ParentNode.RemoveChild(n, true);
                                break;
                        }

                        if (n.Attributes.Count > 0) { n.Attributes.RemoveAll(); }
                    });
            }
        }

        internal static class Magento {
            internal static MagentoOrder MagentoBuilder(string HTMLBody) {
                HTMLBody = Regex.Replace(HTMLBody, "\n|\r|\t", "");
                HtmlDocument HTMLDoc = new HtmlDocument();
                HTMLDoc.LoadHtml(HTMLBody);
                //_WipeHTMLNodes(HTMLDoc);

                string shop =
                   HTMLDoc.DocumentNode.SelectSingleNode(XPathInfo.Magento.Franchise).InnerText.Trim(' ');
                string deliveryDate =
                     HTMLDoc.DocumentNode.SelectSingleNode(XPathInfo.Magento.DeliveryDate).InnerText.Trim(' ');
                Tracer.TraceEvent(TraceEventType.Information, 0, $"\tShop: {shop}\n\tDate: {deliveryDate}");

                List<MagentoProduct> products = _ReadMagentoProductTable(
                    HTMLDoc.DocumentNode.SelectSingleNode(XPathInfo.Magento.ProductTable));

                OrderMetadata meta = OrderMetadata.None;
                if (products.Count == 0) {
                    meta |= OrderMetadata.DoNotProcess;
                }
                
                return new MagentoOrder(
                    XmlHandler.GetFranchiseInfo(shop),
                    DateTime.Parse(deliveryDate),
                    meta,
                    products);
            }

            private static List<MagentoProduct> _ReadMagentoProductTable(HtmlNode productTable) {
                Tracer.TraceEvent(TraceEventType.Verbose, 0, $"\tProducts:");
                List<MagentoProduct> products = new List<MagentoProduct>();
                HtmlNodeCollection rows = productTable.SelectNodes(@".//tr");

                for (int i = 1; i < rows.Count - 1; i++) {
                    if (rows[i].ParentNode.Name != "tbody")
                        continue;
                    if (rows[i].Descendants("td").Count() < 2)
                        continue;
                    HtmlNode toOrderCell = rows[i].SelectSingleNode(".//td[2]");
                    if (toOrderCell == null)
                        continue;

                    HtmlNode productCell = rows[i].SelectSingleNode(".//td[1]");

                    string productName;
                    string skuCode;
                    bool ignoreProduct;
                    try {
                        productName = productCell.SelectSingleNode(".//p").InnerText.Trim(' ');
                        skuCode = _SanitiseSku(productCell.SelectSingleNode(".//p[2]").InnerText.Trim(' '));
                        ignoreProduct = toOrderCell.FirstChild.InnerText == "Y";
                    }
                    catch (Exception) {
                        continue;
                    }
                    Tracer.TraceEvent(TraceEventType.Verbose, 0, $"\t\t{productName} - {skuCode}");
                    if (!ignoreProduct) {
                        ProductType type;
                        try {
                            type = XmlHandler.GetProductType(skuCode);
                        }catch (ArgumentException) {
                            type = XmlHandler.DefaultProductType;
                            Tracer.TraceEvent(TraceEventType.Warning, 0, 
                                $"## Product {productName} with SKU code {skuCode} has no ProductType entry,\n" +
                                $"## or does not match any known skuTag\n" +
                                $"## Assigned default type, consider adding to Config.xml");
                        }
                        products.Add(new MagentoProduct(productName, skuCode, type));
                    }
                    else
                        continue;
                }

                return products;
            }

            private static string _SanitiseSku(string sku) {
                if (sku.StartsWith("SKU: "))
                    sku = sku.Remove(0, 5);

                if (sku.Length > 35)
                    return "PHOTOCAKE";
                else  
                    return sku; 
            }
            private static void _ReStyleProductTable(HtmlNode productTableNode, bool saveToFile = false) {
                string tableBGColour = @"background-color:#AA7777";

                HtmlAttribute style;
                if (productTableNode.Attributes.Contains("style")) {
                    style = productTableNode.ChildAttributes("style").First();

                    if (!style.Value.Contains(tableBGColour)) {
                        style.Value = string.Join("; ", style.Value, tableBGColour);

                    }
                }
                else { productTableNode.Attributes.Add("style", tableBGColour); }

                //foreach (HtmlNode node in productTable.Descendants("tr")) {
                //node.Attributes.Add("style", "background-color:#FFFFFF");
                //}
            }
            private static void _WipeHTMLNodes(HtmlNode baseNode) {
                baseNode.DescendantsAndSelf().ToList()
                   .ForEach(n => {
                       if (n.Attributes.Count > 0) { n.Attributes.RemoveAll(); }
                   });
            }
        }
    }

    internal static class XPathInfo {
        internal static class Magento {
            internal const string ProductTable =
                "//div/table/tbody/tr/td/table/tbody/tr[7]/td/span/table";
            internal const string DeliveryTable =
                "//div/table/tbody/tr/td/table/tbody/tr[6]/td/table/tbody/tr";
            internal const string DeliveryDate =
                "//div/table/tbody/tr/td/table/tbody/tr[6]/td/table/tbody/tr/td[2]/strong";
            internal const string Franchise =
                "//div/table/tbody/tr/td/table/tbody/tr[6]/td/table/tbody/tr/td[1]/strong[1]";
        }
        internal static class Wufoo {
            internal static class DecoratedCake {
                internal const string OrderNumber = "//div/table/tbody/tr[1]/td/div";
                internal const string Franchise = "//div/table/tbody/tr[2]/td/div";
                internal const string DeliveryDate = "//div/table/tbody/tr[7]/td/abbr";
                internal const string DeliveryDay = "//div/table/tbody/tr[8]/td/div";
                internal const string Email = "//div/table/tbody/tr[3]/td/a";
            }
        }
    }
}
