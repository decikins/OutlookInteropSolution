using System;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;


namespace FPBInterop {
    public class HTMLHandling {
        /// PROPERTIES ///
        private static readonly TraceSource Tracer = new TraceSource("FPBInterop.HTMLHandling");

        /// METHODS ///
        public static class Wufoo {
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

        public static class Magento {
            

            public static bool ParseOrder(string HTMLBody) {
                HTMLBody = Regex.Replace(HTMLBody, "\n|\r|\t", "");
                HtmlDocument HTMLDoc = new HtmlDocument();
                HTMLDoc.LoadHtml(HTMLBody);

                _WipeHTMLNodes(HTMLDoc);

                return CheckProductTable(HTMLDoc);
            }

            public static bool CheckProductTable(HtmlDocument HtmlBody) {

                HtmlNode productTable = HtmlBody.DocumentNode.SelectSingleNode(XPathInfo.Magento.productTableXPath);
                /*HtmlBody.DocumentNode.Descendants().Where(
                    n => n.NodeType == HtmlNodeType.Text
                    && n.InnerText.Trim(' ') == "Quantity").Single().Ancestors("table").First();*/

                HtmlNodeCollection rows = productTable.SelectNodes(@".//tr");

                StringBuilder traceInfo = new StringBuilder();
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
                    string singlePrice;
                    string quantity;
                    string totalPrice;
                    bool ignoreProduct;
                    try {
                        productName = productCell.SelectSingleNode(".//p").InnerText.Trim(' ');
                        skuCode = productCell.SelectSingleNode(".//p[2]").InnerText.Trim(' ');
                        singlePrice = rows[i].SelectSingleNode(".//td[3]/span").InnerText.Trim(' ');
                        quantity = rows[i].SelectSingleNode(".//td[4]").InnerText.Trim(' ');
                        totalPrice = rows[i].SelectSingleNode(".//td[5]").InnerText.Trim(' ');
                        ignoreProduct = toOrderCell.FirstChild.InnerText == "Y";
                    } catch (Exception) {
                        continue;
                    }

                    if (skuCode.Length > 35)
                        skuCode = $"{skuCode.Substring(0, 35)}...";

                    string deliveryDate =
                        HtmlBody.DocumentNode.SelectSingleNode(XPathInfo.Magento.deliveryDate).InnerText.Trim(' ');
                    string shop =
                        HtmlBody.DocumentNode.SelectSingleNode(XPathInfo.Magento.shop).InnerText.Trim(' ');

                    traceInfo.Append(
                        $"|\tShop:\t{shop};\n"+
                        $"|\tDelivery Date:\t{deliveryDate};\n"+
                        $"|\t{productName}\n" +
                        $"|\t\t{skuCode};\n" +
                        $"|\t\tPrice:\t{singlePrice};\n" +
                        $"|\t\tQuantity:\t{quantity};\n" +
                        $"|\t\tTotal:\t{totalPrice};\n" +
                        $"|\tIgnore product: {ignoreProduct}\n");

                    if (!ignoreProduct) {
                        traceInfo.Append("|Must be processed\n");
                        Tracer.TraceEvent(TraceEventType.Information, 0, traceInfo.ToString());
                        return true;
                    }
                    else
                        continue;
                }
                traceInfo.Append("|Processing not necessary, moving to Deleted Items\n");
                Tracer.TraceEvent(TraceEventType.Information, 0, traceInfo.ToString());
                return false;
            }
            public static void ReStyleProductTable(HtmlNode productTableNode, bool saveToFile = false) {
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

            private static void _WipeHTMLNodes(HtmlDocument doc) {
                doc.DocumentNode.DescendantsAndSelf().ToList()
                   .ForEach(n => {
                       if (n.Attributes.Count > 0) { n.Attributes.RemoveAll(); }
                   });
            }

            public static bool ParseOrderSpecial(string HTMLBody) {
                HTMLBody = Regex.Replace(HTMLBody, "\n|\r|\t", "");
                HtmlDocument HTMLDoc = new HtmlDocument();
                HTMLDoc.LoadHtml(HTMLBody);

                return CheckProductTableSpecial(HTMLDoc);
            }

            public static bool CheckProductTableSpecial(HtmlDocument HtmlBody) {
                HtmlNode productTable = //HtmlBody.DocumentNode.SelectSingleNode(productTableXPath);
                HtmlBody.DocumentNode.Descendants().Where(
                    n => n.NodeType == HtmlNodeType.Text
                    && n.InnerText.Trim(' ') == "Quantity").Single().Ancestors("table").First();

                HtmlNodeCollection rows;
                try {
                    rows = productTable.SelectNodes(@".//tr");
                }
                catch (NullReferenceException) {
                    Trace.WriteLine(" SKIPPED ");
                    return false;
                }

                bool mustProcess = false;
                bool firstProductIgnored = false;
                for (int i = 1; i < rows.Count - 1; i++) {
                    if (rows[i].ParentNode.Name != "tbody")
                        continue;

                    if (rows[i].Descendants("td").Count() < 2)
                        continue;

                    HtmlNode toOrderCell = rows[i].SelectSingleNode(".//td[2]");
                    if (toOrderCell == null)
                        continue;

                    HtmlNode productCell = rows[i].SelectSingleNode(".//td[1]");
                    bool ignoreProduct;
                    try {
                        ignoreProduct = toOrderCell.FirstChild.InnerText == "Y";
                    }
                    catch (Exception) {
                        continue;
                    }
                    if (i == 1 & ignoreProduct == true) {
                        //Trace.WriteLine("\tIgnored First Product: True");
                        firstProductIgnored = true;
                    }

                    if (!ignoreProduct) {
                        // Trace.WriteLine("\tMust Process: True");
                        mustProcess = true;
                        break;
                    }
                    else
                        continue;
                }
                if (mustProcess & firstProductIgnored) {
                    //Trace.WriteLine("\tProbably fuck up, move back to deleted items");
                    return true;
                }
                else
                    Trace.WriteLine(" safe");
                return false;
            }
        }
    }

    internal static class XPathInfo {
        internal static class Magento {
            internal static string productTableXPath =
                    "//div/table/tbody/tr/td/table/tbody/tr[7]/td/span/table";
            internal static string deliveryInfoXpath =
                "//div/table/tbody/tr/td/table/tbody/tr[6]/td/table/tbody/tr";
            internal static string deliveryDate =
                "//div/table/tbody/tr/td/table/tbody/tr[6]/td/table/tbody/tr/td[2]/strong";
            internal static string shop =
                "//div/table/tbody/tr/td/table/tbody/tr[6]/td/table/tbody/tr/td[1]/strong[1]";
        }
    }
}
