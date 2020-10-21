using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Diagnostics;
using HtmlAgilityPack;


namespace FPBInterop {
    public class HTMLHandling {
        /// PROPERTIES ///

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
                HtmlNode productTable =
                  HtmlBody.DocumentNode.Descendants().Where(
                      n => n.NodeType == HtmlNodeType.Text
                      && n.InnerText.Trim(' ') == "Quantity").Single().Ancestors("table").First();

                HtmlNodeCollection rows = productTable.SelectNodes(@".//tr");

                for (int i = 1; i < rows.Count - 1; i++) {
                    if (rows[i].Descendants("td").Count() < 2)
                        continue;

                    HtmlNode cellData = rows[i].SelectSingleNode(".//td[2]");
                    if (cellData == null)
                        continue;

                    Trace.WriteLine(cellData.FirstChild.XPath);
                    Trace.WriteLine($"Ignore order: {cellData.FirstChild.InnerText}");
                    return (cellData.FirstChild.InnerText == "N");
                }

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
        }
    }
}
