using HtmlAgilityPack;
using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SMAUsefulFunctions
{
    public static class SelObjEx
    {

        public static string GetLastPartialWord(this IHTMLTxtRange selObj)
        {

            if (selObj == null)
                return null;

            var duplicate = selObj.duplicate();

            bool found = false;
            while (duplicate.moveStart("character", -1) == -1)
            {

                if (string.IsNullOrEmpty(duplicate.text))
                    return null;

                if (char.IsWhiteSpace(selObj.text.First()))
                {
                    found = true;
                    break;
                }

            }

            return selObj.text;

        }
       
        /// <summary>
        /// Get the end index of the selection object as an inner text index.
        /// NOTE: Counts \r\n incorrectly.
        /// </summary>
        /// <param name="selObj"></param>
        /// <returns>index or -1</returns>
        public static int GetSelectionTextEndIdx(this IHTMLTxtRange selObj)
        {

            int MaxTextLength = 2000000000;
            int result = -1;
            if (selObj != null)
            {
                var duplicate = selObj.duplicate();
                result = Math.Abs(duplicate.moveEnd("character", -MaxTextLength));
            }
            return result;

        }

        /// <summary>
        /// Get the start index of the selection object as an inner text index.
        /// NOTE: Counts \r\n incorrectly.
        /// </summary>
        /// <param name="selObj"></param>
        /// <returns>index or -1</returns>
        public static int GetSelectionTextStartIdx(this IHTMLTxtRange selObj)
        {

            int MaxTextLength = 2000000000;
            int result = -1;
            if (selObj != null)
            {
                var duplicate = selObj.duplicate();
                result = Math.Abs(duplicate.moveStart("character", -MaxTextLength));
            }
            return result;

        }

        /// <summary>
        /// Convert a text index to the equivalent position in the html string.
        /// </summary>
        /// <param name="html"></param>
        /// <param name="textIdx"></param>
        /// <returns>index or -1</returns>
        public static int ConvertTextIdxToHtmlIdx(string html, int textIdx)
        {

            if (string.IsNullOrEmpty(html))
                return -1;

            if (textIdx < 0)
                return -1;

            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            var nodes = doc.DocumentNode
                          ?.Descendants()
                          ?.Where(x => x.Name == "#text" || x.Name == "br");

            if (nodes == null || !nodes.Any())
                return -1;

            // Return -1 if not found
            int htmlIdx = -1;

            int startIdx = -1;
            int endIdx = -1;


            // Last pair of characters
            Tuple<char, char> last = new Tuple<char, char>('x', 'x');

            foreach (var node in nodes)
            {

                if (node.Name == "br")
                {
                    startIdx++;
                    endIdx++;
                    continue;
                }

                var decoded = HttpUtility.HtmlDecode(node.InnerText);

                // Consecutive \r\n get turned into a single null object in the IHTMLTxtRange
                // Increment TextIdx for each consecutive \r\n

                if (decoded.Length == 1)
                {
                    last = new Tuple<char, char>(last.Item2, decoded[0]);
                }
                else
                {
                    for (int i = 0, j = 1; j < decoded.Length; i++, j++)
                    {
                        if (decoded[i] == '\r' && decoded[j] == '\n')
                        {
                            if (last.Item1 == '\r' && last.Item2 == '\n')
                            {
                                textIdx++;
                            }
                        }
                        last = new Tuple<char, char>(decoded[i], decoded[j]);
                    }
                }

                var length = decoded.Replace("\r\n", " ").Length;
                endIdx += length;

                if (startIdx <= textIdx && textIdx <= endIdx)
                {
                    if (startIdx == -1)
                        startIdx = 0;

                    var textIdxDiff = (textIdx - startIdx);
                    var htmlIdxDiff = 0;

                    // Add the length of each html encoded character before the target
                    // to the htmlIdx
                    // TODO: This encodes punctuation like ' 
                    // TODO: Check which characters I need to manage

                    //for (int i = 0; i < textIdxDiff; i++)
                    //  htmlIdxDiff += HttpUtility.HtmlEncode(decoded[i].ToString()).Length;

                    htmlIdx = node.InnerStartIndex + textIdxDiff;
                    break;
                }

                // TODO: Testing moving here
                if (endIdx > -1)
                    startIdx = endIdx + 1;
            }

            return htmlIdx;
        }
    }
}
