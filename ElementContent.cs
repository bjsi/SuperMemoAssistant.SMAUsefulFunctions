using HtmlAgilityPack;
using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows;

namespace SMAUsefulFunctions
{
    public static class ElementContent
    {

        // TODO: Consider converting to ext. method on HtmlCtrl
        /// <summary>
        /// Get the selection object representing the currently highlighted text in SM.
        /// </summary>
        /// <returns>IHTMLTxtRange object or null</returns>
        public static IHTMLTxtRange GetSelectionObject()
        {

            var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
            var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
            var htmlDoc = htmlCtrl?.GetDocument();
            var sel = htmlDoc?.selection;

            if (!(sel?.createRange() is IHTMLTxtRange textSel))
                return null;

            return textSel;

        }

        // TODO: Consider converting to ext. method on HtmlCtrl
        /// <summary>
        /// Get the selection object representing the currently highlighted control in SM.
        /// </summary>
        /// <returns>IHTMLControlRange object or null</returns>
        public static IHTMLControlRange GetControlSelectionObject()
        {

            var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
            var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
            var htmlDoc = htmlCtrl?.GetDocument();
            var sel = htmlDoc?.selection;

            if (!(sel?.createRange() is IHTMLControlRange ctrlSel))
                return null;

            return ctrlSel;
        }


        /// <summary>
        /// Get the IHTMLDocument2 object representing the first html control of the current element.
        /// </summary>
        /// <returns>IHTMLDocument2 object or null</returns>
        public static IHTMLDocument2 GetFirstHtmlDocument()
        {

            var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
            var htmlCtrl = ctrlGroup?.GetFirstHtmlControl()?.AsHtml();
            return htmlCtrl?.GetDocument();

        }

        /// <summary>
        /// Get the IHTMLDocument2 object representing the focused html control of the current element.
        /// </summary>
        /// <returns>IHTMLDocument2 object or null</returns>
        public static IHTMLDocument2 GetFocusedHtmlDocument()
        {

            var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
            var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
            return htmlCtrl?.GetDocument();

        }

        /// <summary>
        /// Get the IHTMLWindow2 object for the currently focused HtmlControl
        /// </summary>
        /// <returns>IHTMLWindow2 object or null</returns>
        public static IHTMLWindow2 GetFocusedHtmlWindow()
        {

            var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
            var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
            var htmlDoc = htmlCtrl?.GetDocument();
            if (htmlDoc == null)
                return null;

            return Application.Current.Dispatcher.Invoke(() =>
            {
                return htmlDoc.parentWindow;
            });

        }
    }
}
