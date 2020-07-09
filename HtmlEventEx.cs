﻿using mshtml;
using SuperMemoAssistant.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;

namespace SMAUsefulFunctions
{
    public static class HtmlEventEx
    {
        public enum EventType
        {

            onkeydown,
            onclick,
            ondblclick,
            onkeypress,
            onkeyup,
            onmousedown,
            onmousemove,
            onmouseout,
            onmouseover,
            onmouseup,
            onselectstart,
            onbeforecopy,
            onbeforecut,
            onbeforepaste,
            oncontextmenu,
            oncopy,
            oncut,
            ondrag,
            ondragstart,
            ondragend,
            ondragenter,
            ondragleave,
            ondragover,
            ondrop,
            onfocus,
            onlosecapture,
            onpaste,
            onpropertychange,
            onreadystatechange,
            onresize,
            onactivate,
            onbeforedeactivate,
            oncontrolselect,
            ondeactivate,
            onmouseenter,
            onmouseleave,
            onmove,
            onmoveend,
            onmovestart,
            onpage,
            onresizeend,
            onresizestart,
            onfocusin,
            onfocusout,
            onmousewheel,
            onbeforeeditfocus,
            onafterupdate,
            onbeforeupdate,
            ondataavailable,
            ondatasetchanged,
            ondatasetcomplete,
            onerrorupdate,
            onfilterchange,
            onhelp,
            onrowenter,
            onrowexit,
            onlayoutcomplete,
            onblur,
            onrowsdelete,
            onrowsinserted,

        }

        public static bool SubscribeTo(this IHTMLElement2 element, EventType eventType, IControlHtmlEvent handlerObj)
        {
            try
            {

                return element.IsNull()
                  ? false
                  : element.attachEvent(eventType.Name(), handlerObj);

            }
            catch (RemotingException) { }
            catch (UnauthorizedAccessException) { }

            return false;
        }

        public static void UnsubscribeFrom(this IHTMLElement2 element, EventType eventType, IControlHtmlEvent handlerObj)
        {
            try
            {
                element?.detachEvent(eventType.Name(), handlerObj);
            }
            catch (RemotingException) { }
            catch (UnauthorizedAccessException) { }
        }

        public static bool SubscribeTo(this IHTMLElement2 element, string @event, object pdisp)
        {

            try
            {

                return element.IsNull()
                  ? false
                  : element.attachEvent(@event, pdisp);

            }
            catch (RemotingException) { }
            catch (UnauthorizedAccessException) { }

            return false;
        }

        public static void UnsubscribeFrom(this IHTMLElement2 element, string @event, object pdisp)
        {
            try
            {
                element?.detachEvent(@event, pdisp);
            }
            catch (RemotingException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
