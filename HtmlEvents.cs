using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SMAUsefulFunctions
{
    public interface IControlHtmlEvent
    {
        event EventHandler<IControlHtmlEventArgs> OnEvent;
        void handler(IHTMLEventObj e);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class HtmlEvent : IControlHtmlEvent
    {
        public event EventHandler<IControlHtmlEventArgs> OnEvent;

        [DispId(0)]
        public void handler(IHTMLEventObj e)
        {
            if (!OnEvent.IsNull())
                OnEvent(this, new IControlHtmlEventArgs(e));
        }
    }

    public class IControlHtmlEventArgs
    {
        public IHTMLEventObj EventObj { get; set; }
        public IControlHtmlEventArgs(IHTMLEventObj EventObj)
        {
            this.EventObj = EventObj;
        }
    }

}
