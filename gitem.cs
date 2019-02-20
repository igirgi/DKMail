using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DKMail
{
    class gitem
    {
        private string _Mail;
        private string _Display;
        private string _Group;

        public gitem(string Group, string Display, string Mail)
        {
            this._Display = Display;
            this._Group = Group;
            this._Mail = Mail;
        }
        public string Group
        {
            get
            {
                return _Group;
            }
        }
        public string Display
        {
            get
            {
                return _Display;
            }
        }
        public string Mail
        {
            get
            {
                return _Mail;
            }
        }
    }
}
