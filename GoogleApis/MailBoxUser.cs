using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleApis
{
    class MailBoxUser
    {
        public string PrimaryEmail { get; set; }

        public List<string> NonEditableAliases { get; set; }

        public List<string> Aliases { get; set; }

        public List<string> Emails { get; set; }
    }
}
