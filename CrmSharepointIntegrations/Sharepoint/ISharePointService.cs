using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrmSharepointIntegrations
{
    public interface ISharePointService
    {
    
        SharePointOutput CreateListItem(string listName, Dictionary<string, string> fields);

        SharePointOutput UpdateList(string listName, string id, Dictionary<string, string> fields);

        SharePointOutput GetLIFullName(string listName);

    }
}
