using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrmSharepointIntegrations
{
    public sealed class UpdateSharepointListItem : WorkFlowActivityBase
    {
        [Input("Title of List Item")]
        public InArgument<string> Title { get; set; }

        [RequiredArgument]
        [Input("List Item Id")]
        public InArgument<string> ListItemId { get; set; }

        [Output("Succesfully Executed")]
        public OutArgument<bool> Success { get; set; }

        [Output("Error Details")]
        public OutArgument<string> ExceptionDetails { get; set; }


        public override void ExecuteCRMWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext crmWorkflowContext)
        {
            // You might want to work on configuring these 3 settings - URL, User Name, Password !!
            ISharePointService spService = new SPService("https://xyz.sharepoint.com", "user@xyz.onmicrosoft.com", "password");

            // These are the are part of Sharepoint List Item. Think of better way to handle
            // the field changes.. !!
            Dictionary<string, string> fields = new Dictionary<string, string>();
            fields.Add("Title", Title.Get(context));

            // This is the list name what we see in the Sharepoint. 
            // You might want to keep it as configurable value !!
            var result = spService.UpdateList("Leads from Flow", ListItemId.Get(context), fields);

            Success.Set(context, result.Success);
            if (!result.Success)
            {
                ExceptionDetails.Set(context, result.Exception.ToString());
            }
        }
    }
}
