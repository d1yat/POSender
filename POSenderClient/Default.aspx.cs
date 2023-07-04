using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using POSenderClient.POSenderServices;
//using POSenderClient.POSenderService4130; // yg ke server Production

namespace POSenderClient
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void BtnClickMe_Click(object sender, EventArgs e)
        {
            string PONumber = string.Empty;
            PONumber = TxtPONumber.Text.Trim();

            try
            {
                Service1 service = new Service1();
                string responseMsg = service.SendPO(PONumber);
                LblPONumber.Text = responseMsg;
            }
            catch (Exception exception)
            {

                LblPONumber.Text = exception.Message;
            }
        }


    }
}
