using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                if (Session[Util.Constants.USER] != null)
                {
                    Entities.User user = (Entities.User)Session[Util.Constants.USER];
                    lblUserName.Text = user.UserName;
                    imgUserPhoto.ImageUrl = user.Photo;
                }
                else
                {
                    Response.Redirect("Login.aspx");
                    return;
                }

                LoadSongs();
            }
        }

        private void LoadSongs()
        {
            DataTable dtSongs = new BLL.Song(strConnString).Read();
            gvSongs.DataSource = dtSongs;
            gvSongs.DataBind();
        }
    }
}
