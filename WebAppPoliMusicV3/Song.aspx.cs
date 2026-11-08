using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV3
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                if (Session[EntityLayer.PoliMusicV3.Util.Constants.USER] != null)
                {
                    EntityLayer.PoliMusicV3.User user = (EntityLayer.PoliMusicV3.User)Session[EntityLayer.PoliMusicV3.Util.Constants.USER];
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
            DataTable dtSongs = new BLL.PoliMusicV3.Song(strConnString).Read();
            gvSongs.DataSource = dtSongs;
            gvSongs.DataBind();
        }
    }
}
