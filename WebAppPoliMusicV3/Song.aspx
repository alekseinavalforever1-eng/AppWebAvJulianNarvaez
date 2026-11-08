<%@ Page Title="Canciones" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV3.Song" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <style type="text/css">
        .header-bar {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-top: 15px;
            margin-bottom: 10px;
        }
        .user-photo {
            width: 36px;
            height: 36px;
            border-radius: 50%;
            object-fit: cover;
            vertical-align: middle;
        }
        .logout-link {
            display: inline-block;
            margin-bottom: 20px;
            color: #0056b3;
            text-decoration: underline;
            font-weight: 500;
        }
        .grid-songs {
            border-collapse: collapse;
            background-color: #fff;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            width: 100%;
        }
        .grid-songs th {
            background-color: #2c3e50;
            color: #ffffff;
            padding: 10px 15px;
            text-align: center;
            border: 1px solid #ccc;
        }
        .grid-songs td {
            padding: 8px 15px;
            border: 1px solid #ccc;
            vertical-align: middle;
        }
        .grid-songs tr:nth-child(even) {
            background-color: #f9f9f9;
        }
    </style>

    <div class="header-bar">
        <h2>Welcome <asp:Label ID="lblUserName" runat="server"></asp:Label></h2>
        <asp:Image ID="imgUserPhoto" runat="server" CssClass="user-photo" />
    </div>
    <div>
        <asp:HyperLink ID="hlLogout" runat="server" NavigateUrl="~/Logout.aspx" CssClass="logout-link">Logout</asp:HyperLink>
    </div>
    <h3>Song list:</h3>
    <div>
        <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="False" CssClass="grid-songs">
            <Columns>
                <asp:BoundField DataField="ID_SONG" HeaderText="ID Canción" />
                <asp:BoundField DataField="SONG_NAME" HeaderText="Nombre" />
                <asp:BoundField DataField="SONG_PATH" HeaderText="Ruta" />
                <asp:BoundField DataField="PLAYS" HeaderText="Reproducciones" />
                <asp:TemplateField HeaderText="Reproducir">
                    <ItemTemplate>
                        <audio controls="controls">
                            <source src='<%# ResolveUrl(Eval("SONG_PATH").ToString().Replace("../", "~/")) %>' type="audio/mpeg" />
                        </audio>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    </div>

    <script type="text/javascript">
        document.addEventListener('play', function (e) {
            var audios = document.getElementsByTagName('audio');
            for (var i = 0, len = audios.length; i < len; i++) {
                if (audios[i] !== e.target) {
                    audios[i].pause();
                }
            }
        }, true);
    </script>
</asp:Content>
