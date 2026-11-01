<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>PoliMusic - Song List</title>
    <style type="text/css">
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f4f6f9;
            color: #333;
            margin: 30px;
        }
        .header-bar {
            display: flex;
            align-items: center;
            gap: 10px;
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
</head>
<body>
    <form id="formSong" runat="server">
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
    </form>
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
</body>
</html>
