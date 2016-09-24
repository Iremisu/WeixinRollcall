<%@ Page Language="C#" AutoEventWireup="True" CodeBehind="SelectClass.aspx.cs" Inherits="weixinRollCall.SelectClass" %>
<%@  Import   Namespace="weixinRollCall.DAL.Model"   %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <title>课程列表</title>
    <link href="css/bootstrap.min.css" rel="stylesheet"/>
    <link href="css/main.css" rel="stylesheet" />
    <link href="css/SelectClass.css" rel="stylesheet"/>  
</head>
<body>
     <form id="form1" runat="server">
    <nav class="navbar navbar-default navbar-fixed-top navbar-inverse navbar-93E0FF">
        <div class="container">
            <div class="row">
                <div class="col-xs-9" style="padding:0px">
        <a class="navbar-brand" >
            <img alt="Brand" src="img/logo.png" height='24px;'/>
          </a>
                </div>
                                  <div class="col-xs-3 none-padding3">
                                      <asp:Button ID="Button1" runat="server" Text="切换账号" class="btn btn-success" autocomplete="off" OnClick="Button1_Click" style="margin-left:-5px"/>
                    </div>
            </div>                  
            </div>
            </nav>
      
        <asp:Repeater ID="Repeater1" runat="server">

         <HeaderTemplate>
             <div class="list-group">
                <div  class="list-group-item list-group-1ABC9C"> 
                    <h3 class="list-group-item-heading " style="color:white"><%=TeacherName %>老师</h3>
                    <p class="list-group-item-text">点击下载按钮可将点名的数据下载至校内邮箱，耗时较长，请耐心等待。</p>
                </div>
         </HeaderTemplate>

            <ItemTemplate>           
                <div class="list-group-item list-group-34495E"> 
                    <div class="container">
                        <div class="row"> 
                            <a href='RollCall.aspx?ClassID=<%#((TeacherClass)Container.DataItem).ClassID %>&LessonNum=<%#((TeacherClass)Container.DataItem).LessonNum %>&ClassName=<%#((TeacherClass)Container.DataItem).ClassName%>' class="col-xs-9"style="padding-left:0px">                
                    <h4 class="list-group-item-heading fontwhite"><strong><%#((TeacherClass)Container.DataItem).ClassName %></strong></h4>                                       
                                </a>
                            <div class="col-xs-3">
                                <asp:Button id="submit" class="btn btn-warning download"  runat="server" Text="下载" OnCommand="Download" CommandArgument="<%#((TeacherClass)Container.DataItem).ClassID %>" CommandName="<%#((TeacherClass)Container.DataItem).ClassName%>"/>                                
                            </div>
                            </div>
                        </div>
                    <div class="list-group-item-text p-6285a8" onclick="location.href='RollCall.aspx?ClassID=<%#((TeacherClass)Container.DataItem).ClassID %>&LessonNum=<%#((TeacherClass)Container.DataItem).LessonNum %>&ClassName=<%#((TeacherClass)Container.DataItem).ClassName%>'">
                    <p>
                        <%#((TeacherClass)Container.DataItem).TimeAndLocation%><br /><%#((TeacherClass)Container.DataItem).ClassHead%>
                    </p>
                     </div>
                </div>
            </ItemTemplate>

            <FooterTemplate>
                </div>
            </FooterTemplate>

        </asp:Repeater>

    </form>

</body>
</html>
