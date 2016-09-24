<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="RollCall.aspx.cs" Inherits="weixinRollCall.RollCall" %>
<%@  Import   Namespace="weixinRollCall.DAL.Model"   %>

<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>学生名单</title>
    <link href="css/bootstrap.min.css" rel="stylesheet">
    <link href="css/main.css" rel="stylesheet" />
    <link href="css/RollCall.css" rel="stylesheet"/>
    <link href="css/bootstrap-switch.min.css" rel="stylesheet" />
    <script src="http://cdn.bootcss.com/jquery/2.2.1/jquery.min.js""></script>
    <script src="js/bootstrap.min.js"></script>       
    <script src="js/bootstrap-switch.min.js"></script>
</head>
<body>
    <form runat="server">
    <nav class="navbar navbar-default navbar-fixed-top navbar-inverse navbar-93E0FF">
        <div class="container">
            <div class="row">
                <div class="col-xs-9" style="padding:0px">
        <a class="navbar-brand" >
            <img alt="Brand" src="img/logo.png" height='24px;'/>
          </a>
                </div>
                                  <div class="col-xs-3 none-padding3">
                        <asp:button runat="server" id="myButton"  text="提交" class="btn btn-success" autocomplete="off" OnClick="Button1_Click"/>                                                
                    </div>
            </div>                  
            </div>
            </nav>
    
        <asp:Repeater ID="Repeater1" runat="server">

         <HeaderTemplate>
    <div class="list-group">
              <div  class="list-group-item list-group-top"> 
                    <h3 class="list-group-item-heading "><%=ClassName %></h3>
                  <p class="list-group-item-text" ><span class="glyphicon glyphicon-ok"id="tips"></span> 已到 </a><span class="glyphicon glyphicon-remove"id="tips"></span> 未到 <span class="glyphicon glyphicon-user"id="tips"></span> 请假<br />点名完成后,按右上角绿色按钮即可提交</p>
                </div>
        </HeaderTemplate>
            <ItemTemplate>
            <div class="list-group-item list-group-item-custom">
                <div class="container">
                    <div class="row">
                        <div class="col-xs-6">
                            <h4 class="list-group-item-heading"><strong><%#((ClassStudent)Container.DataItem).StudentName %> </strong></h4><span class="list-group-item-text class"><%#((ClassStudent)Container.DataItem).StudentClass %></span>
                            <p class="list-group-item-text"><%#((ClassStudent)Container.DataItem).StudentID%></p>
                        </div>
                        <div class="col-xs-6 none-padding">

                            <div class="btn-group" data-toggle="buttons" runat="server">
                                <label class="btn btn-3498DB" id="option1">
                                    <input type="radio" name="<%#((ClassStudent)Container.DataItem).StudentID%>"  value="已到"><span class="glyphicon glyphicon-ok"></span> 
                                </label>
                                <label class="btn btn-3498DB" id="option2">
                                    <input type="radio" name="<%#((ClassStudent)Container.DataItem).StudentID%>" value="未到"><span class="glyphicon glyphicon-remove"></span> 
                                </label>
                                <label class="btn btn-3498DB" id="option3">
                                    <input type="radio" name="<%#((ClassStudent)Container.DataItem).StudentID%>" value="请假"><span class="glyphicon glyphicon-user"></span>
                                </label>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </ItemTemplate>
            <FooterTemplate>
                </div>
                </FooterTemplate>
            </asp:Repeater>

        </form>
</body>
        <script>
            document.body.addEventListener('touchstart', function () { }); 
    </script>
</html>        


