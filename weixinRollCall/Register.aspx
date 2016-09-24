<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Register.aspx.cs" Inherits="weixinRollCall.Register" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-CN">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <title>信息学院点名系统</title>
    <link href="css/bootstrap.min.css" rel="stylesheet"/>
    <link href="css/signin.css" rel="stylesheet" />
    <link href="css/main.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" >
         <nav class="navbar navbar-default navbar-static-top navbar-inverse navbar-93E0FF">
      <div class="container ">
        <div class="navbar-header">
          <a class="navbar-brand" href="#">
            <img alt="Brand" src="img/logo.png" height='24px;'/>
          </a>
        </div>
      </div>
    </nav>
        </form>
    <div class="container">
      <form class="form-signin" runat="server">
        <h3 class="form-signin-heading">初次使用，请登录。</h3>
        <label for="inputID" class="sr-only">工号</label>
        <input type="text" id="inputID" class="form-control" placeholder="工号" required="" autofocus="" name="inputID"/>
        <label for="inputPassword" class="sr-only">密码</label>
        <input type="password" id="inputPassword" class="form-control" placeholder="密码" required="" name="inputPassword"/>
        <asp:Button ID="Button1" runat="server" Text="登录" CssClass="btn btn-lg btn-info btn-block" OnClick="Button1_Click"/>
          <div><p>
              <br />
              1.本系统供浙江工业大学教师使用<br />
              2.用户名和密码均为校原创系统的用户名<br />
              3.如遇问题，可联系褚衍清老师（13588243425/693425）,cyq@zjut.edu.cn
               </p></div>
      </form>
    </div>
</body>
</html>
