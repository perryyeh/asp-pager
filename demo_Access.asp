<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="Cls_Page.asp"-->
<%
Dim startime,endtime
startime=timer()

Dim Db,Conn,rs,nav,rc,Page,i
Db = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("db/IP.mdb")
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.open Db
%>
<html>
<head>
<title>叶子ASP分页类-access调用示范</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
div {float:left;width:25%;}
</style>
</head>
<body>
<div>ID</div><div>标题</div><div>内容</div><div>时间</div>
<%
Set Page = new Cls_Page				'创建对象
Set Page.Conn = conn				'得到数据库连接对象
With Page
	.PageSize = 10					'每页记录条数
	.PageParm = "p"					'页参数
	'.PageIndex = 10				'当前页，可选参数，一般是生成静态时需要
	.Database = "ac"				'数据库类型,AC为access,MSSQL为sqlserver2000存储过程版,MYSQL为mysql,PGSQL为PostGreSql
	.Pkey="MID"						'主键
	.Field="MID,ip2,country,city"	'字段
	.Table="dv_address"				'表名
	.Condition=""					'条件,不需要where
	.OrderBy=""						'排序,不需要order by,需要asc或者desc
	.RecordCount = -1				'总记录数，可以外部赋值，0不保存（适合搜索），-1存为session，-2存为cookies，-3存为applacation

	.NumericJump = 5 '数字上下页个数，可选参数，默认为3，负数为跳转个数，0为显示所有
	.Template = "总记录数：{$RecordCount} 总页数：{$PageCount} 每页记录数：{$PageSize} 当前页数：{$PageIndex} {$FirstPage} {$PreviousPage} {$NumericPage} {$NextPage} {$LastPage} {$InputPage} {$SelectPage}" '整体模板，可选参数，有默认值
	.FirstPage = "首页" '可选参数，有默认值
	.PreviousPage = "上一页" '可选参数，有默认值
	.NextPage = "下一页" '可选参数，有默认值
	.LastPage = "尾页" '可选参数，有默认值
	.NumericPage = " {$PageNum} " '数字分页部分模板，可选参数，有默认值
End With

rs = Page.ResultSet() '记录集
'rc = Page.RowCount() '可选，输出总记录数
nav = Page.Nav() '分页样式

If IsNull(rs) Then
	Response.Write "<div >暂无记录</div>"
Else
	For i=0 To Ubound(rs,2)
		Response.Write "<div>"&rs(0,i)&"</div><div>"&rs(1,i)&"</div><div>"&rs(2,i)&"</div><div>"&rs(3,i)&"</div>"
	Next
End If
%>
<br><%Response.Write nav%>
<br><%endtime=timer()%>本页面执行时间：<%=FormatNumber((endtime-startime)*1000,3)%>毫秒
</body>
</html>
<%
Page = Null
Set Page = Nothing
%>