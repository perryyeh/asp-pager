<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="Cls_Pager.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
</head>
<body>
<%
Dim Pager
Set Pager = new Cls_Pager
With Pager
	.PageSize = 9 '每页记录数,可选参数，默认为10
	'.PageCount = 50 '总页数,半必选参数,PageCount和RecordCount必须指定其一，若2者都指定则RecordCount优先
	.RecordCount = 100 '总记录数,半必选参数,PageCount和RecordCount必须指定其一，若2者都指定则RecordCount优先
	.NumericJump = 3 '数字上下页个数，可选参数，默认为3，负数为跳转个数，0为显示所有
	'.PageIndex = Request("P") '当前页数，可选参数，如果为空则按照页参数取值，实际上这个是留给生成静态页面的
	'.PageUrl = "demo_Pager.asp?P={$PageNum}&a=2" '超链接模板，可选参数，同上留给生成静态页面的
	.PageParm = "pa" '页参数，可选参数，默认为'p'，如果PageIndex和PageUrl指定，则此不用指定
	.Template = "<div style='pager'>总记录数：{$RecordCount} 总页数：{$PageCount} 每页记录数：{$PageSize} 当前页数：{$PageIndex} {$FirstPage} {$PreviousPage} {$NumericPage} {$NextPage} {$LastPage} {$InputPage} {$SelectPage}</div>" '整体模板，可选参数，有默认值
	.FirstPage = "首页" '可选参数，有默认值
	.PreviousPage = "上一页" '可选参数，有默认值
	.NextPage = "下一页" '可选参数，有默认值
	.LastPage = "尾页" '可选参数，有默认值
	.NumericPage = " {$PageNum} " '数字分页部分模板，可选参数，有默认值
End With

Response.Write Pager.Nav()
%>
</body>
</html>