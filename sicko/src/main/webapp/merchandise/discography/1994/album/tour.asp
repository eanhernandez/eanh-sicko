<%@ LANGUAGE="VBScript"%>
<% OPTION EXPLICIT %>
<%
Function BuildDropDown()
	
	dim	rs
	dim strSQL
	dim strFunctionResult

	set rs = CreateObject("ADODB.Recordset")
	strSQL = " SELECT PhotoId, StoryTitle FROM tblTourStory"
	rs.Open strSQL,"DATABASE=sicko;UID=Jrubin;PWD=k8M8G4f2s;DSN=sicko"
	
	strFunctionResult = strFunctionResult & "<SELECT NAME=""PHOTO""  LANGUAGE=""JAVASCRIPT"" onChange=""SubmitIt()"" >"

	Do Until rs.Eof

		strFunctionResult = strFunctionResult & "<OPTION VALUE=""" & rs("PhotoId") & """>" & rs("StoryTitle") & "&nbsp;&nbsp;&nbsp;&nbsp;</option>" & vbcrlf

		rs.movenext 

	Loop	
	

	rs.Close

	strFunctionResult = strFunctionResult & "</SELECT>"




	BuildDropDown =  strFunctionResult


End Function 
'************************************************************
%>
<HTML>
<HEAD>
<TITLE>Sicko: Pictures</TITLE>

<SCRIPT LANGUAGE="JAVASCRIPT">
<!--
function SubmitIt(){

	self.document.frmTourStory.submit();

}
//-->
</SCRIPT>
</HEAD>

<BODY BGCOLOR="#FFFFFF" LINK="#999933" VLINK="#999933" ALINK="#000000">
<CENTER>
<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0 WIDTH=100%>
<TR>
<TD WIDTH=20><BR></TD>
<TD ALIGN=CENTER>
<BR><IMG SRC="../../images/h_tour.gif" WIDTH=245 HEIGHT=100 ALT="rock and roll tour 1996">
<BR><BR>

<FORM NAME="frmTourStory" METHOD="POST" ACTION="DisplayStory.Asp">

<%=BuildDropDown%>

<!--<input type="IMAGE" SRC="../../images/gobutton.gif" BORDER="0">-->
</form>

</TD>
</TR>
<TR><TD HEIGHT=20 COLSPAN=2><BR></TD></TR>
<TR>
<TD WIDTH=20><BR></TD>
<TD ALIGN=CENTER>
	<TABLE WIDTH=380 CELLSPACING=1 CELLPADDING=0 BORDER=0>
	<TR>
	<TD WIDTH=20 ALIGN=RIGHT VALIGN=TOP><IMG SRC="../../images/bullet6.gif" WIDTH=9 HEIGHT=19 ALT="*"></TD>
	<TD WIDTH=360><FONT FACE="Arial" SIZE=2>On stage at the Holleywood Alley in <A HREF="more/mesa.htm">Mesa, AZ</A></FONT>
	</TD>
	</TR>
	<TR>
	<TD ALIGN=RIGHT VALIGN=TOP><BR><IMG SRC="../../images/bullet6.gif" WIDTH=9 HEIGHT=19 ALT="*"></TD>
	<TD><BR><FONT FACE="Arial" SIZE=2>Denny sitting on a speaker <A HREF="more/pnutbutr.htm">playing his bass</A></FONT></TD>
	</TR>
	<TR>
	<TD ALIGN=RIGHT VALIGN=TOP><BR><IMG SRC="../../images/bullet6.gif" WIDTH=9 HEIGHT=19 ALT="*"></TD>
	<TD><BR><FONT FACE="Arial" SIZE=2>Denny, Josh and Ean waiting for tacos in <A HREF="more/mexico.htm">Mexico</A></FONT></TD>
	</TR>
	<TR>
	<TD ALIGN=RIGHT VALIGN=TOP><BR><IMG SRC="../../images/bullet6.gif" WIDTH=9 HEIGHT=19 ALT="*"></TD>
	<TD><BR><FONT FACE="Arial" SIZE=2>Josh at The Alamo in front of the <A HREF="more/shrub.htm">Texas shrubbery</A></FONT></TD>
        </TR>
	<TR>
	<TD ALIGN=RIGHT VALIGN=TOP><BR><IMG SRC="../../images/bullet6.gif" WIDTH=9 HEIGHT=19 ALT="*"></TD>
	<TD><BR><FONT FACE="Arial" SIZE=2>Picture in front of the <A HREF="more/smith-1.htm"> Smithsonian Air and Space Museum</A></FONT></TD>    
	</TR>
	<TR>
	<TD ALIGN=RIGHT VALIGN=TOP><BR><IMG SRC="../../images/bullet6.gif" WIDTH=9 HEIGHT=19 ALT="*"></TD>
	<TD><BR><FONT FACE="Arial" SIZE=2>Steve (from Tinkle) and Ean with a <A HREF="more/dc-ean.htm"> package of pickle powder</A></FONT></TD>   
	</TR>
	<TR>
	<TD  ALIGN=RIGHT VALIGN=TOP><BR><IMG SRC="../../images/bullet6.gif" WIDTH=9 HEIGHT=19 ALT="*"></TD>   
	<TD><BR><FONT FACE="Arial" SIZE=2><A HREF="more/dc-couch.htm"> Ean, Mr. Benson, and Denny</A> watching Starblazers</FONT></TD>
	</TR>
	</TABLE>
<BR>
<A HREF="home.htm"><IMG SRC="../../images/back.gif" WIDTH=157 HEIGHT=15 ALT="back to picture index" BORDER=0></A>
</TD>
</TR>
</TABLE>
</CENTER>
</BODY>
</HTML>
