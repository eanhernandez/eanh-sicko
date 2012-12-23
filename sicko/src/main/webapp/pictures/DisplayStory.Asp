<%

'******************************************************************
'DEBUGGING
'******************************************************************
'for each x in request.form
'response.write (x & " = " &request(x) & "<BR>")
'next
'******************************************************************



	
'******************************************************************
'CHECK TO SEE IF THE ARG IS ACTUALLY A NUMBER, AND ALSO IF ONE WAS
'PASSED IN... REDIRECTS TO DROP DOWN IF NOT
'******************************************************************
If ((IsNumeric(Request("Photo")))  AND (Len(Request("Photo"))>0)) Then


	'******************************************************************
	'GETS THE HIGHEST PHOTO ID IN THE DB, AND ALLOWS FORWARD AND BACKWARD
	'BUTTONS BASED ON THIS
	'******************************************************************
	set rs1 = CreateObject("ADODB.Recordset")
	strSQL = "Select TheCount = max(photoid) from tbltourstory "
	rs1.Open strSQL,"DATABASE=sicko;UID=Jrubin;PWD=k8M8G4f2s;DSN=sicko"
	intTheCount = rs1("TheCount")
	rs1.Close

	'response.write(intthecount)


	If (Cint(Request("Photo")) = Cint(intTheCount) ) Then

		boolShowNextButtons = FALSE
		boolShowPreviousButtons = TRUE

	ElseIf (Cint(Request("Photo")) < 2 ) Then

		boolShowPreviousButtons = FALSE
		boolShowNextButtons = TRUE

	Else

			boolShowNextButtons = TRUE
			boolShowPreviousButtons = TRUE

	End If


	'******************************************************************
	'GETS STORY TEXT
	'******************************************************************
	set rs = CreateObject("ADODB.Recordset")
	strSQL = "SELECT StoryText FROM tblTourStory WHERE PhotoId = Convert(Int,'"&Request("PHOTO")&"')"
	rs.Open strSQL,"DATABASE=sicko;UID=Jrubin;PWD=k8M8G4f2s;DSN=sicko",3,3

	If rs.RecordCount = 0 Then

		strStoryText = "<CENTER>No story text found!</CENTER>"

	 Else

		strStoryText = rs("StoryText")

	End If

	rs.Close

	intNextStory = Cint(Request("Photo")) + 1
	intPreviousStory = Cint(Request("Photo")) - 1


Else

	response.redirect("tour.asp")

End If

%>


<HTML>
<HEAD>
<TITLE>Sicko -  Silly story #<%=Request("Photo")%></TITLE>
</HEAD>
<BODY BGCOLOR="#FFFFFF" LINK="#666666" VLINK="#666666" ALINK="#000000">

<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0 WIDTH=100%>
	<TR>
		<TD WIDTH=20 ROWSPAN=2>
			<BR>
		</TD>
		<TD ALIGN=CENTER>
			<BR><NOBR><IMG SRC="../../images/h_story.gif" WIDTH=176 HEIGHT=44 ALT="Sicko Story"><IMG SRC="../../tour/no<%=Request("Photo")%>.gif" ALT="#<%=Request("Photo")%>"></NOBR>
		</TD>
	</TR>
	<TR>
		<TD>
			<FONT FACE="Arial" SIZE=2>
			<P><IMG SRC="../../tour/tp<%=Request("Photo")%>.jpg" ALIGN=LEFT HSPACE=10 VSPACE=10>
			
			<%=strStoryText%>

		</TD>
	</TR>
	<TR>
		<TD COLSPAN=2>
			<BR>
			<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0 WIDTH=100% BGCOLOR="#000000">
				<TR>
					<TD HEIGHT=23>
						<A HREF="../../backintheday/pictures/home.htm">
						<IMG SRC="../../images/p_wayback.gif" WIDTH=115 HEIGHT=23 BORDER=0 ALT="way back">
						</A>
					</TD>
					<TD HEIGHT=23>
						<A HREF="../../backintheday/pictures/tour.asp">
						<IMG SRC="../../images/p_back.gif" WIDTH=69 HEIGHT=23 BORDER=0 ALT="back">
						</A>
					</TD>
					<TD WIDTH=30>
						<BR>
					</TD>
						
				<% If (boolShowPreviousButtons = TRUE) Then %>
				
					<TD HEIGHT=23>
						<FORM METHOD="POST" ACTION="DisplayStory.Asp">
						<INPUT TYPE="hidden" NAME="PHOTO" VALUE="<%=intPreviousStory%>">
						<INPUT TYPE="image" SRC="../../images/p_previous.gif" BORDER=0>
						</FORM>
					</TD>

				<% End If %>

				<% If boolShowNextButtons Then %>

					<TD HEIGHT=23 ALIGN=RIGHT>
						<FORM METHOD="POST" ACTION="DisplayStory.Asp">
						<INPUT TYPE="hidden" NAME="PHOTO" VALUE="<%=intNextStory%>"	>
						<INPUT TYPE="image" SRC="../../images/p_next.gif" BORDER=0>
						</FORM>
					</TD>

				<% End If %>



				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>