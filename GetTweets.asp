<!--#include file="Libs/ASPTwitter/ASPTwitter.asp"-->
<%
const TWITTER_API_CONSUMER_KEY = ""
const TWITTER_API_CONSUMER_SECRET = ""
const TWITTER_BEARER_TOKEN = ""

TWITTER_ACCOUNTS = Array("","")

KEYWORDS = Array("","")
SENDER_EMAIL = ""
RECIPIENT_EMAILS = ""
EMAIL_SUBJECT = "Tweets Update Email"

if TWITTER_API_CONSUMER_KEY = "" OR TWITTER_API_CONSUMER_SECRET = "" OR TWITTER_BEARER_TOKEN = "" Then
  response.write "Please fill in your Twitter API Keys"
  response.end
End if

if SENDER_EMAIL = "" OR RECIPIENT_EMAILS = "" Then
  response.write "Please fill in your the twitter accounts, keywords and email details before running this script"
  response.end
End if

' Twitter API client.
Dim objASPTwitter
Dim emailBody

' Tweets will be obtained by parsing data from Twitter API.
Dim objTweets

Set objASPTwitter = New ASPTwitter
Call objASPTwitter.Configure(TWITTER_API_CONSUMER_KEY, TWITTER_API_CONSUMER_SECRET)
Call objASPTwitter.ConfigureOAuth(TWITTER_API_OAUTH_TOKEN, TWITTER_API_OAUTH_TOKEN_SECRET)

Call objASPTwitter.Login
objASPTwitter.strBearerToken = TWITTER_BEARER_TOKEN
For Each twitter_account In TWITTER_ACCOUNTS
	Call LoadTweetsUserTimeline(twitter_account)
	Call WriteTweetsUserTimeline
  For Each search_criteria in KEYWORDS
    Call LoadTweetsSearch(twitter_account, search_criteria)
	Call WriteTweetSearch
  Next
Next

Sub LoadTweetsUserTimeline(useraccount)

	' Configure the API call.
	Dim sUsername : sUsername = useraccount
	Dim iCount : iCount = 10
	Dim bExcludeReplies : bExcludeReplies = False
	Dim bIncludeRTs : bIncludeRTs = True

	Set objTweets = objASPTwitter.GetUserTimeline(sUsername, iCount, bExcludeReplies, bIncludeRTs)

End Sub

Sub WriteTweetsUserTimeline()

	%>
	<h2>User Timeline</h2>

	<ol id="Tweets"><%

    ' Assumes TypeName(objTweets) = "JScriptTypeInfo"
    If Not HasKey(objTweets, "length") Then
    	%><li>GetTweets.asp: No tweets.</li><%
        Exit Sub
    End If

	If objTweets.length = 0 Then
		%><li>GetTweets.asp: No tweets.</li><%
	End If

	If Err Then
		%><li>GetTweets.asp: invalid API response.</li><%
	End if

	Dim oTweet
	For Each oTweet In objTweets

		' Workarounds.
		' JSON parser bug workaround:
		'	- API can return invalid tweets, probably due to characters.
		' Twitter API bugs:
		'	- Filtering by the API can return additional invalid items, and seems to filter only after retrieving the requested number of items, so you get less than you asked for.
		'	- API sometimes seems to exclude replies even if that filter is not set, resulting in "*up to* count" responses and associated issues.
		If IsTweet(oTweet) Or IsRetweet(oTweet) Then

			' NOTE: A JSON viewer can be useful here: http://www.jsoneditoronline.org/
			Dim screen_name, text
			If Not IsRetweet(oTweet) Then
				screen_name = oTweet.user.screen_name
				text = URLsBecomeLinks(oTweet.text)
			Else
				screen_name = oTweet.retweeted_status.user.screen_name
				text = URLsBecomeLinks(oTweet.retweeted_status.text)
			End If

			%>
		<li>
			<b class="screen_name">@<%= screen_name %></b>
			<span class="text"><%= text %></span>
		</li><%

		End If

	Next

	%>
	</ol><%

	Response.Flush()

End Sub

Sub LoadTweetsSearch(accountuser, keyword_criteria)

	sQuery = keyword_criteria
	iCount = 10
	lMaxID = Null
	Set objTweets = objASPTwitter.GetSearch(accountuser, sQuery, iCount, lMaxID)

End Sub

Sub WriteTweetSearch()
' Assumes TypeName(objTweets) = "JScriptTypeInfo"
If Not HasKey(objTweets, "statuses") Then
	%><li>Tweets.asp: No tweets.</li><%
    	Exit Sub
End If

If objTweets.statuses.length = 0 Then
	%><li>Tweets.asp: No tweets.</li><%
End If

If Err Then
	%><li>Tweets.asp: invalid API response.</li><%
End if

	If objTweets.statuses.length > 0 Then

	Dim oTweet
	For Each oTweet In objTweets.statuses
		If IsTweet(oTweet) Or IsRetweet(oTweet) Then
			Dim screen_name, text
			If Not IsRetweet(oTweet) Then
				screen_name = oTweet.user.screen_name
				text = URLsBecomeLinks(oTweet.text)
				tweet_date = oTweet.created_at
			Else
				screen_name = oTweet.retweeted_status.user.screen_name
				text = URLsBecomeLinks(oTweet.retweeted_status.text)
				tweet_date = oTweet.created_at
			End If
			emailBody = "<p>@" & screen_name & "<br>" & URLsBecomeLinks(text) & "<br>" & tweet_date & "</p>" & emailBody
		End If
	Next

  Set objEmail = Server.CreateObject("CDO.Message")
  objEmail.BodyPart.Charset = "utf-8"
  objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
  objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="localhost"
  objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
  objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
  objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 5000
  objEmail.Configuration.Fields.Update
  objEmail.From = SENDER_EMAIL
  objEmail.To = RECIPIENT_EMAILS
  objEmail.Subject = EMAIL_SUBJECT
  objEmail.HTMLBody = emailBody
  objEmail.Send
  set objEmail = Nothing
  end if
	Response.Flush()
End Sub

Function IsTweet(ByRef oTweet)
	IsTweet = HasKey(oTweet, "user")
End Function

Function IsRetweet(ByRef oTweet)
	IsRetweet = HasKey(oTweet, "retweeted_status")
End Function

Function IsReply(ByRef oTweet)
	IsReply = Not oTweet.get("in_reply_to_user_id") = Null
End Function

Function HasKey(ByRef oTweet, ByVal sKeyName)
	HasKey = Not CStr("" & oTweet.get(sKeyName)) = ""
End Function

Function URLsBecomeLinks(sText)
	' Wrap URLs in text with HTML link anchor tags.
	Dim objRegExp
	Set objRegExp = New RegExp
	objRegExp.Pattern = "(http://[^\s<]*)"
	objRegExp.Global = True
	objRegExp.ignorecase = True
	UrlsBecomeLinks = "" & objRegExp.Replace(sText, "<a href=""$1"" target=""_blank"">$1</a>")
	Set objRegExp = Nothing
End Function
%>
