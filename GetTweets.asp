<!--#include file="Libs/ASPTwitter/ASPTwitter.asp"-->
<%
const TWITTER_API_CONSUMER_KEY = 	    	""
const TWITTER_API_CONSUMER_SECRET = 		""
const TWITTER_BEARER_TOKEN = ""

TWITTER_ACCOUNTS = Array("@microsoft","@oracle")

KEYWORDS = Array("ASP.NET 8","Vietnam")
SENDER_EMAIL = ""
RECIPIENT_EMAILS = ""
EMAIL_SUBJECT = "Tweets Update Email"

if TWITTER_API_CONSUMER_KEY = "" OR TWITTER_API_CONSUMER_SECRET = "" OR TWITTER_BEARER_TOKEN = "" Then
  response.write "Please fill in your Twitter API Keys"
  response.end
End if

if KEYWORDS = "" OR SENDER_EMAIL = "" OR RECIPIENT_EMAILS = "" Then
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
  For Each search_criteria in KEYWORDS
    Call LoadTweetsSearch(twitter_account, search_criteria)
	Call WriteTweetSearch
  Next
Next

Sub LoadTweetsSearch(accountuser, keyword_criteria)

	sQuery = keyword_criteria
	iCount = 10
	lMaxID = Null
	Set objTweets = objASPTwitter.GetSearch(accountuser, sQuery, iCount, lMaxID)

End Sub

Sub WriteTweetSearch()
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
