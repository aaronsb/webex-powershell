
Class WebExUserAccount
{
    [boolean]$RunEnabled
    [string]$WebexID
    [string]$firstname
    [string]$lastname
    [string]$title
    [string]$categoryid
    [string]$description
    [string]$officeGreeting
    [string]$company
    [string]$email
    [string]$password
    [string]$passwordHint
    [string]$passwordHintAnswer
    [string]$language
    [string]$locale
    [string]$timeZoneId
    [string]$active
    [string]$address1
    [string]$address2
    [string]$city
    [string]$state
    [string]$zipcode
    [string]$country
    [boolean]$host
    [boolean]$siteAdmin
    [boolean]$roSiteAdmin
    [boolean]$teleConfTollFreeCallIn
    [boolean]$teleConfCallOutInternational
    [boolean]$teleConfCallIn
    [boolean]$voiceOverIp
    [boolean]$labAdmin
    [boolean]$otherTelephony
    [boolean]$teleConfCallInInternational
    [boolean]$attendeeOnly
    [boolean]$recordingEditor
    [boolean]$meetingAssist
    [boolean]$HQvideo
    [boolean]$forceChangePassword
    [boolean]$resetPassword
    [boolean]$lockAccount
    [boolean]$isMyWebExPro
    [boolean]$myContact
    [boolean]$myProfile
    [boolean]$myMeetings
    [boolean]$myFolders
    [boolean]$trainingRecordings
    [boolean]$recordedEvents
    [string]$totalStorageSize
    [boolean]$myReports
    [string]$myComputer
    [boolean]$personalMeetingRoom
    [boolean]$myPartnerLinks
    [boolean]$myWorkspaces
}


Class WebExAuthContext
{
    [string]$WebExServiceAccountName
    [string]$WebExServiceAccountPassword
    [string]$WebExSiteID
    [string]$WebExPartnerID
}