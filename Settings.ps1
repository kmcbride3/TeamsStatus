# Configure the variables below that will be used in the script
$OHUser = "<Insert username>" # OpenHAB Cloud Connector User Example: openhab@openhab.org
$OHPassword = "<Insert token>" # OpenHAB Cloud Connector Password Example: eyJ0eXAiOiJKV1...
$UserName = "<Insert Windows UserName>" # Windows User; When not sure, open a command prompt and type: echo %USERNAME%
$OHUrl = "https://myopenhab.org" # Url to OpenHAB Cloud Connector

# Set language variables below
$lgAvailable = "Available"
$lgBusy = "Busy"
$lgOnThePhone = "On the phone"
$lgAway = "Away"
$lgBeRightBack = "Be right back"
$lgDoNotDisturb = "Do not disturb"
$lgPresenting = "Presenting"
$lgFocusing = "Focusing"
$lgInAMeeting = "In a meeting"
$lgOffline = "Offline"
$lgNotInACall = "Not in a call"
$lgInACall = "In a call"

# Set entities to post to
$entityStatus = "<Insert status name>" # Name of Teams Status Variable in OpenHAB
$entityActivity = "<Insert activity name>" # Name of Teams Activity Variable in OpenHAB
