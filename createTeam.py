################################################################################
##     This script uses Microsoft Graph to create a new Teams space
## using a custom template. 
################################################################################
import requests
#from requests_toolbelt.utils import dump
import json
# config file with site-specific values
from config import strClientID, strClientSecret, strGraphAuthURL, strTenantID
################################################################################
# Function definitions
################################################################################
################################################################################
# End of functions
################################################################################
iPlanID = 5 # In a real-world scenario, create the planner first (https://docs.microsoft.com/en-us/graph/api/planner-post-plans?view=graph-rest-1.0)
postData = {"grant_type": "client_credentials","client_id" : strClientID,"client_secret": strClientSecret,"scope": "https://graph.microsoft.com/.default"}
    
r = requests.post(strGraphAuthURL, data=postData)

strJSONResponse = r.text
if len(strJSONResponse) > 5:
    jsonResponse = json.loads(strJSONResponse)
    print(jsonResponse)
    strAccessToken = jsonResponse['access_token']

    getHeader = {"Authorization": "Bearer " + strAccessToken }

    strBody  = json.dumps({
    "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates/standard",
    "visibility": "Private",
    "displayName": "Fancy Programmatic Team",
    "description": "This is a sample Teams space created using the Graph API",
    "owners@odata.bind": [
           "https://graph.microsoft.com/beta/users('3942de0d-e478-4d83-afa1-6cc12c319595')"
    ],
    "channels": [
        {
            "displayName": "Announcements",
            "isFavoriteByDefault": "true",
            "description": "This is a sample announcements channel that is favorited by default. Use this channel to make important team, product, and service announcements."
        },
        {
            "displayName": "Training",
            "isFavoriteByDefault": "false",
            "description": "Channel with pre-configured web tab.",
            "tabs": [
                {
                    "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.web')",
                    "name": "Graph API Documentation",
                    "configuration": {
                        "contentUrl": "https://docs.microsoft.com/en-us/graph/overview?toc=./toc.json&view=graph-rest-beta"
                    }
                }
            ]
        },
        {
            "displayName": "App Development and Issues",
            "description": "Channel used to discuss issue prioritization -- Planner backlog is in here.",
            "tabs": [
                {
                    "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.planner')",
                    "name": "Development Backlog",
                    "configuration": {
                        "contentUrl": "https://tasks.office.com/%s/Home/PlannerFrame?page=7&planId=%s" % (strTenantID, iPlanID),
                        "removeUrl": "https://tasks.office.com/%s/Home/PlannerFrame?page=7&planId=%s" % (strTenantID, iPlanID),
                        "websiteUrl": "https://tasks.office.com/%s/Home/PlannerFrame?page=7&planId=%s" % (strTenantID, iPlanID)
                    }
                },
                {
                    "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.file.staticviewer.powerpoint')",
                    "name": "PowerPoint Product Roadmap",
                    "configuration": {
                        "contentUrl": "https://windstream.sharepoint.com/:p:/r/sites/LJRSandboxTeam/_layouts/15/Doc.aspx?sourcedoc=%7B5e5cb031-9e28-4859-a478-b96cdbe1b7e7%7D&action=edit&uid=%7B5E5CB031-9E28-4859-A478-B96CDBE1B7E7%7D&ListItemId=79&ListId=%7B7AA859DA-3CFC-4F68-8DDA-2608F246F158%7D&odsp=1&env=prod",
                        "entityID": "%7B5e5cb031-9e28-4859-a478-b96cdbe1b7e7%7D"
                    }
                }
            ]
        }
    ],
    "memberSettings": {
        "allowCreateUpdateChannels": "false",
        "allowDeleteChannels": "false",
        "allowAddRemoveApps": "false",
        "allowCreateUpdateRemoveTabs": "false",
        "allowCreateUpdateRemoveConnectors": "false"
    },
    "guestSettings": {
        "allowCreateUpdateChannels": "false",
        "allowDeleteChannels": "false"
    },
    "funSettings": {
        "allowGiphy": "true",
        "giphyContentRating": "Moderate",
        "allowStickersAndMemes": "true",
        "allowCustomMemes": "true"
    },
    "messagingSettings": {
        "allowUserEditMessages": "true",
        "allowUserDeleteMessages": "true",
        "allowOwnerDeleteMessages": "true",
        "allowTeamMentions": "false",
        "allowChannelMentions": "false"
    },
    "installedApps": [
        {
            "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.vsts')"
        },
        {
            "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('d58f3268-9fe3-44f5-8a4c-abef78b77134')"   # RememberThis
        },
        {
            "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('75a6c3a3-aadd-4f97-8118-99e7c2335cb2')"   # Twitter
        }
    ]
    })

    postRecord = requests.post("https://graph.microsoft.com/beta/teams",headers={"Content-Length": str(len(json.dumps(strBody))), 'content-Type': "application/json", "Authorization": "Bearer " + strAccessToken}, data=strBody)
 #   data = dump.dump_all(postRecord)
 #   print("Session data:\t%s" % data.decode('utf-8'))
    print("HTTP Status Code:\t%s\nResult code content:\t%s" % (postRecord.status_code, postRecord.content))