import re
import inquirer
from azure.identity import InteractiveBrowserCredential
from msgraph.core import GraphClient
import sys, readchar

# Fix platform related backspace issue
if any(x in sys.platform for x in ['darwin', 'linux']):
    readchar.key.BACKSPACE = '\x7F'

browser_credential = InteractiveBrowserCredential(client_id='0d498bab-7c7a-4c20-a54b-e44fcda36532')
client = GraphClient(credential=browser_credential,
                     scopes=['TeamSettings.ReadWrite.All', 'TeamMember.ReadWrite.All', 'User.Invite.All', 'Mail.Send'])

result = client.get('/me/joinedTeams')
teams = result.json()['value']
data = {}
for team in teams:
    data[team['displayName']] = team['id']

questions = [
    inquirer.Text('email', message="輸入學員Email："),
    inquirer.List('team',
                  message="要加入哪個群組呢？",
                  choices=data.keys())
]
answers = inquirer.prompt(questions)

team = client.get(f'/teams/{data[answers["team"]]}').json()
invite = client.post(f'/invitations', json={
    "invitedUserEmailAddress": answers['email'],
    "inviteRedirectUrl": team['webUrl'],
}).json()

redeem = invite['inviteRedeemUrl']

client.post(f'/teams/{data[answers["team"]]}/members/add', json={
    "values": [{
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": ["guest"],
        "user@odata.bind": f"""https://graph.microsoft.com/v1.0/users('{invite["invitedUser"]["id"]}')"""
    }]
}).json()

client.post('/me/sendMail', json={
    "message": {
        "subject": "Microsoft Teams 邀請函通知",
        "body": {
            "contentType": "HTML",
            "content": f"""
            您好，已經邀請您至 <b>{team['displayName']}</b><br/>
            請點擊這個連結加入組織： <a href="{redeem}" target="_blank">邀請連結</a><br/>
            邀請的 Teams 團隊連結： <a href="{team['webUrl']}" target="_blank">{team['displayName']}</a>
            """
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": answers['email']
                }
            }
        ]
    },
    "saveToSentItems": "true"
})
