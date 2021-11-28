import requests


CONTENT_TYPE = "application/json"


def get_headers(bearer_token):
    """This is the headers for the Microsoft Graph API calls"""
    return {
        "Accept": CONTENT_TYPE,
        "Authorization": f"Bearer {bearer_token}",
        "ConsistencyLevel": "eventual",
    }


def get_token_for_user_application(client_id, tenant_id, username, password):
    """
    Get Token on behalf of a user using username/password
    """

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    payload = (
        f"grant_type=password&client_id={client_id}&username={username}&password={password}"
        "&scope=User.Read"
    )
    headers = {}

    resp = requests.request("POST", url, headers=headers, data=payload)
    if resp.status_code != 200:
        return None
    return resp.json()["access_token"]


def get_token_for_client_application(client_id, client_secret, tenant_id):
    """
    Get Token on behalf of a client application using client_secret/client_id
    """

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    payload = (
        f"grant_type=client_credentials&client_id={client_id}&client_secret={client_secret}"
        "&scope=https%3A//graph.microsoft.com/.default"
    )
    headers = {}

    resp = requests.request("POST", url, headers=headers, data=payload)
    if resp.status_code != 200:
        return None
    return resp.json()["access_token"]


def get_signedin_user_data(bearer_token):
    """
    Get SignedIn user data
    """

    resp = requests.get(f"https://graph.microsoft.com/v1.0/me",
                        headers=get_headers(bearer_token))

    json_resp = resp.json()
    return json_resp


def get_ms_teams_users(bearer_token, filters=""):
    """
    Get/Search MS Teams users
    """

    if filters:
        filters = f"$filter={filters}"

    url = f"https://graph.microsoft.com/beta/users?{filters}"
    resp = requests.get(url, headers=get_headers(bearer_token))
    if resp.status_code != 200:
        print(resp.json())
        return None

    json_resp = resp.json()
    try:
        return json_resp["value"]
    except KeyError as err:
        return []


def send_message_to_ms_teams_user(bearer_token, sender_ms_team_id, user_ms_teams_id, message):
    """
    Send Message to MS Teams user is done in 2 steps:
        1: Create chat
        2: Use chat-id created in 1st step and send message to the user.
    """
    # 1st step: Create chat
    creat_chat_url = "https://graph.microsoft.com/v1.0/chats"
    data = {
        "chatType": "oneOnOne",
        "members": [
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_ms_teams_id}')",
            },
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{sender_ms_team_id}')",
            },
        ],
    }

    resp = requests.post(
        creat_chat_url, headers=get_headers(bearer_token), json=data)
    json_resp = resp.json()
    if resp.status_code not in [200, 201]:
        return False

    # 2nd step: Use created chat-id and send message to it.
    chat_id = json_resp["id"]
    send_message_url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"

    messsage_data = {"body": {"contentType": "html", "content": message}}
    resp = requests.post(send_message_url, headers=get_headers(
        bearer_token), json=messsage_data)
    json_resp = resp.json()
    if resp.status_code not in [200, 201]:
        return False

    return True


def get_ms_teams_users_using_emails(bearer_token, emails=[]):
    filters = [f"mail eq '{email}'" for email in emails]
    filters = " OR ".join(filters)
    users = get_ms_teams_users(bearer_token, filters=filters)

    return users
