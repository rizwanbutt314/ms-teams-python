import ms_teams

# MS Teams App configs
CLIENT_ID = "test_client_id"
CLIENT_SECRET = "test_client_secret"
TENANT_ID = "test_tenant_id"
USERNAME = "test_user"
PASSWORD = "test_password"


def main(message):

    # Get Client application token
    client_app_token = ms_teams.get_token_for_client_application(
        CLIENT_ID, CLIENT_SECRET, TENANT_ID)

    # Get User application token
    user_app_token = ms_teams.get_token_for_user_application(
        CLIENT_ID, TENANT_ID, USERNAME, PASSWORD)

    # Get SignedIn user data
    signedin_user_data = ms_teams.get_signedin_user_data(user_app_token)
    sender_ms_teams_id = signedin_user_data["id"]

    # Search user(s) with email
    ms_teams_users = ms_teams.get_ms_teams_users_using_emails(
        client_app_token, emails=["test@test.com"])

    # Get first user id of above search
    ms_teams_user_id = ms_teams_users[0]["id"]

    # Send message
    is_message_sent = ms_teams.send_message_to_ms_teams_user(
        user_app_token, sender_ms_teams_id, ms_teams_user_id, message)

    if is_message_sent:
        print("Message sent")
    else:
        print("Message sending Failed")


if __name__ == "__main__":
    message = "Hello World!"
    main(message)
