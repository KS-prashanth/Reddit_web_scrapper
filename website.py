import streamlit as st
import requests
import pandas as pd
from io import BytesIO

# Step 1: Define Reddit API credentials
CLIENT_ID = '1AZGBEgbai9_WoZxrKiuqA'  # Replace with your Reddit app's Client ID
CLIENT_SECRET = '1H8nI4z-r-uMgCTuYE9SW3jshJ4ZYw'  # Replace with your Client Secret
USERNAME = 'Historical-Event-615'  # Replace with your Reddit username
PASSWORD = 'K$p9059558215'  # Replace with your Reddit password
USER_AGENT = 'ChangeMeClient/0.1 by Historical-Event-615'  # Replace with a valid User-Agent

# Step 2: Set up authentication and headers
client_auth = requests.auth.HTTPBasicAuth(CLIENT_ID, CLIENT_SECRET)
post_data = {
    "grant_type": "password",
    "username": USERNAME,
    "password": PASSWORD
}
headers = {"User-Agent": USER_AGENT}

# Function to get access token
def get_access_token():
    response = requests.post(
        "https://www.reddit.com/api/v1/access_token",
        auth=client_auth,
        data=post_data,
        headers=headers
    )
    if response.status_code == 200:
        return response.json()['access_token']
    else:
        st.error(f"Error fetching access token: {response.status_code}")
        return None

# Function to fetch Reddit posts based on a query
def fetch_posts(access_token, query, limit=10):
    url = "https://oauth.reddit.com/search"
    headers = {
        "Authorization": f"bearer {access_token}",
        "User-Agent": USER_AGENT
    }
    params = {
        "q": query,
        "limit": limit,
        "sort": "relevance"
    }
    response = requests.get(url, headers=headers, params=params)
    if response.status_code == 200:
        return response.json()['data']['children']
    else:
        st.error(f"Error fetching posts: {response.status_code}")
        return []

# Function to save posts to an Excel file
def save_to_excel(posts, query):
    # Prepare data for the Excel file
    posts_data = []
    comments_data = []

    for post in posts:
        post_data = {
            "Title": post['data']['title'],
            "Author": post['data']['author'],
            "Subreddit": post['data']['subreddit'],
            "URL": post['data']['url'],
            "Score": post['data']['score'],
            "Created UTC": post['data']['created_utc']
        }
        posts_data.append(post_data)

        for comment in post['data'].get('comments', []):
            comments_data.append({
                "Post Title": post['data']['title'],
                "Comment Author": comment['author'],
                "Comment Body": comment['body'],
                "Comment Score": comment['score']
            })

    # Create a BytesIO object for the Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame(posts_data).to_excel(writer, sheet_name='Posts', index=False)
        pd.DataFrame(comments_data).to_excel(writer, sheet_name='Comments', index=False)

    output.seek(0)
    return f"{query.replace(' ', '_')}.xlsx", output

# Main function to build the Streamlit app
def main():
    st.title("Reddit Scraper")

    # Input for the query
    query = st.text_input("Enter search query:", "problems with AI")
    num_posts = st.slider("Number of posts to fetch:", 1, 100, 10)

    if st.button("Fetch Data"):
        st.info("Fetching data from Reddit...")

        # Get access token
        access_token = get_access_token()
        if not access_token:
            return

        # Fetch posts
        posts = fetch_posts(access_token, query, num_posts)
        if not posts:
            st.warning("No posts found.")
            return

        # Save to Excel
        filename, excel_file = save_to_excel(posts, query)
        st.success("Data fetched successfully!")

        # Download button for the Excel file
        st.download_button(
            label="Download Excel File",
            data=excel_file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
