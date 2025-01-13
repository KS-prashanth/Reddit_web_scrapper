import streamlit as st
import requests
import pandas as pd
import json
import os
from io import BytesIO

# Reddit API credentials
CLIENT_ID = '1AZGBEgbai9_WoZxrKiuqA'
CLIENT_SECRET = '1H8nI4z-r-uMgCTuYE9SW3jshJ4ZYw'
USERNAME = 'Historical-Event-615'
PASSWORD = 'K$p9059558215'
USER_AGENT = 'ChangeMeClient/0.1 by Historical-Event-615'

# Function to get access token
def get_access_token():
    try:
        auth = requests.auth.HTTPBasicAuth(CLIENT_ID, CLIENT_SECRET)
        post_data = {"grant_type": "password", "username": USERNAME, "password": PASSWORD}
        headers = {"User-Agent": USER_AGENT}
        response = requests.post("https://www.reddit.com/api/v1/access_token", auth=auth, data=post_data, headers=headers)
        if response.status_code == 200:
            return response.json().get('access_token')
        else:
            st.error(f"Access token error: {response.status_code}, {response.text}")
            return None
    except Exception as e:
        st.error(f"Exception during token fetch: {str(e)}")
        return None

# Function to fetch posts
def fetch_reddit_posts(access_token, query, limit=10):
    try:
        url = "https://oauth.reddit.com/search"
        headers = {"Authorization": f"bearer {access_token}", "User-Agent": USER_AGENT}
        params = {"q": query, "limit": limit, "sort": "relevance"}
        response = requests.get(url, headers=headers, params=params)
        if response.status_code == 200:
            return response.json()
        else:
            st.error(f"Error fetching posts: {response.status_code}, {response.text}")
            return None
    except Exception as e:
        st.error(f"Exception fetching posts: {str(e)}")
        return None

# Function to fetch comments
def fetch_post_comments(access_token, post_id):
    try:
        url = f"https://oauth.reddit.com/comments/{post_id}"
        headers = {"Authorization": f"bearer {access_token}", "User-Agent": USER_AGENT}
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json()
        else:
            st.error(f"Error fetching comments: {response.status_code}, {response.text}")
            return None
    except Exception as e:
        st.error(f"Exception fetching comments: {str(e)}")
        return None

# Function to save posts and comments to an Excel file
# Function to save posts and comments to an Excel file
# Function to save posts and comments to an Excel file
def save_to_excel(posts, filename):
    def clean_sheet_name(name):
        # Limit to 31 characters and remove invalid characters
        name = ''.join(c for c in name if c not in r'\\/?*[]:').strip()
        return name[:31]  # Limit to 31 characters

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # Save posts to the first sheet
        posts_df = pd.DataFrame(posts)
        posts_df.to_excel(writer, sheet_name='Posts', index=False)

        # Save comments for each post in separate sheets
        existing_sheet_names = set()
        for post in posts:
            if 'Comments' in post:
                comments_df = pd.DataFrame(post['Comments'])
                sheet_name = clean_sheet_name(post['Title'])

                # Ensure sheet name is unique
                counter = 1
                original_sheet_name = sheet_name
                while sheet_name in existing_sheet_names or not sheet_name:
                    sheet_name = f"{original_sheet_name[:25]}_{counter}"  # Truncate and add counter
                    counter += 1

                existing_sheet_names.add(sheet_name)  # Add to the set of used names
                comments_df.to_excel(writer, sheet_name=sheet_name, index=False)


# Streamlit app
def main():
    st.title("Reddit Data Fetcher")

    query = st.text_input("Enter search query:", "problems with AI")
    limit = st.slider("Number of posts to fetch:", 1, 100, 10)

    if st.button("Fetch Data"):
        st.info("Fetching data from Reddit...")
        access_token = get_access_token()

        if not access_token:
            st.error("Failed to retrieve access token. Check your credentials.")
            return

        st.info("Access token retrieved. Fetching posts...")
        posts_data = fetch_reddit_posts(access_token, query, limit)

        if not posts_data or 'data' not in posts_data or 'children' not in posts_data['data']:
            st.error("No posts found or error in fetching posts.")
            return

        posts = posts_data['data']['children']
        if not posts:
            st.warning("No posts available for the given query.")
            return

        posts_list = []

        for post in posts:
            post_id = post['data'].get('id', "N/A")
            post_info = {
                "Title": post['data'].get('title', "No Title"),
                "Author": post['data'].get('author', "Unknown"),
                "Subreddit": post['data'].get('subreddit', "Unknown"),
                "Score": post['data'].get('score', 0),
                "Post ID": post_id
            }

            # Fetch comments for each post
            comments_data = fetch_post_comments(access_token, post_id)
            if comments_data and len(comments_data) > 1 and 'data' in comments_data[1]:
                comments = comments_data[1]['data'].get('children', [])
                comments_list = [
                    {
                        "Comment Author": comment['data'].get('author', 'N/A'),
                        "Comment Body": comment['data'].get('body', 'N/A'),
                        "Comment Score": comment['data'].get('score', 0)
                    }
                    for comment in comments if 'body' in comment.get('data', {})
                ]
                post_info['Comments'] = comments_list
            else:
                post_info['Comments'] = []

            posts_list.append(post_info)

        # Save data to Excel
        excel_file = BytesIO()
        save_to_excel(posts_list, excel_file)
        excel_file.seek(0)

        st.success("Data fetched and saved successfully!")

        # Provide download link
        st.download_button(
            label="Download Excel File",
            data=excel_file,
            file_name="reddit_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
