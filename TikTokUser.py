import re

from TikTokApi import TikTokApi
import asyncio
import os
import xlsxwriter
import openpyxl

workbook = xlsxwriter.Workbook('Creator_Info.xlsx')
worksheet = workbook.add_worksheet()
handleList = []

ms_token = os.environ.get(
    "ms_token", None
) 
context_options = {
    'viewport' : { 'width': 1280, 'height': 1024 },
    'user_agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36'
}


async def get_user_info(api, handle_name):
    # Avoid "Can not find this acc" status
    try:
        user = api.user(handle_name)
        user_data = await user.info()
        return user_data
    except Exception as e:
        print(f"Error retrieving user information for handle {handle_name}: {e}")
        return None

# Get Info by search their handle name
async def search_handle_name():
    async with TikTokApi() as api:
        await api.create_sessions(ms_tokens=[ms_token], num_sessions=1, sleep_after=3, context_options=context_options)
        dataframe = openpyxl.load_workbook("HandleList.xlsx")
        sheet = dataframe.active
        handle_list = [cell.value for cell in sheet['A']]
        row = 0

        for name in handle_list:
            user_data = await get_user_info(api, name)
            if user_data is None:
                row += 1
                continue
            user_info = user_data["userInfo"]
            follower_count = user_info["stats"]["followerCount"]
            signature = user_info["user"]["signature"]
            worksheet.write(row, 0, name)
            worksheet.write(row, 1, follower_count)
            worksheet.write(row, 2, signature)
            signature = modify_email(signature)
            worksheet.write(row, 3, signature)
            print(follower_count)
            print()
            row += 1

    workbook.close()

async def get_hash_tag(api, hashtag):
    try:
        tag = api.hashtag(hashtag)
        if tag is None:
            print(f"Hashtag '{hashtag}' does not exist.")
            return None
        return tag
    except Exception as e:
        print(f"Error retrieving hashtag information for handle {hashtag}: {e}")
        return None

async def trending_videos():
    async with TikTokApi() as api:
        await api.create_sessions(ms_tokens=[ms_token], num_sessions=1, sleep_after=3, context_options=context_options)
        dataframe = openpyxl.load_workbook("HashTags.xlsx")
        sheet = dataframe.active
        hashtags = [cell.value for cell in sheet['A']]
        row = 0
        # Change the tag name for different type of videos
        for hashtag in hashtags:
            tag = await get_hash_tag(api, hashtag)
            if tag is None:
                continue
            try:
                async for video in tag.videos():
                    # print("DICT HERE: ")
                    # print(video)
                    # print(video.as_dict)
                    # print("DICT END")
                    user_name = video.author.username
                    if user_name in handleList:
                        continue
                    handleList.append(user_name)
                    print(user_name)
                    user = api.user(user_name)
                    user_data = await get_user_info(api, user_name)
                    if user_data is None:
                        continue
                    user_info = user_data["userInfo"]
                    # print(user_data)
                    follower_count = user_info["stats"]["followerCount"]
                    signature = user_info["user"]["signature"]
                    worksheet.write(row, 0, user_name)
                    worksheet.write(row, 1, follower_count)
                    worksheet.write(row, 2, signature)
                    signature = modify_email(signature)
                    worksheet.write(row, 3, signature)
                    print(follower_count)
                    # print("Signature: " + signature)
                    print()
                    row += 1
            except AttributeError as e:
                print(f"AttributeError encountered for hashtag '{hashtag}' : {e}")
                continue

    workbook.close()


def modify_email(signature):
    email_pattern = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4})'

    found_emails = re.findall(email_pattern, signature)
    signature = found_emails[0] if found_emails else None

    valid_formats = ['.com', '.co', '.net', '.org']
    if isinstance(signature, str):
        if '@yahoo' in signature:
            if not signature.endswith('.com'):
                signature = signature.split('@yahoo')[0] + '@yahoo.com'

        if '@hotmail' in signature:
            if not signature.endswith('.com'):
                signature = signature.split('@hotmail')[0] + '@hotmail.com'

        if '@icloud' in signature:
            if not signature.endswith('.com'):
                signature = signature.split('@icloud')[0] + '@icloud.com'

        if '@gmail' in signature:
            if not signature.endswith('.com'):
                signature = signature.split('@gmail')[0] + '@gmail.com'

        if not any(format in signature for format in valid_formats):
            signature = None

        # if signature is not None:
        #     if signature.startswith('Collab'):
        #         signature = None
        #     if signature.startswith(('Business')):
        #         signature = None

        return signature



if __name__ == "__main__":
    asyncio.run(trending_videos())