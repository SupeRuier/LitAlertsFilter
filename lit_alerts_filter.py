import imaplib
import email
from email.header import decode_header
import re
import datetime
import json
import os
import html

# Set the proxy (设置代理)
os.environ['http_proxy'] = "http://127.0.0.1:7890"
os.environ['https_proxy'] = "http://127.0.0.1:7890"

# Login to the IMAP server (登录到IMAP服务器)
def login(email_address, email_password):
    mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    mail.login(email_address, email_password)
    return mail

# Search for emails within a specified date range (搜索指定日期范围内的邮件)
def search_emails_in_date_range(mail, start_date, end_date):
    mail.select('inbox')
    # Format the date (格式化日期)
    date_format = "%d-%b-%Y"
    since_date = start_date.strftime(date_format)
    before_date = end_date.strftime(date_format)
    # Build the search command (构建搜索命令)
    search_criteria = f'(SINCE "{since_date}" BEFORE "{before_date}")'
    status, messages = mail.search(None, search_criteria)
    return messages

# Decode email subject and content (解码邮件主题和内容)
def decode_email_subject_and_body(msg):
    subject = decode_header(msg["Subject"])[0][0]
    if isinstance(subject, bytes):
        subject = subject.decode()
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            if content_type == "text/plain" and "attachment" not in content_disposition:
                body = part.get_payload(decode=True).decode()
                return subject, body
    else:
        body = msg.get_payload(decode=True).decode()
        return subject, body
    return None, None

# Parse the email content (解析邮件内容)
def parse_email(mail, email_id):
    res, msg = mail.fetch(email_id, "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # Parse the email content (解析邮件内容)
            msg = email.message_from_bytes(response[1])
            # Get the email sender (获取邮件发件人)
            from_ = decode_header(msg.get("From"))[-1][0]
            
            if isinstance(from_, bytes):
                from_ = from_.decode()

            # Check if it is from Google Scholar (检查是否来自Google Scholar)
            if "scholaralerts" in from_: #"scholaralerts-noreply@google.com"
                subject, body = decode_email_subject_and_body(msg)
                return subject, body
    return None, None

# Extract literature information from the email body (从邮件正文中提取文献信息)
def extract_literature_info(body):
    if not isinstance(body, str):
        raise ValueError("The body must be a string.")

    # Updated regular expression to extract the direct link from the redirect URL (更新后的正则表达式，用于提取重定向URL中的直接链接)
    literature_pattern = re.compile(
        r'<a href="https://scholar.google.com/scholar_url\?url=([^"&]+)[^"]*" class="gse_alrt_title"[^>]*>(.*?)</a>'
        r'.*?<div style="color:#006621;line-height:18px">(.*?)(?:\s+-\s+([^<]+))?</div>',
        re.DOTALL
    )

    matches = literature_pattern.findall(body)
    literature_list = []
    
    for url, title, authors, year_journal in matches:
        # Decode HTML entities (解码HTML实体)
        decoded_url = html.unescape(url)
        literature_dict = {
            "title": title.strip(),
            "authors": authors.strip(),
            "year_journal": year_journal.strip(),
            "url": decoded_url  # Use the decoded direct URL (使用解码后的直接URL)
        }
        literature_list.append(literature_dict)
    
    return literature_list  # Return the list of dictionaries directly (直接返回字典列表)

# Save the list of literature information to a JSON file (将文献信息列表保存到JSON文件)
def save_to_json(literatures, filename="literature_info.json"):
    """
    Save the list of literature information to a JSON file.
    (将文献信息列表保存到JSON文件。)

    :param literatures: List of dictionaries, where each dictionary contains info of a literature. (文献信息的字典列表，每个字典包含一篇文献的信息。)
    :param filename: The filename for the JSON file. (JSON文件的文件名。)
    """
    with open(filename, 'w', encoding='utf-8') as file:
        # Convert the list of literature information to a JSON formatted string and write to the file. (将文献信息列表转换为JSON格式的字符串并写入文件。)
        json.dump(literatures, file, ensure_ascii=False, indent=4)

# Main function to run the process (运行过程的主函数)
def main(email_address, email_password, start_date, end_date, keywords):
    mail = login(email_address, email_password)
    print('Successfully connected!')
    messages = search_emails_in_date_range(mail, start_date, end_date)
    unique_titles = set()  # Set to store processed titles (集合用于存储已处理的标题)
    all_literatures = []  # List to store all literatures (存储所有文献的列表)

    print('Start processing!')
    if messages[0]:  # Check if there are returned emails (检查是否有邮件返回)
        for email_id in messages[0].split():
            _, body = parse_email(mail, email_id)  # Assume the subject returned by parse_email is not important (假设parse_email返回的主题不重要)
            if body:
                literatures = extract_literature_info(body)  # Now returns a list of dictionaries (现在返回的是字典列表)
                # literatures = json.loads(extract_literature_info(body))  # Parse literature information and convert to JSON (解析文献信息并转换为JSON)

                for literature in literatures:
                    title = literature['title']
                    # If the title is not in the known title set, process the literature and add it to the set (如果标题不在已知标题集合中，处理文献并添加到集合中)
                    if title not in unique_titles:
                        unique_titles.add(title)
                        all_literatures.append(literature)

    # Save non-duplicate literature information as a JSON file (将非重复的文献信息保存为JSON文件)
    start_date_str = start_date.strftime("%Y%m%d")
    end_date_str = end_date.strftime("%Y%m%d")
    raw_filename = f"raw_{start_date_str}{end_date_str}.json"
    save_to_json(all_literatures, raw_filename)

    # Filter literatures (筛选文献)
    selected_literatures = [lit for lit in all_literatures if any(kw.lower() in lit['title'].lower() for kw in keywords)]

    # Save filtered literature information (保存筛选后的文献信息)
    keywords_str = "_".join(keywords)
    selected_filename = f"selected_{start_date_str}{end_date_str}_{keywords_str}.json"
    save_to_json(selected_literatures, selected_filename)

# Run the main function (运行主函数)
if __name__ == "__main__":
    email_address = "your address"
    email_password = "your password"
    # Set the start and end dates for searching emails (设置搜索邮件的起始和结束日期)
    start_date = datetime.datetime.strptime("2023-12-01", "%Y-%m-%d")
    end_date = datetime.datetime.strptime("2023-12-31", "%Y-%m-%d")
    # end_date = datetime.datetime.now()  # Or set to a specific end date (或设置为特定的结束日期)
    # Define a list of keywords (定义关键词列表)
    keywords = ['active']
    main(email_address, email_password, start_date, end_date, keywords)
