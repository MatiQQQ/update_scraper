from bs4 import BeautifulSoup, Tag
import requests
from datetime import datetime
import json
import re

initial_links = [
    {
        'url': "https://www.catalog.update.microsoft.com/Search.aspx?q=CURRENT_YEAR-CURRENT_MONTH%20Cumulative%20Update%20for%20Windows%2010%20Version%2021H2%20for%20x64-based%20Systems%20",
        'type': 'cumulative',
        'version': '21H2',
        'text': 'Cumulative Update for Windows 10 Version 21H2 for x64-based Systems'
    },
    {
        'url': "https://www.catalog.update.microsoft.com/Search.aspx?q=CURRENT_YEAR-CURRENT_MONTH%20Cumulative%20Update%20for%20Windows%2010%20Version%2022H2%20for%20x64-based%20Systems%20",
        'type': 'cumulative',
        'version': '22H2',
        'text': 'Cumulative Update for Windows 10 Version 22H2 for x64-based Systems'
    },
    {
        'url': "https://www.catalog.update.microsoft.com/Search.aspx?q=CURRENT_YEAR-CURRENT_MONTH%20Cumulative%20Update%20for%20.NET%20Framework%203.5%20and%204.8.1%20for%20Windows%2010%20Version%2021H2%20for%20x64",
        'type': 'net',
        'version': '21H2',
        'text': 'Cumulative Update for .NET Framework 3.5 and 4.8.1 for Windows 10 Version 21H2 for x64'
    },
    {
        'url': "https://www.catalog.update.microsoft.com/Search.aspx?q=CURRENT_YEAR-CURRENT_MONTH%20Cumulative%20Update%20for%20.NET%20Framework%203.5%20and%204.8.1%20for%20Windows%2010%20Version%2022H2%20for%20x64",
        'type': 'net',
        'version': '22H2',
        'text': 'Cumulative Update for .NET Framework 3.5 and 4.8.1 for Windows 10 Version 22H2 for x64'
    },
]


def check_connection(url: str) -> bool:
    response = requests.get(url)
    if response.status_code == 200:
        return True
    else:
        return False


def create_url(url_template: str, p_current_year: int, p_current_month: int) -> str:
    new_url = url_template.replace('CURRENT_YEAR', str(p_current_year)).replace('CURRENT_MONTH',
                                                                                str(p_current_month).zfill(2))
    return new_url


def check_element_visibility(url_template: str, search_text: str, max_depth: int = 5, current_depth: int = 0,
                             year: int | None = None, month: int | None = None) -> dict:
    if current_depth > max_depth:
        return {'error': 'Max depth reached without finding the element.'}

    if year is None or month is None:
        now = datetime.now()
        year, month = now.year, now.month

    # Ensure year and month are treated as integers for arithmetic operations
    year, month = int(year), int(month)

    current_url = create_url(url_template, year, month)
    try:
        response = requests.get(current_url)
        if response.ok:  # Better check for successful response
            soup = BeautifulSoup(response.content, 'html.parser')

            # Find the table by class 'resultsBackGround', then navigate to <tbody> > <tr> > <td> > <a>
            table = soup.find('table', class_='resultsBackGround')
            if table:
                trs = table.find_all('tr')
                for tr in trs[1:]:
                    # Assuming there's only one <td> and <a> per <tr>
                    td = tr.find_all('td')
                    a_tag = td[1].find('a') if td else None
                    if a_tag and search_text in a_tag.get_text(strip=True):
                        return {'status': 'Found', 'depth': current_depth, 'html_content': response.content,
                                'year': year,
                                'month': month}

            # If not found, decrement the month and year if necessary
            month -= 1
            if month == 0:
                month = 12
                year -= 1

            return check_element_visibility(url_template, search_text, max_depth, current_depth + 1, year, month)
    except Exception as e:
        return {'error': str(e)}

    # If the function reaches this point without returning, it means the text wasn't found
    return {'error': 'Text not found in any anchor within the specified depth.'}


def find_element_with_text(html_content: str, search_text: str, exclude_text: str = "Dynamic") -> Tag | None:
    """
    Searches through a table for an element (bs4 Tag) that contains specific text and does not contain another specified text.
    Returns the bs4 Tag if found, or None if no such element exists.
    """

    try:

        soup = BeautifulSoup(html_content, 'html.parser')

        # Find the table by class 'resultsBackGround'
        table = soup.find('table', class_='resultsBackGround')
        if table:
            trs = table.find_all('tr')
            for tr in trs[1:]:
                tds = tr.find_all('td')
                td = tds[1]
                a_tag = td.find('a')
                a_tag_text = a_tag.get_text(strip=True)
                if search_text in a_tag_text and exclude_text not in a_tag_text:
                    return td  # Return the bs4 Tag object directly

    except Exception as e:
        print(f"Error: {e}")
        return None

    # If no matching element is found, return None
    return None


def find_kb_number(element: Tag) -> str:
    anchor_tag_text = element.find("a").text
    last_element_of_split_text = anchor_tag_text.strip().split(" ")[-1]
    parsed_kb_number = last_element_of_split_text.removeprefix("(KB").removesuffix(")")
    return str(parsed_kb_number)


def get_build_number(kb_number: str) -> dict:
    response = requests.get(f"https://support.microsoft.com/help/{kb_number}")
    soup = BeautifulSoup(response.content, "html.parser")
    header = soup.find('header',
                       attrs={'class': 'ocpArticleTitleSection', 'aria-labelledby': 'page-header', 'role': 'banner'})
    title = header.find('h1').text.strip()
    pattern = r'1904[45]\.\d+'

    # Find all matches in the text
    matches = re.findall(pattern, title)

    # Initialize the dictionary to store categorized version numbers
    version_dict = {'21H2_number': '', '22H2_number': ''}

    # Categorize each match into '21H2_number' or '22H2_number'
    for match in matches:
        if match.startswith('19044'):
            version_dict['21H2_number'] = match
        elif match.startswith('19045'):
            version_dict['22H2_number'] = match

    return version_dict


def find_anchor_id(element: Tag) -> str | None:
    """
    Finds the first anchor (<a>) tag within the given BeautifulSoup element's children
    and returns the value of its 'id' attribute.

    Parameters:
    - element: A BeautifulSoup element to search within.

    Returns:
    - The value of the 'id' attribute of the first found anchor tag, or None if no such tag or attribute is found.
    """
    # Find the first anchor tag within the given element
    anchor_tag = element.find('a')

    if anchor_tag and anchor_tag.has_attr('id'):
        # If the anchor tag is found and has an 'id' attribute, return its value
        return str(anchor_tag['id']).removesuffix('_link')
    else:
        # If no anchor tag with an 'id' attribute is found, return None
        return None


def fetch_update_html_content(update_id: str) -> str:
    """
    Fetches the HTML content for a given update ID from the Microsoft Update Catalog.

    Parameters:
    - update_id: The unique identifier for the update.

    Returns:
    - The HTML content of the response as a string.
    """
    url = 'https://www.catalog.update.microsoft.com/DownloadDialog.aspx'

    # Headers with minimal necessary information
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        "Referer": "https://www.catalog.update.microsoft.com/DownloadDialog.aspx",
        "Referrer-Policy": "strict-origin-when-cross-origin"
    }
    print(f'Fetching update {update_id}...')

    form_data = f"updateIDs=%5B%7B%22size%22%3A0%2C%22languages%22%3A%22{update_id}%22%2C%22updateID%22%3A%22{update_id}%22%7D%5D&updateIDsBlockedForImport=&wsusApiPresent=&contentImport=&sku=&serverName=&ssl=&portNumber=&version="

    # Make the POST request with URL-encoded data
    response = requests.post(url, data=form_data, headers=headers)

    # Returning the HTML content of the response
    return response.text


def get_update_final_link(html_content: str) -> str | None:
    match = re.search(r"downloadInformation\[0\].files\[0\].url = '([^']*)'", str(html_content))
    if match:
        url = match.group(1)
        print("Extracted URL:", url)
        return url
    else:
        print("URL not found.")
        return None


def create_list_of_updates(list_of_initial_urls: list[dict]) -> list[dict]:
    current_year = datetime.now().year
    current_month = datetime.now().month
    result_list = []
    for initial_url in list_of_initial_urls:
        url = create_url(initial_url['url'], p_current_year=current_year, p_current_month=current_month)
        connection_success = check_connection(url)
        if connection_success:
            final_dict = {}
            update_dict = check_element_visibility(initial_url['url'], initial_url['text'])
            element_table_update = find_element_with_text(update_dict['html_content'], initial_url['text'])
            update_id = find_anchor_id(element_table_update)
            html_content_popup_dialog = fetch_update_html_content(update_id)
            final_link = get_update_final_link(html_content_popup_dialog)
            final_dict['version'] = initial_url.get('version')
            final_dict['link'] = final_link
            final_dict['type'] = initial_url.get('type')
            final_dict['build_number'] = 'na'
            kb_number = find_kb_number(element_table_update)
            final_dict['kb_number'] = kb_number
            if initial_url['type'] == 'cumulative':
                build_number_dict = get_build_number(kb_number)
                final_dict['build_number'] = build_number_dict[f'{initial_url['version']}_number']
            result_list.append(final_dict)
    return result_list


def replace_content_between_markers(file_path, start_marker, end_marker, new_content):
    """
    Replaces content in a file between specified markers with new content.

    Parameters:
    - file_path: Path to the file to modify.
    - start_marker: The marker indicating the start of the content to replace.
    - end_marker: The marker indicating the end of the content to replace.
    - new_content: The content to insert between the markers.
    """
    with open(file_path, 'r', encoding='utf-8') as file:
        file_content = file.read()

    # Define the regular expression pattern to match content between markers
    pattern = re.compile(re.escape(start_marker) + r'.*?' + re.escape(end_marker), re.DOTALL)

    # Replace the content between markers with new content
    modified_content = re.sub(pattern, f"{start_marker}{new_content}{end_marker}", file_content)

    # Write the modified content back to the file
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(modified_content)


def create_replace_content(final_list_of_updates: list[dict]) -> str:
    string_of_objects = ''
    for update in final_list_of_updates:
        string_of_objects += f"""[pscustomobject]@{{
            Type = '{update['type']}'
            FeatureVersion = '{update['version']}'
            BuildNumber = '{update['build_number']}'
            PatchNumber = '{update['kb_number']}'
            PatchLink = {update['link']}
        }},"""
    result = f"""
$arrayOfUpdates=@(
    {string_of_objects.rstrip(string_of_objects[-1])}
)   
"""
    return result


if __name__ == '__main__':
    list_of_updates = create_list_of_updates(initial_links)
    print(json.dumps(list_of_updates, indent=4))
    content_to_replace = create_replace_content(list_of_updates)
    print(content_to_replace)
    replace_content_between_markers('./update_tool.ps1', '##REPLACE START', '##REPLACE END', content_to_replace)
