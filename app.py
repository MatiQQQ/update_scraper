import requests
from datetime import datetime
from bs4 import BeautifulSoup
import os
import copy
import json

currentMonth = datetime.now().month
currentYear = datetime.now().year

path_to_script = "./update_tool.ps1"
index_replace_start = 0
index_replace_end = 0
lines_to_safe = []
string_of_objects = ""
updates_results = []

update_catalog_links_request = [
    {
        "type": "net",
        "windows_feature_version": "22H2",
        "link": f"https://www.catalog.update.microsoft.com/Search.aspx?q={currentYear}-{currentMonth:02d}%20Update%20for%20.NET%20Framework%203.5%20and%204.8%20for%20Windows%2010%20Version%2022H2%20for%20x64",
    },
    {
        "type": "net",
        "windows_feature_version": "21H2",
        "link": f"https://www.catalog.update.microsoft.com/Search.aspx?q={currentYear}-{currentMonth:02d}%20Update%20for%20.NET%20Framework%203.5%20and%204.8%20for%20Windows%2010%20Version%2021H2%20for%20x64",
    },
    {
        "type": "cumulative",
        "windows_feature_version": "22H2",
        "link": f"https://www.catalog.update.microsoft.com/Search.aspx?q={currentYear}-{currentMonth:02d}%20Cumulative%20Update%20for%20Windows%2010%20Version%2022H2%20for%20x64-based%20Systems",
    },
    {
        "type": "cumulative",
        "windows_feature_version": "21H2",
        "link": f"https://www.catalog.update.microsoft.com/Search.aspx?q={currentYear}-{currentMonth:02d}%20Cumulative%20Update%20for%20Windows%2010%20Version%2021H2%20for%20x64-based%20Systems",
    },
]


def get_link(id):
    result = ""
    request = requests.post(
        "https://www.catalog.update.microsoft.com/DownloadDialog.aspx",
        data=f"updateIDs=%5B%7B%22size%22%3A0%2C%22languages%22%3A%22{id}%22%2C%22updateID%22%3A%22{id}%22%7D%5D&updateIDsBlockedForImport=&wsusApiPresent=&contentImport=&sku=&serverName=&ssl=&portNumber=&version=",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    with open(f"./update_{id}.html", "w") as html_file_input:
        html_file_input.write(request.text)
        with open(f"./update_{id}.html", "r") as html_file_output:
            lines = html_file_output.readlines()
            for line in lines:
                if "downloadInformation[0].files[0].url" in line:
                    result = line.split("=")[1].replace(" ", "")
        os.remove(f"./update_{id}.html")
    return result


def format_id(string):
    return str(string).split("_")[0]


def format_patch_update_number(string):
    return str(string).split("(")[1].replace(")", "").strip()


def get_build_number(kb_number, feature_version):
    formatted_kb_number = str(kb_number).split("KB")[1]
    html = requests.get(
        f"https://support.microsoft.com/en-us/help/{formatted_kb_number}"
    )
    soup = BeautifulSoup(html.text, features="html.parser")
    title = soup.find("h1", attrs={"id": "page-header"})
    if feature_version == "22H2":
        return (str(title.text).split(" ")[7]).removesuffix(")")
    elif feature_version == "21H2":
        return str(title.text).split(" ")[5]


def get_link_and_patch_number(link):
    link_result = ""
    patch_number_result = ""
    patch_build_number = "na"
    html = requests.get(link["link"])
    soup = BeautifulSoup(html.text, features="html.parser")
    links = soup.find_all("a", attrs={"class": "contentTextItemSpacerNoBreakLink"})
    for link_item in links:
        if ("Dynamic" not in link_item.string and link["type"] == "cumulative") or (
            link["type"] == "net" and ("4.8.1" not in link_item.string)
        ):
            update_id = format_id(link_item["id"])
            link_result = get_link(update_id)
            patch_number_result = format_patch_update_number(link_item.string)
            if link["type"] == "cumulative":
                patch_build_number = get_build_number(
                    patch_number_result, link["windows_feature_version"]
                )
            break
    return {
        "Type": link["type"],
        "FeatureVersion": link["windows_feature_version"],
        "BuildNumber": patch_build_number,
        "PatchNumber": patch_number_result,
        "PatchLink": str(link_result).strip().replace(";", ""),
    }


for i, request_link in enumerate(update_catalog_links_request):
    link_dict = get_link_and_patch_number(request_link)
    if link_dict["PatchLink"] == "":
        currentMonth_copied = copy.deepcopy(currentMonth)
        if currentMonth_copied == 1:
            currentMonth = 12
            currentYear = currentYear - 1
            index = 1
            while link_dict["PatchLink"] == "":
                currentMonth -= index
                if request_link["type"] == "cumulative":
                    request_dict = {
                        "type": "cumulative",
                        "windows_feature_version": request_link[
                            "windows_feature_version"
                        ],
                        "link": f"https://www.catalog.update.microsoft.com/Search.aspx?q={currentYear}-{currentMonth:02d}%20Cumulative%20Update%20for%20Windows%2010%20Version%20{request_link['windows_feature_version']}%20for%20x64-based%20Systems",
                    }
                else:
                    request_dict = {
                        "type": "net",
                        "windows_feature_version": request_link[
                            "windows_feature_version"
                        ],
                        "link": f"https://www.catalog.update.microsoft.com/Search.aspx?q={currentYear}-{currentMonth:02d}%20Update%20for%20.NET%20Framework%203.5%20and%204.8%20for%20Windows%2010%20Version%20{request_link['windows_feature_version']}%20for%20x64",
                    }
                link_dict = get_link_and_patch_number(request_dict)
                if link_dict["PatchLink"] == "":
                    continue
                updates_results.append(link_dict)
                index += 1

        else:
            index = 1
            while link_dict["PatchLink"] == "" or currentMonth - index == 1:
                currentMonth = datetime.now().month - index
                if request_link["type"] == "cumulative":
                    request_dict = {
                        "type": "cumulative",
                        "windows_feature_version": request_link[
                            "windows_feature_version"
                        ],
                        "link": f"https://www.catalog.update.microsoft.com/Search.aspx?q={currentYear}-{currentMonth:02d}%20Cumulative%20Update%20for%20Windows%2010%20Version%20{request_link['windows_feature_version']}%20for%20x64-based%20Systems",
                    }
                else:
                    request_dict = {
                        "type": "net",
                        "windows_feature_version": request_link[
                            "windows_feature_version"
                        ],
                        "link": f"https://www.catalog.update.microsoft.com/Search.aspx?q={currentYear}-{currentMonth:02d}%20Update%20for%20.NET%20Framework%203.5%20and%204.8%20for%20Windows%2010%20Version%20{request_link['windows_feature_version']}%20for%20x64",
                    }
                link_dict = get_link_and_patch_number(request_dict)
                if link_dict["PatchLink"] == "":
                    continue
                updates_results.append(link_dict)
                index += 1
    else:
        updates_results.append(link_dict)
print(json.dumps(updates_results, indent=2))

with open(f"{path_to_script}", "r") as file:
    lines = file.readlines()
    for index, line in enumerate(lines):
        if line.startswith("## REPLACE START"):
            index_replace_start = index
        if line.startswith("## REPLACE END"):
            index_replace_end = index
            break
    for index, line in enumerate(lines):
        if index in range(index_replace_start + 1, index_replace_end):
            continue
        lines_to_safe.append(line)


for update in updates_results:
    string_of_objects += f"""[pscustomobject]@{{
        Type = '{update['Type']}'
        FeatureVersion = '{update['FeatureVersion']}'
        BuildNumber = '{update['BuildNumber']}'
        PatchNumber = '{update['PatchNumber']}'
        PatchLink = {update['PatchLink']}
    }},"""

lines_to_safe[
    index_replace_start + 1
] = f"""$arrayOfUpdates=@(
    {string_of_objects.rstrip(string_of_objects[-1])}
)
## REPLACE END
"""

with open(f"{path_to_script}", "w") as file_write:
    for index, line in enumerate(lines_to_safe):
        if line.startswith("## REPLACE_DATE"):
            lines_to_safe[
                index
            ] = f"## REPLACE_DATE {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            break
    file_write.writelines(lines_to_safe)
