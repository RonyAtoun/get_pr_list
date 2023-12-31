import argparse
import concurrent.futures
import sys

import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime
import time
from dotenv import load_dotenv
import os
import xlwt
from xlwt import Workbook


def get_pr_list():
    args = get_arguments()  # command line: -projectName <name of bitbucket project> for example "GEN2"
    load_dotenv('bitBucket.env')
    username = os.getenv('USERNAME1')  ## for some reason failing to load username with the key USENAME from .env
    password = os.getenv('PASSWORD')

    start = time.perf_counter()
    repo_slugs = get_repo_slugs(username, password, args.project_name)
    finish_get_repos = time.perf_counter()
    print(f'Finished get repos in {round(finish_get_repos - start, 2)} seconds')

    URLS = []
    for repo in repo_slugs:
        URLS.append(f"https://api.bitbucket.org/2.0/repositories/softimize/" + repo + "/pullrequests?state=MERGED")
    with concurrent.futures.ThreadPoolExecutor() as executor:
        repo_result = {executor.submit(get_pull_requests, url, username, password): url for url in URLS}
        results = concurrent.futures.wait(repo_result)

        pull_requests_in_date_range = []
        for future in results.done:
            repo = repo_result[future]
            pull_requests_in_date_range.append(future.result())

    finish_get_pull_requests = time.perf_counter()
    print(f'Finished get PRs in {round(finish_get_pull_requests - finish_get_repos, 2)} seconds')

    finish = time.perf_counter()
    print(f'Finished in {round(finish - start, 2)} seconds')

    ob = Workbook()
    ws = ob.add_sheet('Merged PRs', True)

    generate_report(pull_requests_in_date_range, ws)
    ob.save(args.project_name + '-Merged-PR-report.xls')


def get_repo_slugs(username, password, project_name):
    page = 1
    # Define the API URL to get all repos
    url = "https://api.bitbucket.org/2.0/repositories/softimize?q=project.key%3D%22" + project_name + "%22"
    repo_slugs = []
    while url is not None:  # handle pagination by checking of there is a next page
        response = requests.get(url, auth=HTTPBasicAuth(username, password))

        data = response.json()
        if 'next' in data:
            url = data['next']
        else:
            url = None

        for value in data['values']:
            repo_slugs.append(value['slug'])
    return repo_slugs


def get_pull_requests(url, username, password):
    # Define your date range
    start_date = datetime(2023, 8, 1).timestamp()
    end_date = datetime(2023, 12, 31).timestamp()

    pull_requests_in_date_range = []

    while url is not None:
        response = requests.get(url, auth=(username, password))
        if response is not None and "Rate limit" not in response:
            data = response.json()
        else:
            sys.exit("API Rate limit exceeded. Try again in one hour")

        if 'next' in data:
            url = data['next']
        else:
            url = None

        # Filter the pull requests by date
        if 'values' in data:
            paginated_pull_requests_in_date_range = [
                pr for pr in data['values']
                if
                start_date <= datetime.strptime(pr['created_on'], '%Y-%m-%dT%H:%M:%S.%f%z').timestamp() <= end_date
            ]
            pull_requests_in_date_range.extend(paginated_pull_requests_in_date_range)
    return pull_requests_in_date_range


def generate_report(pr_list, ws):
    date_format = xlwt.XFStyle()
    bold_style_head = xlwt.easyxf('font: bold 1, color blue; align: wrap on, vert centre, horiz center')
    bold_style = xlwt.easyxf('font: bold 1, color black')
    reg_style = xlwt.easyxf('font: color black; align: horiz center')
    date_format.num_format_str = 'yyyy/mm/dd'

    row = 0
    ws.write(row, 0, "Title", bold_style_head)
    ws.write(row, 1, "Link", bold_style_head)
    ws.write(row, 2, "Date created", bold_style_head)
    ws.write(row, 3, "Merge Destination", bold_style_head)
    ws.write(row, 4, "Repository", bold_style_head)

    row = row + 1
    for repo in pr_list:
        for value in repo:
            if len(value) != 0:
                title = value['title']
                link = value['merge_commit']['links']['html']['href']
                created_on = value['created_on'].split("T")[0]
                merge_dest = value['destination']['branch']['name']
                repo = value['destination']['repository']['name']
                style = reg_style

                ws.write(row, 0, title, bold_style)
                ws.write(row, 1, link, bold_style)
                ws.write(row, 2, created_on, style)
                ws.write(row, 3, merge_dest, style)
                ws.write(row, 4, repo, style)

                row = row + 1


def get_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument('-projectName', dest="project_name", action="store", required=True,
                        help='name of bitBucket project')

    return parser.parse_args()


if __name__ == '__main__':
    get_pr_list()
