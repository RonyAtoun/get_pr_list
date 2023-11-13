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


def single_repo_pr_list():
    args = get_arguments()  # command line: -repoName <name of bitbucket project>
    load_dotenv('bitBucket.env')
    # username = os.getenv('USERNAME')  ## for some reason failing to load username from .env
    username = 'Tzvika_Lifshitz'
    password = os.getenv('PASSWORD')

    url = f"https://api.bitbucket.org/2.0/repositories/softimize/" + args.repo_name + "/pullrequests?state=MERGED"
    pull_requests_in_date_range = get_pull_requests(url, username, password)

    ob = Workbook()
    ws = ob.add_sheet('Merged PRs', True)

    generate_report(pull_requests_in_date_range, ws)
    ob.save(args.repo_name + '-Merged-PR-report.xls')


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

    for value in pr_list:
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
    parser.add_argument('-repoName', dest="repo_name", action="store", required=True,
                        help='name of bitBucket project')

    return parser.parse_args()


if __name__ == '__main__':
    single_repo_pr_list()
