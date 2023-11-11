
import requests
from datetime import datetime
import xlwt
from xlwt import Workbook

def get_pr_list():
    # Define your repository and credentials
    from requests.auth import HTTPBasicAuth

    # Define your Bitbucket project and credentials
    project_key = 'projects/GEN2'
    username = 'Tzvika_Lifshitz'
    password = 'YkXKS8yPFzaMb9KddQ2v'
    repo_slug = 'https://bitbucket.org/softimize/gen2-plugin-lib'
    workspace_id = "https://bitbucket.org/softimize/workspace"
    project_key = "GEN2"
    # Define your date range
    start_date = datetime(2023, 8, 1).timestamp()
    end_date = datetime(2023, 12, 31).timestamp()

    repo_slugs = []
    url = "https://api.bitbucket.org/2.0/repositories/softimize?q=project.key%3D%22GEN2%22"

    while (url != None):  #handle pagination by checking of there is a next page
        # Define the API URL to get all repos
        # Send a GET request to the Bitbucket API
        response = requests.get(url, auth=HTTPBasicAuth(username, password))

        data = response.json()
        if 'next' in data:
            url = data['next']
        else:
            url = None

        for value in data['values']:
            repo_slugs.append(value['name'])

    pull_requests_in_date_range = []
    for repo in repo_slugs:
        url = f"https://api.bitbucket.org/2.0/repositories/softimize/"+repo+"/pullrequests?state=MERGED"
        while (url != None):
             response = requests.get(url, auth=(username, password))

             data = response.json()

             if 'next' in data:
                 url = data['next']
             else:
                 url = None

             # Filter the pull requests by date

             if 'values' in data:
                 paginated_pull_requests_in_date_range = [
                 pr for pr in data['values']
                 if start_date <= datetime.strptime(pr['created_on'], '%Y-%m-%dT%H:%M:%S.%f%z').timestamp() <= end_date
                 ]
                 pull_requests_in_date_range.extend(paginated_pull_requests_in_date_range)

    ob = Workbook()
    ws = ob.add_sheet('Merged PRs', True)
    generate_report(pull_requests_in_date_range, ws)

    ob.save('Merged-PR-report.xls')


def generate_report(pr_list, ws):
    date_format = xlwt.XFStyle()
    bold_style_head = xlwt.easyxf('font: bold 1, color blue; align: wrap on, vert centre, horiz center')
    bold_style = xlwt.easyxf('font: bold 1, color black')
    reg_style = xlwt.easyxf('font: color black; align: horiz center')
    date_format.num_format_str = 'yyyy/mm/dd'

    row = 0
    ws.write(row, 0, "Title", bold_style_head)
    ws.write(row, 1, "Date created", bold_style_head)
    ws.write(row, 2, "Merge Destination", bold_style_head)
    ws.write(row, 3, "Repository", bold_style_head)

    row = row + 1

    for value in pr_list:
        if len(value) != 0:
            title = value['title']
            created_on = value['created_on'].split("T")[0]
            merge_dest = value['destination']['branch']['name']
            repo = value['destination']['repository']['name']
            style = reg_style

            ws.write(row, 0, title, bold_style)
            ws.write(row, 1, created_on, style)
            ws.write(row, 2, merge_dest, style)
            ws.write(row, 3, repo, style)

            row = row + 1

if __name__ == '__main__':
    get_pr_list()
