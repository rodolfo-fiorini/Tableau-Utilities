# Python script using TSC to publish PPT workbook to server in workflow.

import argparse
import tableauserverclient as TSC
from tableauserverclient import ConnectionCredentials, ConnectionItem

parser = argparse.ArgumentParser(description='Publish a workbook to server.')

parser.add_argument('--server', '-s', required=True, help='server address') # tableau.apogeeintegration.com
parser.add_argument('--tokenname', '-t', required=True, help='Token name') # test_token
parser.add_argument('--workbookpath', '-w', required = True, help = 'Tableau file to publish to server') # ppt_imported_story.twbx

args = parser.parse_args()

token_value = input("Please enter token value:") #'G2YYOZ/bTgmJWUoAnmnQlQ==:ydeQWjbpdhEYsGppUPNYL0HpYS7L7q3F'
tableau_auth = TSC.PersonalAccessTokenAuth(args.tokenname, token_value, '')
overwrite_true = TSC.Server.PublishMode.Overwrite

server = TSC.Server(args.server, use_server_version = True)


with server.auth.sign_in(tableau_auth):

        # Step 2: Get all the projects on server, then look for the default one.
        all_projects, pagination_item = server.projects.get()
        default_project = next((project for project in all_projects if project.is_default()), None)

        connection1 = ConnectionItem()
        connection1.server_address = "mssql.test.com"
        connection1.connection_credentials = ConnectionCredentials("test", "password", True)

        connection2 = ConnectionItem()
        connection2.server_address = "postgres.test.com"
        connection2.server_port = "5432"
        connection2.connection_credentials = ConnectionCredentials("test", "password", True)

        all_connections = list()
        all_connections.append(connection1)
        all_connections.append(connection2)

        # Step 3: If default project is found, form a new workbook item and publish.
        if default_project is not None:
            new_workbook = TSC.WorkbookItem(default_project.id)

            new_job = server.workbooks.publish(new_workbook, args.workbookpath, overwrite_true,
                                                           connections=all_connections, as_job= True,
                                                           skip_connection_check=True)
            print("Workbook published. JOB ID: {0}".format(new_job.id))

        else:
            error = "The default project could not be found."
            raise LookupError(error)
