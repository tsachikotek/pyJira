from jira import JIRA
import re
import os, sys, getopt
import Word

def main(argv):
    jira_url = ''
    jira_password = ''
    jira_username = ''
    jira_filter = ''

    try:
        opts, args = getopt.getopt(argv, "l:u:p:f:", ["url=", "username=", "password=", "filter="])
        print('ARGS:')

        for opt, arg in opts:
            print(opt, arg)

    except getopt.GetoptError:
        print('ERROR: Jira.py -l <jira url> -u <username> -p <password> -f <filter name>')
        sys.exit(2)

        print('PARSING ARGS:')
    for opt, arg in opts:
        if opt == '-h':
            print('Jira.py -l <jira url> -u <username> -p <password> -f <filter name>')
            sys.exit()
        else:
            print(opt, arg)
            if opt in ("-l", "--url"):
                jira_url = arg
            if opt in ("-u", "--username"):
                jira_username = arg
            if opt in ("-p", "--password"):
                jira_password = arg
            if opt in ("-f", "--filter"):
                jira_filter = arg

    issues = export_jira_issues_by_filter(jira_url, jira_username, jira_password, jira_filter)

    word_docuemnt_filename = 'export2.docx'
    if not(os.path.isfile(word_docuemnt_filename)):
        Word.new_document(word_docuemnt_filename)

    for issue in issues:
        print('Current issue:', issue.key)
        jira_issue = issue.key + ': ' + issue.fields.summary
        jira_body = issue.fields.description
        Word.insert_Issue(jira_issue, jira_body, issue.fields.created, issue.fields.status.name, word_docuemnt_filename)

    Word.updateTOC (word_docuemnt_filename)

def export_jira_issues_by_filter (jira_url, jira_username, jira_password, jira_filter):
    print('export_jira_issues_by_filter:')
    # jira = JIRA()
    auth_jira = JIRA(auth=(jira_username, jira_password), options={'server': jira_url})
    #projects = auth_jira.projects()

    block_size = 100
    block_num = 0
    start_idx = 0
    jql = "filter=" + jira_filter
    print('CONNECTING TO JIRA')
    issues = auth_jira.search_issues(jql, start_idx, block_size)
    print('NUMBER OF ISSUES RETRIEVED: ', issues.__len__())

    return issues

if __name__ == "__main__":
   main(sys.argv[1:])



