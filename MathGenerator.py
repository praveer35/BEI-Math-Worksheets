from __future__ import print_function
import os.path
import os
import random
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/documents']

# The ID of a sample document.
#DOCUMENT_ID = '1wfSsP1HtsUQKPtrko1mu_G1Cdoax2U9QR5KoSCrT8Vw'

INSTRUCTIONS = {
    'Aster Place': {
        'boxes': 0,
        'qInBox': 0,
        'evenDistribution': 'n',
        'weight': [1, 1, 1, 1]
    },
    'Creasy Springs': {
        'boxes': 6,
        'qInBox': 15,
        'evenDistribution': 'n',
        'weight': [1, 1, 1, 1]
    },
    'Cumberland': {
        'boxes': 1,
        'qInBox': 96,
        'evenDistribution': 'y'
    },
    'Digby Place': {
        'boxes': 1,
        'qInBox': 96,
        'evenDistribution': 'y'
    },
    'Five Star': {
        'boxes': 1,
        'qInBox': 96,
        'evenDistribution': 'n',
        'weight': [1, 1, 0.75, 1]
    },
    'Glasswater': {
        'boxes': 1,
        'qInBox': 96,
        'evenDistribution': 'n',
        'weight': [1, 1, 0.75, 1]
    },
    'Mulberry Health': {
        'boxes': 5,
        'qInBox': 15,
        'evenDistribution': 'n',
        'weight': [1, 1, 1, 1]
    },
    'Rosewalk': {
        'boxes': 4,
        'qInBox': 15,
        'evenDistribution': 'n',
        'weight': [1, 1, 0.75, 1]
    },
    'St. Mary': {
        'boxes': 5,
        'qInBox': 15,
        'evenDistribution': 'n',
        'weight': [1, 1, 0.75, 1]
    },
    'The Springs': {
        'boxes': 1,
        'qInBox': 75,
        'evenDistribution': 'n',
        'weight': [1, 1, 0.75, 1]
    },
    'Westminster': {
        'boxes': 1,
        'qInBox': 102,
        'evenDistribution': 'n',
        'weight': [1, 1, 1.25, 1]
    },
    'Wickshire': {
        'boxes': 1,
        'qInBox': 102,
        'evenDistribution': 'n',
        'weight': [1, 1, 0.75, 1]
    }
}

def calc_weight(questionsLeft):
    sum = 0
    cumulativeSum = 0
    weight = [0] * len(questionsLeft)
    for i in range(len(questionsLeft)):
        sum += questionsLeft[i]
    for i in range(len(questionsLeft)):
        weight[i] = questionsLeft[i] / sum + cumulativeSum
        cumulativeSum = weight[i]
    return weight

def main():
    print("URL: ", end="")
    URL = input()
    index1 = URL.index('/d/') + 3
    index2 = URL.index('/edit')
    DOCUMENT_ID = URL[index1:index2]

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    try:
        service = build('docs', 'v1', credentials=creds)
        document = service.documents().get(documentId=DOCUMENT_ID).execute()
        doc_content = document.get('body').get('content')
        print('The title of the document is: {}'.format(document.get('title')))
        requests = []
        i = len(doc_content) - 1
        startIndex = endIndex = 0
        weight = [0.25, 0.25, 0.25, 0.25]
        report = [0] * 4
        boxes = qInBox = 0
        evenDistribution = ''
        for residential_home in INSTRUCTIONS:
            if residential_home in document.get('title'):
                boxes = INSTRUCTIONS[residential_home]['boxes']
                qInBox = INSTRUCTIONS[residential_home]['qInBox']
                evenDistribution = INSTRUCTIONS[residential_home]['evenDistribution']
                if (evenDistribution.upper() == 'N'):
                    weight = INSTRUCTIONS[residential_home]['weight']
                break
        weight = calc_weight(weight)
        while i >= 0:
            if ('paragraph' in doc_content[i] and 'elements' in doc_content[i]['paragraph'] and len(doc_content[i]['paragraph']['elements']) > 0 and 'textRun' in doc_content[i]['paragraph']['elements'][0]):
                if ("Start time:   ___:___" in doc_content[i]['paragraph']['elements'][0]['textRun']['content']):
                    temp = doc_content[i]['paragraph']['elements'][0]['startIndex'] + doc_content[i]['paragraph']['elements'][0]['textRun']['content'].index('Start time:   ___:___')
                    startIndex = temp + len('Start time:   ___:___')
                    requests.append({'deleteContentRange': {
                        'range': {
                            'startIndex': startIndex,
                            'endIndex': endIndex,
                        }
                    }})
                    requests.append({'insertText': {
                        'text': '\n',
                        'location': {
                            'index': startIndex
                        }
                    }})
                    startIndex += 1
                    service.documents().batchUpdate(
                    documentId=DOCUMENT_ID, body={'requests': requests}).execute()
                    requests.clear()
                if ("End time:   ___:___" in doc_content[i]['paragraph']['elements'][0]['textRun']['content']):
                    endIndex = doc_content[i]['paragraph']['elements'][0]['startIndex'] + doc_content[i]['paragraph']['elements'][0]['textRun']['content'].index('End time:   ___:___')
            i -= 1
        """weight = [0.25, 0.25, 0.25, 0.25]
        report = [0] * 4
        print("Boxes (1 if none): ", end="")
        boxes = int(input())
        print("Questions per box: ", end="")
        qInBox = int(input())
        print("They have to be equally distributed, Yes(y) or No(n): ", end="")
        evenDistribution = input()
        if evenDistribution.upper() != "Y":
            print("Enter the ratio, + - x ÷ (i.e. 1 1 1 1 or 1 1 1.5 1): ", end="")
            weightStr = input()
            weight = weightStr.split(" ")
            for i in range(len(weight)):
                weight[i] = float(weight[i])
        weight = calc_weight(weight)"""
        questionsLeft = [qInBox*boxes/4] * 4
        problems = []
        operator = 0
        num1 = 0
        num2 = 0
        operator2 = 0
        for i in range(boxes):
            mathStr = ""
            j = 0
            newLineBuffer = 0
            while j < qInBox:
                tabException = False
                if evenDistribution.upper() == "Y":
                    weight = calc_weight(questionsLeft)
                operator = random.uniform(0, 1)
                if operator <= weight[0]:
                    num1 = random.randint(1, 20)
                    num2 = random.randint(1, num1)
                    num1 -= num2
                    out = str(num2) + " + " + str(num1)
                    operator2 = 0
                elif operator <= weight[1]:
                    num1 = random.randint(0, 20)
                    num2 = random.randint(0, num1)
                    out = str(num1) + " - " + str(num2)
                    operator2 = 1
                    if (num1 < 10 and num2 == 1) or (num1 == 1 and num2 < 10):
                        tabException = True
                elif operator <= weight[2]:
                    num1 = random.randint(1, 10)
                    num2 = random.randint(1, 10)
                    out = str(num1) + " x " + str(num2)
                    operator2 = 2
                else:
                    num1 = random.randint(1, 10)
                    num2 = random.randint(1, 10)
                    num1 *= num2
                    out = str(num1) + " ÷ " + str(num2)
                    operator2 = 3
                if out in problems:
                    continue
                else:
                    questionsLeft[operator2] -= 1
                    report[operator2] += 1
                    problems.append(out)
                    if j % 3 == 2 and j > 0:
                        mathStr += out + " = "
                        #if boxes != 1:
                        #    mathStr += "\n
                        if j != qInBox - 1:
                            mathStr += "\n"
                    else:
                        mathStr += out + " =   \t\t\t"
                        if tabException:
                            mathStr += "\t"
                j += 1
            if boxes == 1:
                requests.append({
                    'insertText': {
                        'text': '*' + "\n",
                        'location': {
                            'index': startIndex
                        }
                    }
                })
                requests.append({
                    'updateParagraphStyle': {
                        'range': {
                            'startIndex': startIndex,
                            'endIndex': startIndex + 1
                        },
                        'fields': 'lineSpacing',
                        'paragraphStyle': {
                            'lineSpacing': 150
                        }
                    }
                })
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': startIndex,
                            'endIndex': startIndex + 1
                        },
                        'fields': 'fontSize',
                        'textStyle': {
                            'fontSize': {
                                'magnitude': 22,
                                'unit': 'PT'
                            }
                        }
                    }
                })
                requests.append({
                    'replaceAllText': {
                        'containsText': {
                            'text': '*',
                            'matchCase':  'true'
                        },
                        'replaceText': mathStr
                    }
                })
            else:
                if i == boxes - 2:
                    requests.append({
                        'insertText': {
                            'text': '\n\n',
                            'location': {
                                'index': startIndex
                            }
                        }
                    })
                    newLineBuffer = 2
                if i == 0 and boxes % 2 == 0:
                    requests.append({
                        'insertText': {
                            'text': '\n',
                            'location': {
                                'index': startIndex
                            }
                        }
                    })
                requests.append({
                    'insertTable': {
                        'rows': 1,
                        'columns': 1,
                        'location': {
                            'index': startIndex + newLineBuffer
                        }
                    }
                })
                requests.append({
                    'insertText': {
                        'text': '*',
                        'location': {
                            'index': startIndex + newLineBuffer + 4
                        }
                    }
                })
                requests.append({
                    'updateParagraphStyle': {
                        'range': {
                            'startIndex': startIndex + newLineBuffer + 4,
                            'endIndex': startIndex + newLineBuffer + 5
                        },
                        'fields': 'lineSpacing',
                        'paragraphStyle': {
                            'lineSpacing': 197
                        }
                    }
                })
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': startIndex,
                            'endIndex': startIndex + 1
                        },
                        'fields': 'fontSize',
                        'textStyle': {
                            'fontSize': {
                                'magnitude': 22,
                                'unit': 'PT'
                            }
                        }
                    }
                })
                requests.append({
                    'replaceAllText': {
                        'containsText': {
                            'text': '*',
                            'matchCase':  'true'
                        },
                        'replaceText': mathStr
                    }
                })
            service.documents().batchUpdate(
            documentId=DOCUMENT_ID, body={'requests': requests}).execute()
            requests.clear()
        print("Report" + str(report))
    except HttpError:
        print('Error in completion.')


if __name__ == '__main__':
    main()