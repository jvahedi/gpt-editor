#!pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org aspose-words

import urllib.request, json
import aspose.words as aw
import numpy as np
import glob
from datetime import datetime as dt

def config():
    config = []
    params = []
    with open('Config.txt') as f:
        for line in f:
            split = line.split()
            config.append(split)
        params.append(config[0][-1])
        params.append(float(config[1][-1]))
        params.append(float(config[2][-1]))
        params.append(int(config[3][-1]))
    return params

vals = config()
print(vals)
def gptRespond(prompt, t = 1, c = 1, GPT = 3):
    #Place personal key here in string format
    KEY = vals[0]

    try:
        url = "https://apigw.rand.org/openai/RAND/inference/deployments/gpt-35-turbo-v0301-base/chat/completions?api-version=2023-03-15-preview"

        if (GPT == 3):
            model = 'gpt-35-turbo'
            key = KEY
        elif (GPT == 4):
            model = 'gpt-4-v0314-base'
            key = KEY
        else:
            print("Model not available.")
            return 
            
        hdr ={
        # Request headers
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
        'Ocp-Apim-Subscription-Key': key,
        }

        # Request body
        data = {'model': model, 'messages': [{'role': 'user', 'content': prompt}], 'temperature': t, 'top_p' : c}
        data = json.dumps(data)
        req = urllib.request.Request(url, headers=hdr, data = bytes(data.encode("utf-8")))

        req.get_method = lambda: 'POST'
        response = urllib.request.urlopen(req)

        content = bytes.decode(response.read(), 'utf-8') #return string value
        res = json.loads(content)
        return res['choices'][0]['message']['content']
    except Exception as e:
        print(e)

def gptMergeDoc(temp = vals[1], top_c = vals[2], GPT = vals[3]):
    ##License section
    # editWordLicense = aw.License()
    # editWordLicense.set_license("Aspose.Word.lic")
    
    #Finding all files with ".docx" as an extension
    for file in glob.glob('*.docx'):
        print(file)
        ind = file.find('.doc')
        #Doc Name
        inname = file[:ind]
        #extension Type
        ext = file[ind:]
        #Importing First Word Document
        docToRead = aw.Document(file)

        #Extract text from Word Document
        text = np.array([],dtype = object)
        for paragraph in docToRead.get_child_nodes(aw.NodeType.PARAGRAPH, True) :    
            paragraph = paragraph.as_paragraph()
            text = np.append(text, paragraph.to_string(aw.SaveFormat.TEXT))
        text = text[1:-2]

        body = ''
        for t in text:
            body = body + t

        # The target document doc2.
        doc2 = aw.Document()
        builder = aw.DocumentBuilder(doc2)
        request = 'Can you edit the following for grammatical errors?'
        answer = gptRespond(request + '\n' + body, temp, top_c, GPT)
        builder.writeln(answer)

        # DocToRead now contains changes as revisions.
        docToRead.compare(doc2, "GPT", dt.today())
        # save merged document in local directory
        docToRead.save(inname + '_edited'+ ext)
    
gptMergeDoc()
