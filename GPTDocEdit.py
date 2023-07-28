!pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org aspose-words

import urllib.request, json
import aspose.words as aw
import numpy as np


#Queries ChatGPT with the prompt (includes request and main body)
#Optional parameters of Temprature, Top_percent, and GPT Class 3/4
def gptRespond(prompt, t = 1, c = 1, GPT = 3):
    try:
        url = "https://apigw.rand.org/openai/RAND/inference/deployments/gpt-35-turbo-v0301-base/chat/completions?api-version=2023-03-15-preview"
        
        #Place personal key here in string format
        #KEY = 'INSERT_KEY_HERE'
        
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

        

def gptMergeDoc(inname = 'test',outname = 'merged', temp =2, top_c = .7, GPT =3):
    ##License section
    # editWordLicense = aw.License()
    # editWordLicense.set_license("Aspose.Word.lic")
    
    #Importing Word Document
    docToRead = aw.Document(inname + ".docx")
    
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

    # Doc1 now contains changes as revisions.
    docToRead.compare(doc2, "GPT", dt.today())
    # save document in local directory
    name = 'merged'
    docToRead.save(name + ".docx")
    
gptMergeDoc()
