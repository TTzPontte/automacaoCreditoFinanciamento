from urllib import request

def makeDownload(url, pathAndName):
    response = request.urlretrieve(url, pathAndName)
    print(response)

